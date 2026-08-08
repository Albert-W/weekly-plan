#!/usr/bin/env python3
"""Enrich local diary entries with environment + LLM fields.

Runs after a diary's day has ended. For each row where `enriched_at` is NULL:
  - weekday / ISO week / season            (stdlib)
  - Chinese lunar date                     (classic LUNAR_INFO table, 1900-2100)
  - moon phase                             (classic synodic-month formula)
  - weather for that date                  (Open-Meteo HISTORICAL API, no key)
  - tags / worry_controllable / summary    (local liteLLM proxy, OpenAI-compatible)

Pure Python stdlib (urllib + sqlite3). The LLM call goes to your local
liteLLM proxy's OpenAI-compatible /v1/chat/completions endpoint — set
LITELLM_URL / LITELLM_MODEL / LITELLM_API_KEY in the config below or via env.

Run on a schedule (e.g. daily after midnight):
    python3 enrich_diary.py
"""

import datetime
import json
import os
import sqlite3
import sys
import urllib.error
import urllib.parse
import urllib.request

try:
    import dotenv  # local/dotenv.py — tiny .env loader
except Exception:  # pragma: no cover - import path differs in some runners
    dotenv = None

# ---- Config (edit me in local/.env, NOT in this file) ----
HERE = os.path.dirname(os.path.abspath(__file__))
if dotenv:
    dotenv.load_dotenv(os.path.join(HERE, ".env"))

DB_PATH = os.path.join(HERE, os.environ.get("DB_PATH", "bot.db"))
LAT = float(os.environ.get("LAT", "53.3498"))       # Dublin; set to your city
LON = float(os.environ.get("LON", "-6.2603"))
WEATHER_TZ = os.environ.get("WEATHER_TZ", "Europe/Dublin")  # raw — urlencode encodes it
# Local liteLLM proxy (OpenAI-compatible /v1/chat/completions).
LITELLM_URL = os.environ.get("LITELLM_URL", "http://localhost:4000/v1/chat/completions")
LITELLM_MODEL = os.environ.get("LITELLM_MODEL", "default")  # model your proxy exposes
LITELLM_API_KEY = os.environ.get("LITELLM_API_KEY", "")    # proxy master key (in .env)

SCHEMA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "schema.sql")

# ---- Chinese calendar constants (lunar year 1900..2100) ----
# Low 4 bits = leap month (0 = none); bit 0x10000 = leap month has 30 days;
# bits 0x8000..0x10 = whether month 1..12 has 30 days (else 29).
LUNAR_INFO = [
    0x04bd8, 0x04ae0, 0x0a570, 0x054d5, 0x0d260, 0x0d950, 0x16554, 0x056a0, 0x09ad0, 0x055d2,
    0x04ae0, 0x0a5b6, 0x0a4d0, 0x0d250, 0x1d255, 0x0b540, 0x0d6a0, 0x0ada2, 0x095b0, 0x14977,
    0x04970, 0x0a4b0, 0x0b4b5, 0x06a50, 0x06d40, 0x1ab54, 0x02b60, 0x09570, 0x052f2, 0x04970,
    0x06566, 0x0d4a0, 0x0ea50, 0x06e95, 0x05ad0, 0x02b60, 0x186e3, 0x092e0, 0x1c8d7, 0x0c950,
    0x0d4a0, 0x1d8a6, 0x0b550, 0x056a0, 0x1a5b4, 0x025d0, 0x092d0, 0x0d2b2, 0x0a950, 0x0b557,
    0x06ca0, 0x0b550, 0x15355, 0x04da0, 0x0a5b0, 0x14573, 0x052b0, 0x0a9a8, 0x0e950, 0x06aa0,
    0x0aea6, 0x0ab50, 0x04b60, 0x0aae4, 0x0a570, 0x05260, 0x0f263, 0x0d950, 0x05b57, 0x056a0,
    0x096d0, 0x04dd5, 0x04ad0, 0x0a4d0, 0x0d4d4, 0x0d250, 0x0d558, 0x0b540, 0x0b6a0, 0x195a6,
    0x095b0, 0x049b0, 0x0a974, 0x0a4b0, 0x0b27a, 0x06a50, 0x06d40, 0x0af46, 0x0ab60, 0x09570,
    0x04af5, 0x04970, 0x064b0, 0x074a3, 0x0ea50, 0x06b58, 0x055c0, 0x0ab60, 0x096d5, 0x092e0,
    0x0c960, 0x0d954, 0x0d4a0, 0x0da50, 0x07552, 0x056a0, 0x0abb7, 0x025d0, 0x092d0, 0x0cab5,
    0x0a950, 0x0b4a0, 0x0baa4, 0x0ad50, 0x055d9, 0x04ba0, 0x0a5b0, 0x15176, 0x052b0, 0x0a930,
    0x07954, 0x06aa0, 0x0ad50, 0x05b52, 0x04b60, 0x0a6e6, 0x0a4e0, 0x0d260, 0x0ea65, 0x0d530,
    0x05aa0, 0x076a3, 0x096d0, 0x04afb, 0x04ad0, 0x0a4d0, 0x1d0b6, 0x0d250, 0x0d520, 0x0dd45,
    0x0b5a0, 0x056d0, 0x055b2, 0x049b0, 0x0a577, 0x0a4b0, 0x0aa50, 0x1b255, 0x06d20, 0x0ada0,
    0x14b63, 0x09370, 0x049f8, 0x04970, 0x064b0, 0x168a6, 0x0ea50, 0x06b20, 0x1a6c4, 0x0aae0,
    0x092e0, 0x0d2e3, 0x0c960, 0x0d557, 0x0d4a0, 0x0da50, 0x05d55, 0x056a0, 0x0a6d0, 0x055d4,
    0x052d0, 0x0a9b8, 0x0a950, 0x0b4a0, 0x0b6a6, 0x0ad50, 0x055a0, 0x0aba4, 0x0a5b0, 0x052b0,
    0x0b273, 0x06930, 0x07337, 0x06aa0, 0x0ad50, 0x14b55, 0x04b60, 0x0a570, 0x054e4, 0x0d160,
    0x0e968, 0x0d520, 0x0daa0, 0x16aa6, 0x056d0, 0x04ae0, 0x0a9d4, 0x0a2d0, 0x0d150, 0x0f252,
    0x0d520,
]
LUNAR_BASE = datetime.date(1900, 1, 31)  # lunar 1900-01-01
LUNAR_MONTHS_CN = ["", "正月", "二月", "三月", "四月", "五月", "六月",
                   "七月", "八月", "九月", "十月", "冬月", "腊月"]
LUNAR_DAYS_CN = ["", "初一", "初二", "初三", "初四", "初五", "初六", "初七", "初八", "初九", "初十",
                 "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
                 "廿一", "廿二", "廿三", "廿四", "廿五", "廿六", "廿七", "廿八", "廿九", "三十"]
WEEKDAYS_CN = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
SEASONS_CN = ["春", "春", "春", "夏", "夏", "夏", "秋", "秋", "秋", "冬", "冬", "冬"]
MOON_NAMES = ["新月", "娥眉月", "上弦月", "盈凸月", "满月", "亏凸月", "下弦月", "残月"]

# WMO weather codes -> Chinese condition (daytime).
WMO_CN = {
    0: "晴", 1: "晴间多云", 2: "多云", 3: "阴", 45: "雾", 48: "雾凇",
    51: "毛毛雨", 53: "毛毛雨", 55: "毛毛雨", 61: "小雨", 63: "中雨", 65: "大雨",
    71: "小雪", 73: "中雪", 75: "大雪", 80: "阵雨", 81: "阵雨", 82: "强阵雨",
    95: "雷阵雨", 96: "雷阵雨伴冰雹", 99: "雷阵雨伴冰雹",
}


# ---- Lunar calendar (classic port of the popular JS implementation) ----

def _leap_month(y):
    return LUNAR_INFO[y - 1900] & 0xF


def _leap_days(y):
    return 30 if (LUNAR_INFO[y - 1900] & 0x10000) else 29 if _leap_month(y) else 0


def _month_days(y, m):
    return 30 if (LUNAR_INFO[y - 1900] & (0x10000 >> m)) else 29


def _year_days(y):
    info = LUNAR_INFO[y - 1900]
    total = 348  # 12 months * 29 days
    bit = 0x8000
    while bit > 0x8:
        if info & bit:
            total += 1
        bit >>= 1
    return total + _leap_days(y)


def solar_to_lunar(d):
    """Solar date -> (lunar_year, lunar_month, lunar_day, is_leap_month)."""
    offset = (d - LUNAR_BASE).days
    y = 1900
    temp = 0
    while y < 2101 and offset > 0:
        temp = _year_days(y)
        offset -= temp
        y += 1
    if offset < 0:
        offset += temp
        y -= 1
    leap = _leap_month(y)
    is_leap = False
    m = 1
    while m < 13 and offset > 0:
        if leap > 0 and m == (leap + 1) and not is_leap:
            m -= 1
            is_leap = True
            temp = _leap_days(y)
        else:
            temp = _month_days(y, m)
        if is_leap and m == (leap + 1):
            is_leap = False
        offset -= temp
        m += 1
    if offset == 0 and leap > 0 and m == leap + 1:
        if is_leap:
            is_leap = False
        else:
            is_leap = True
            m -= 1
    if offset < 0:
        offset += temp
        m -= 1
    return y, m, offset + 1, is_leap


def lunar_date_str(d):
    y, m, day, is_leap = solar_to_lunar(d)
    prefix = "闰" if is_leap else ""
    return "%s%s%s" % (prefix, LUNAR_MONTHS_CN[m], LUNAR_DAYS_CN[day])


def moon_phase_str(d):
    # Approximate: new moon epoch 2000-01-06 18:14 UTC; synodic month 29.530588853 d.
    ref = datetime.datetime(2000, 1, 6, 18, 14)
    synodic = 29.530588853
    age = ((datetime.datetime(d.year, d.month, d.day) - ref).total_seconds() / 86400.0) % synodic
    return MOON_NAMES[int(age / (synodic / 8)) % 8]


# ---- Weather (Open-Meteo, no API key) ----
# Today/future come from the LIVE FORECAST feed (api.open-meteo.com) — the
# same source phones/weather apps use, which always has the current day.
# Past dates come from the ARCHIVE feed (archive-api.open-meteo.com). The
# archive can 400 for a very recent day that isn't archived yet, in which
# case we fall back to the forecast feed's `past_days` window.

def _parse_daily(data, idx=0):
    """Human summary + raw JSON from a /daily payload, picking index idx."""
    daily = data.get("daily", {})

    def at(key):
        v = daily.get(key)
        return v[idx] if isinstance(v, list) and idx < len(v) else v

    wc = at("weathercode")
    tmax = at("temperature_2m_max")
    tmin = at("temperature_2m_min")
    cond = WMO_CN.get(wc, "未知")
    if tmax is None:
        return cond, None
    if tmin is not None:
        return "%s %s°C/%s°C" % (cond, round(tmin), round(tmax)), json.dumps(daily, ensure_ascii=False)
    return "%s %s°C" % (cond, round(tmax)), json.dumps(daily, ensure_ascii=False)


def weather_for(date_str):
    today = datetime.date.today().isoformat()
    base = {
        "latitude": LAT, "longitude": LON,
        "daily": "temperature_2m_max,temperature_2m_min,weathercode",
        "timezone": WEATHER_TZ,
    }

    if date_str >= today:
        qs = urllib.parse.urlencode(dict(base, forecast_days=1))
        with urllib.request.urlopen("https://api.open-meteo.com/v1/forecast?" + qs, timeout=30) as resp:
            return _parse_daily(json.load(resp))

    try:
        qs = urllib.parse.urlencode(dict(base, start_date=date_str, end_date=date_str))
        with urllib.request.urlopen("https://archive-api.open-meteo.com/v1/archive?" + qs, timeout=30) as resp:
            return _parse_daily(json.load(resp))
    except urllib.error.HTTPError as e:
        if e.code != 400:
            raise
        # Not archived yet (recent past) — use the forecast feed's past window.
        days_back = (datetime.date.today() - datetime.date.fromisoformat(date_str)).days
        qs = urllib.parse.urlencode(dict(base, past_days=days_back))
        with urllib.request.urlopen("https://api.open-meteo.com/v1/forecast?" + qs, timeout=30) as resp:
            data = json.load(resp)
        times = (data.get("daily") or {}).get("time", [])
        idx = times.index(date_str) if date_str in times else 0
        return _parse_daily(data, idx)


# ---- LLM (local liteLLM proxy, OpenAI-compatible) ----

def llm_analyze(d):
    """Return (tags, worry_controllable, summary_line). tags is a list."""
    prompt = (
        "你是我的个人日记分析助手。分析下面这篇日记，只输出严格 JSON，不要任何多余文字：\n"
        "{\"tags\": [\"标签1\", \"标签2\"], "
        "\"worry_controllable\": \"actionable\" 或 \"rumination\", "
        "\"summary_line\": \"一句话中文摘要\"}\n"
        "tags 是 1-3 个中文主题标签（工作/健康/关系/金钱/成长/家庭 等）。"
        "worry_controllable 判断：这条担忧是否有具体可执行的应对办法（填 actionable）"
        "还是不可控的反复思虑（填 rumination）；若没有担忧就填 null。\n\n"
        "日期：%s\n心情：%s/5\n担忧：%s\n亮点：%s\n明日计划：%s"
        % (d["date"], d.get("mood"), d.get("worry") or "无",
           d.get("highlight") or "无", d.get("tomorrow_plan") or "无")
    )
    body = {
        "model": LITELLM_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.2,
    }
    headers = {"Content-Type": "application/json"}
    if LITELLM_API_KEY:
        headers["Authorization"] = "Bearer " + LITELLM_API_KEY
    req = urllib.request.Request(LITELLM_URL, data=json.dumps(body).encode("utf-8"),
                                 headers=headers)
    with urllib.request.urlopen(req, timeout=120) as resp:
        data = json.load(resp)
    content = data["choices"][0]["message"]["content"]
    text = content.strip().strip("`")
    if text.startswith("json"):
        text = text[4:].strip()
    try:
        obj = json.loads(text)
    except Exception:
        # Fall back to the first balanced {...} block if the model added prose.
        start = text.find("{")
        end = text.rfind("}")
        obj = json.loads(text[start:end + 1])
    tags = obj.get("tags") or []
    wc = obj.get("worry_controllable")
    if wc not in ("actionable", "rumination"):
        wc = None
    return tags, wc, (obj.get("summary_line") or "")


# ---- Main ----

def enrich_row(conn, r):
    d = datetime.date.fromisoformat(r["date"])
    now = datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")

    weekday = WEEKDAYS_CN[d.weekday()]
    iso_week = "%s-W%02d" % d.isocalendar()[:2]
    season = SEASONS_CN[d.month - 1]
    lunar = lunar_date_str(d)
    moon = moon_phase_str(d)

    weather = None
    weather_raw = None
    try:
        weather, weather_raw = weather_for(r["date"])
    except Exception as e:
        print("  weather failed for %s: %s" % (r["date"], e), file=sys.stderr)

    tags = None
    wc = None
    summary = None
    llm_ok = False
    try:
        tags, wc, summary = llm_analyze(r)
        llm_ok = True
    except Exception as e:
        print("  llm failed for %s: %s" % (r["date"], e), file=sys.stderr)

    # Only mark enriched_at when the LLM pass succeeded — otherwise a later
    # run retries it. Env fields above are still written either way.
    with conn:
        conn.execute(
            "UPDATE diary SET weekday=?, iso_week=?, season=?, lunar_date=?, moon_phase=?,"
            " weather=?, weather_raw=?, tags=?, worry_controllable=?, summary_line=?,"
            " enriched_at=? WHERE date=?",
            (weekday, iso_week, season, lunar, moon, weather, weather_raw,
             json.dumps(tags, ensure_ascii=False) if tags is not None else None,
             wc, summary, now if llm_ok else None, r["date"]),
        )


def main():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
            conn.executescript(f.read())
        rows = conn.execute(
            "SELECT date, mood, worry, highlight, tomorrow_plan FROM diary"
            " WHERE enriched_at IS NULL ORDER BY date"
        ).fetchall()
        for r in rows:
            enrich_row(conn, dict(r))
        print("enriched %d diary row(s)" % len(rows))
        return 0
    finally:
        conn.close()


if __name__ == "__main__":
    sys.exit(main())
