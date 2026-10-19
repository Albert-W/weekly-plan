#!/usr/bin/env python3
"""Enrich local diary entries with environment fields — no LLM, no API keys.

Runs after a diary's day has ended. For each row where `enriched_at` is NULL:
  - weekday / ISO week / season            (stdlib)
  - Chinese lunar date                     (classic LUNAR_INFO table, 1900-2100)
  - moon phase                             (classic synodic-month formula)
  - weather for that date                  (Open-Meteo, no key)

LLM fields were removed on request (2026-10-17: cost), after tags and
worry_controllable (2026-10-15). There is no LLM call left in this pipeline at
all — the only network calls are Open-Meteo. The `summary_line` COLUMN is kept
(not dropped) so the summaries already generated stay on their rows; nothing
writes it any more, and the reader hides the summary block when it is empty.

Pure Python stdlib (urllib + sqlite3).

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

from dbschema import ensure_columns  # noqa: E402 — shared column migrations

# ---- Config (edit me in local/.env, NOT in this file) ----
HERE = os.path.dirname(os.path.abspath(__file__))
if dotenv:
    dotenv.load_dotenv(os.path.join(HERE, ".env"))

DB_PATH = os.path.join(HERE, os.environ.get("DB_PATH", "bot.db"))
LAT = float(os.environ.get("LAT", "53.3498"))       # Dublin; set to your city
LON = float(os.environ.get("LON", "-6.2603"))
WEATHER_TZ = os.environ.get("WEATHER_TZ", "Europe/Dublin")  # raw — urlencode encodes it

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
# case we fall back to the forecast feed's `past_days` window. The archive has
# no precipitation probability (and no UV), so those columns stay NULL for past
# days — that is a limitation of the upstream feed, not a bug here.

DAILY_VARS = ",".join((
    "weathercode", "temperature_2m_max", "temperature_2m_min",
    "apparent_temperature_max", "apparent_temperature_min",
    "precipitation_probability_max", "sunrise", "sunset",
))


def _round_or_none(v, nd=1):
    return round(v, nd) if isinstance(v, (int, float)) else None


def _int_or_none(v):
    return int(round(v)) if isinstance(v, (int, float)) else None


def _hhmm(iso):
    """Open-Meteo '2026-10-17T07:55' → '07:55' (None when absent)."""
    if isinstance(iso, str) and "T" in iso:
        return iso.split("T", 1)[1][:5]
    return None


def _parse_daily(data, idx=0):
    """Parse one day out of a /daily payload into the columns we store."""
    daily = data.get("daily") or {}

    def at(key):
        v = daily.get(key)
        if isinstance(v, list):
            return v[idx] if idx < len(v) else None
        return v

    wc = at("weathercode")
    tmax = at("temperature_2m_max")
    tmin = at("temperature_2m_min")
    cond = WMO_CN.get(wc, "未知")
    if tmax is None:
        weather = cond
    elif tmin is not None:
        weather = "%s %s°C/%s°C" % (cond, round(tmin), round(tmax))
    else:
        weather = "%s %s°C" % (cond, round(tmax))
    return {
        "weather": weather,
        "weather_raw": json.dumps(daily, ensure_ascii=False),
        "precip_prob": _int_or_none(at("precipitation_probability_max")),
        "feels_min": _round_or_none(at("apparent_temperature_min")),
        "feels_max": _round_or_none(at("apparent_temperature_max")),
        "sunrise": _hhmm(at("sunrise")),
        "sunset": _hhmm(at("sunset")),
    }


def weather_for(date_str):
    today = datetime.date.today().isoformat()
    base = {
        "latitude": LAT, "longitude": LON,
        "daily": DAILY_VARS,
        "timezone": WEATHER_TZ,
    }

    if date_str >= today:
        # The forecast feed's daily arrays start TODAY, so a future date (the
        # editor lets you pick one) needs enough days and the right index —
        # otherwise you'd silently label today's weather as tomorrow's.
        days_ahead = (datetime.date.fromisoformat(date_str)
                      - datetime.date.today()).days
        qs = urllib.parse.urlencode(dict(base, forecast_days=days_ahead + 1))
        with urllib.request.urlopen("https://api.open-meteo.com/v1/forecast?" + qs, timeout=30) as resp:
            return _parse_daily(json.load(resp), days_ahead)

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


# ---- Air quality + pollen (Open-Meteo air-quality API, no key) ----
# Hourly series for the diary's date, collapsed to the worst hour: for a plan
# you want the peak, not the average. Works for past days too (the feed keeps a
# few months of history), which matters because the batch run enriches
# yesterday — a "current conditions only" call would never fire for it.

POLLEN_VARS = (("birch_pollen", "birch"), ("grass_pollen", "grass"), ("alder_pollen", "alder"))


def air_for(date_str):
    """{'aqi','pm25','pollen'} for that date, or None when the feed has nothing."""
    qs = urllib.parse.urlencode({
        "latitude": LAT, "longitude": LON, "timezone": WEATHER_TZ,
        "hourly": "european_aqi,pm2_5," + ",".join(k for k, _ in POLLEN_VARS),
        "start_date": date_str, "end_date": date_str,
    })
    with urllib.request.urlopen(
            "https://air-quality-api.open-meteo.com/v1/air-quality?" + qs, timeout=30) as resp:
        hourly = (json.load(resp).get("hourly") or {})
    if not hourly.get("time"):
        return None

    def peak(key):
        vals = [v for v in (hourly.get(key) or []) if isinstance(v, (int, float))]
        return max(vals) if vals else None

    pollen = {}
    for key, name in POLLEN_VARS:
        v = peak(key)
        if v is not None and v > 0:
            pollen[name] = round(v, 1)
    return {
        "aqi": _int_or_none(peak("european_aqi")),
        "pm25": _round_or_none(peak("pm2_5")),
        # '{}' means "measured, all zero" — distinct from None ("not measured").
        "pollen": json.dumps(pollen, ensure_ascii=False),
    }


# ---- Main ----

def enrich_row(conn, r):
    d = datetime.date.fromisoformat(r["date"])
    now = datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")

    weekday = WEEKDAYS_CN[d.weekday()]
    iso_week = "%s-W%02d" % d.isocalendar()[:2]
    season = SEASONS_CN[d.month - 1]
    lunar = lunar_date_str(d)
    moon = moon_phase_str(d)

    wx = {}
    try:
        wx = weather_for(r["date"]) or {}
    except Exception as e:
        print("  weather failed for %s: %s" % (r["date"], e), file=sys.stderr)

    # Air + pollen are secondary: if this call fails the row is still stamped
    # (below), so one flaky response can't block the day forever. Clear
    # enriched_at by hand to force a retry for a specific date.
    air = {}
    try:
        air = air_for(r["date"]) or {}
    except Exception as e:
        print("  air failed for %s: %s" % (r["date"], e), file=sys.stderr)

    # Stamp enriched_at only when the weather actually landed: Open-Meteo
    # returns transient 5xx now and then, and stamping anyway would leave that
    # day permanently weather-less (the batch run only picks up NULL rows).
    # The env fields are deterministic, so re-writing them costs nothing.
    with conn:
        conn.execute(
            "UPDATE diary SET weekday=?, iso_week=?, season=?, lunar_date=?, moon_phase=?,"
            " weather=?, weather_raw=?, precip_prob=?, feels_min=?, feels_max=?,"
            " sunrise=?, sunset=?, aqi=?, pm25=?, pollen=?, enriched_at=? WHERE date=?",
            (weekday, iso_week, season, lunar, moon,
             wx.get("weather"), wx.get("weather_raw"),
             wx.get("precip_prob"), wx.get("feels_min"), wx.get("feels_max"),
             wx.get("sunrise"), wx.get("sunset"),
             air.get("aqi"), air.get("pm25"), air.get("pollen"),
             now if wx.get("weather") is not None else None, r["date"]),
        )


def enrich_one(date_str, db_path=None):
    """Enrich one date right now; True if the row existed and ran.

    The reader calls this after a save so today's card shows today's weather /
    lunar / moon immediately instead of waiting for the 00:15 batch run. Weather
    for a date >= today comes from the live forecast feed (see weather_for), so
    this works for the current day. Callers treat it as best-effort.
    """
    conn = sqlite3.connect(db_path or DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        ensure_columns(conn)
        r = conn.execute(
            "SELECT date, mood, worry, highlight, tomorrow_plan, improve FROM diary WHERE date=?",
            (date_str,),
        ).fetchone()
        if not r:
            return False
        enrich_row(conn, dict(r))
        return True
    finally:
        conn.close()


def main():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
            conn.executescript(f.read())
        ensure_columns(conn)
        rows = conn.execute(
            "SELECT date, mood, worry, highlight, tomorrow_plan, improve FROM diary"
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
