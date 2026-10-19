#!/usr/bin/env python3
"""Minimal local server for the diary reader.

Serves DiaryReader.html at / and /api/diary from the local SQLite DB — the
same JSON shape the GAS Web App returns, plus the enriched fields. Pure
Python stdlib (http.server + sqlite3). This is also a reference for how your
Mac bot should expose /api/diary; swap in the bot's framework when ready.

    python3 local/serve.py        # then open http://localhost:18000

Endpoints
    GET  /api/diary[?limit=N]   diary rows from local SQLite (newest first)
    POST /api/diary             save/edit one day → GAS (same Diary sheet row
                                the Google Form writes) → mirror into SQLite
    POST /api/refresh           re-pull the Sheet via GAS into SQLite
                                (same work as pull_diary.py)

The browser never sees syncAuthKey: the page is same-origin with this server
and only this server talks to the Web App.
"""

import json
import os
import re
import sqlite3
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, HTTPServer

try:
    import dotenv  # local/dotenv.py — tiny .env loader
except Exception:  # pragma: no cover - import path differs in some runners
    dotenv = None

from dbschema import ensure_columns  # noqa: E402 — shared column migrations

HERE = os.path.dirname(os.path.abspath(__file__))
if dotenv:
    dotenv.load_dotenv(os.path.join(HERE, ".env"))
DB_PATH = os.path.join(HERE, os.environ.get("DB_PATH", "bot.db"))
PORT = int(os.environ.get("PORT", "18000"))
READER = os.path.join(HERE, "DiaryReader.html")

# Web App + shared secret (same values pull_diary.py uses).
WEBAPP_URL = os.environ.get("WEBAPP_URL", "")
AUTH_KEY = os.environ.get("SYNC_AUTH_KEY", "")

# "Open Weekly Plan" link on the diary page. Get the value once from the Web
# App (?view=links&k=...) and paste it into local/.env — the page must not wait
# on Apps Script, and this keeps the page working offline. Empty = no button.
WEEKLY_PLAN_URL = os.environ.get("WEEKLY_PLAN_URL", "")

# Columns an editor save owns. Everything else on the row (summary_*,
# habits_done, the enrich_diary.py environment/LLM columns) is preserved —
# same rule as pull_diary.py.
FORM_COLUMNS = (
    "mood", "worry", "highlight", "tomorrow_plan", "improve",
    "submitted_at", "updated_at",
)

# Which of the diary's fields the reader exposes as editable, mapped to the
# JSON keys the editor sends (identical names, kept explicit for clarity).
EDITABLE_FIELDS = ("mood", "worry", "highlight", "tomorrow_plan", "improve")


def row_to_dict(r):
    """SQLite row → the JSON shape the reader expects (and the GAS Web App
    returns), including the enrich_diary.py fields."""
    has_summary = any(r[k] is not None for k in ("summary_positive", "summary_negative", "summary_total"))
    return {
        "date": r["date"],
        "mood": r["mood"],
        "worry": r["worry"] or "",
        "highlight": r["highlight"] or "",
        "tomorrow_plan": r["tomorrow_plan"] or "",
        "improve": r["improve"] or "",
        "submitted_at": r["submitted_at"],
        "updated_at": r["updated_at"],
        "summary": {
            "positive": r["summary_positive"],
            "negative": r["summary_negative"],
            "total": r["summary_total"],
        } if has_summary else None,
        "habits_done": r["habits_done"],
        "weekday": r["weekday"],
        "iso_week": r["iso_week"],
        "season": r["season"],
        "lunar_date": r["lunar_date"],
        "moon_phase": r["moon_phase"],
        "weather": r["weather"],
        "precip_prob": r["precip_prob"],
        "feels_min": r["feels_min"],
        "feels_max": r["feels_max"],
        "sunrise": r["sunrise"],
        "sunset": r["sunset"],
        "aqi": r["aqi"],
        "pm25": r["pm25"],
        "pollen": json_or_none(r["pollen"]),
        "summary_line": r["summary_line"],
    }


def json_or_none(s):
    """Stored JSON string → object; None/''/garbage → None.
    '{}' survives as {} ("measured, nothing to report") vs None ("not measured")."""
    if not s:
        return None
    try:
        return json.loads(s)
    except Exception:
        return None


def load_diary(limit):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        ensure_columns(conn)
        sql = "SELECT * FROM diary ORDER BY date DESC"
        args = ()
        if limit:
            sql += " LIMIT ?"
            args = (limit,)
        return [row_to_dict(r) for r in conn.execute(sql, args).fetchall()]
    finally:
        conn.close()


def load_diary_one(date_str):
    """One diary row (ISO date) or None — used to echo a save back to the page
    with the local-only enrichment fields intact."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        ensure_columns(conn)
        r = conn.execute("SELECT * FROM diary WHERE date = ?", (date_str,)).fetchone()
        return row_to_dict(r) if r else None
    finally:
        conn.close()


def upsert_form_row(conn, date_str, values):
    """Mirror an editor save into SQLite: form columns only, upsert by date.

    ON CONFLICT DO UPDATE (not INSERT OR REPLACE) so the system snapshot and
    the enrichment columns survive — exactly what pull_diary.py does.
    `values` maps FORM_COLUMNS names → values.
    """
    cols = ("date",) + FORM_COLUMNS
    assignments = ", ".join("%s=excluded.%s" % (c, c) for c in FORM_COLUMNS)
    sql = (
        "INSERT INTO diary(" + ",".join(cols) + ") VALUES("
        + ",".join("?" for _ in cols) + ")"
        + " ON CONFLICT(date) DO UPDATE SET " + assignments
    )
    vals = (date_str,) + tuple(values.get(c) for c in FORM_COLUMNS)
    with conn:
        conn.execute(sql, vals)


# ---- Web App (GAS) client ----

def gas_post(payload):
    """POST JSON to the Web App and return the parsed response.

    Apps Script answers with a 302 to script.googleusercontent.com; urllib
    follows it as a GET, which is where the script's output actually lives.
    """
    if not WEBAPP_URL or not AUTH_KEY:
        raise RuntimeError("local/.env 缺少 WEBAPP_URL 或 SYNC_AUTH_KEY")
    url = WEBAPP_URL + "?" + urllib.parse.urlencode({"k": AUTH_KEY})
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url, data=body, method="POST",
        headers={"Content-Type": "application/json; charset=utf-8"},
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8"))


def gas_pull():
    """Re-run the Sheet → SQLite pull in-process (same code path as
    pull_diary.py, so the two can't drift)."""
    if HERE not in sys.path:
        sys.path.insert(0, HERE)
    import pull_diary  # noqa: E402 — local sibling module
    pull_diary.main()
    conn = sqlite3.connect(DB_PATH)
    try:
        n = conn.execute("SELECT COUNT(*) FROM diary").fetchone()[0]
    finally:
        conn.close()
    return n


def enrich_one_local(date_str):
    """Fill one day's environment fields (weather / air / lunar / moon) now.

    Without this the reader would show today's card bare until the 00:15 batch
    run. Imported lazily for the same reason as gas_pull.
    """
    if HERE not in sys.path:
        sys.path.insert(0, HERE)
    import enrich_diary  # noqa: E402 — local sibling module
    return enrich_diary.enrich_one(date_str, DB_PATH)


# ---- Today's schedule (Web App ?view=schedule, cached briefly) ----
# The grid changes as the day is edited, but a page load shouldn't wait on Apps
# Script every time — 5 minutes is a good trade for a personal dashboard.
_SCHEDULE_CACHE = {"at": 0.0, "data": None}
SCHEDULE_TTL = 300


def gas_get_schedule():
    """Today's timetable slots (incl. the events Calendar.js synced in)."""
    global _SCHEDULE_CACHE
    now = time.time()
    if _SCHEDULE_CACHE["data"] is not None and now - _SCHEDULE_CACHE["at"] < SCHEDULE_TTL:
        return _SCHEDULE_CACHE["data"]
    if not WEBAPP_URL or not AUTH_KEY:
        raise RuntimeError("local/.env 缺少 WEBAPP_URL 或 SYNC_AUTH_KEY")
    url = WEBAPP_URL + "?" + urllib.parse.urlencode({"view": "schedule", "k": AUTH_KEY})
    with urllib.request.urlopen(url, timeout=30) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    _SCHEDULE_CACHE = {"at": now, "data": data}
    return data




class Handler(BaseHTTPRequestHandler):
    def _send(self, code, content_type, body):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(code, "application/json; charset=utf-8",
                   json.dumps(obj, ensure_ascii=False).encode("utf-8"))

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/diary":
            params = urllib.parse.parse_qs(parsed.query)
            limit = 0
            try:
                limit = int(params.get("limit", ["0"])[0])
            except ValueError:
                limit = 0
            self._json({
                "ok": True,
                "diary": load_diary(limit or None),
                "weekly_plan_url": WEEKLY_PLAN_URL,
            })
        elif parsed.path in ("/", "/index.html"):
            with open(READER, "rb") as f:
                self._send(200, "text/html; charset=utf-8", f.read())
        elif parsed.path == "/api/schedule":
            try:
                self._json(gas_get_schedule())
            except Exception as e:
                self._json({"ok": False, "error": "取今日排程失败: %s" % e}, 502)
        else:
            self.send_error(404)

    # ---- POST: the reader's editor ----

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/diary":
            self._api_save_diary()
        elif parsed.path == "/api/refresh":
            self._api_refresh()
        else:
            self.send_error(404)

    def _read_json(self):
        """Request body → dict, or None after answering 400."""
        try:
            n = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            n = 0
        raw = self.rfile.read(n) if n else b""
        try:
            body = json.loads(raw.decode("utf-8") or "{}")
        except Exception as e:
            self._json({"ok": False, "error": "无效的 JSON: %s" % e}, 400)
            return None
        if not isinstance(body, dict):
            self._json({"ok": False, "error": "JSON 必须是对象"}, 400)
            return None
        return body

    def _api_save_diary(self):
        """Save one day. The write goes to the Sheet through the Web App (the
        single source of truth, shared with the Google Form), then the same
        row is mirrored into SQLite so the page updates immediately."""
        body = self._read_json()
        if body is None:
            return

        date_str = str(body.get("date") or "").strip()
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", date_str):
            self._json({"ok": False, "error": "日期格式应为 YYYY-MM-DD"}, 400)
            return

        payload = {"action": "diary", "date": date_str}
        for k in EDITABLE_FIELDS:
            payload[k] = body.get(k, "")

        try:
            res = gas_post(payload)
        except Exception as e:  # network / config problem
            self._json({"ok": False, "error": "调用 Web App 失败: %s" % e}, 502)
            return

        if not res.get("ok"):
            # e.g. the deployed Web App predates action=diary, or auth failed
            self._json({"ok": False, "error": res.get("error") or "Web App 返回失败"}, 502)
            return

        # The Sheet is authoritative; echo its stored row (submitted_at is
        # preserved on an update, so the page must not invent it).
        row = res.get("row") or {}
        if not row.get("date"):
            # Write succeeded but the read-back came back empty: leave the
            # local row alone rather than blanking its columns, and echo what
            # the page submitted so the UI still shows the save.
            self._json({
                "ok": True,
                "created": bool(res.get("created")),
                "diary": dict({"date": date_str}, **{k: payload.get(k) for k in EDITABLE_FIELDS}),
            })
            return

        try:
            conn = sqlite3.connect(DB_PATH)
            try:
                ensure_columns(conn)
                upsert_form_row(conn, date_str, {
                    "mood": row.get("mood"),
                    "worry": row.get("worry") or "",
                    "highlight": row.get("highlight") or "",
                    "tomorrow_plan": row.get("tomorrow_plan") or "",
                    "improve": row.get("improve") or "",
                    "submitted_at": row.get("submitted_at"),
                    "updated_at": row.get("updated_at"),
                })
            finally:
                conn.close()
        except Exception as e:
            self._json({"ok": False, "error": "已写入表格，但本地缓存更新失败: %s" % e}, 500)
            return

        # Fill this day's environment fields (weather / lunar / moon) right
        # away, so the card shows today's weather instead of waiting for the
        # 00:15 batch. Best-effort: an Open-Meteo hiccup must not fail the save.
        try:
            enrich_one_local(date_str)
        except Exception as e:
            print("enrich_one failed for %s: %s" % (date_str, e), file=sys.stderr)

        self._json({
            "ok": True,
            "created": bool(res.get("created")),
            "diary": load_diary_one(date_str),
        })

    def _api_refresh(self):
        """Re-pull the Sheet into SQLite (the same run pull_diary.py does)."""
        try:
            n = gas_pull()
        except Exception as e:
            self._json({"ok": False, "error": "同步失败: %s" % e}, 502)
            return
        self._json({"ok": True, "rows": n})


if __name__ == "__main__":
    print("Diary reader → http://localhost:%d  (Ctrl+C to stop)" % PORT)
    HTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
