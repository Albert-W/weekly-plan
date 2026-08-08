#!/usr/bin/env python3
"""Pull diary rows from the Weekly Plan Web App into local SQLite.

Pure Python stdlib (urllib + sqlite3), no pip dependencies. The fetch is
server-side, so the Apps Script Web App's missing CORS header is irrelevant.

Run on a schedule (cron/launchd) or before opening the reader:
    python3 pull_diary.py

Upserts by date (INSERT OR REPLACE), so it is safe to run frequently: the
form inputs land on the first pull after a submission, and the system
snapshot (score / habits) is overwritten on later pulls once that day's
scores are final. The environment + LLM fields are filled separately by
enrich_diary.py and are untouched here.
"""

import json
import os
import sqlite3
import sys
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

WEBAPP_URL = os.environ.get("WEBAPP_URL", "")
AUTH_KEY = os.environ.get("SYNC_AUTH_KEY", "")       # same syncAuthKey the Web App checks
DB_PATH = os.path.join(HERE, os.environ.get("DB_PATH", "bot.db"))
SINCE = os.environ.get("SINCE", "20260101")          # YYYYMMDD lower bound (optional)

SCHEMA_PATH = os.path.join(HERE, "schema.sql")

# Columns written by this pass (form inputs + system snapshot). The env/LLM
# columns in schema.sql stay untouched so enrich_diary.py owns them.
COLUMNS = (
    "date", "mood", "worry", "highlight", "tomorrow_plan",
    "submitted_at", "updated_at",
    "summary_positive", "summary_negative", "summary_total", "habits_done",
)


def fetch_diary():
    qs = urllib.parse.urlencode({"view": "diary", "k": AUTH_KEY, "since": SINCE})
    with urllib.request.urlopen(WEBAPP_URL + "?" + qs, timeout=30) as resp:
        payload = json.load(resp)
    if not payload.get("ok"):
        raise RuntimeError("Web App error: %r" % payload.get("error"))
    return payload.get("diary", [])


def upsert(conn, rows):
    sql = (
        "INSERT OR REPLACE INTO diary(" + ",".join(COLUMNS) + ") VALUES(" +
        ",".join("?" for _ in COLUMNS) + ")"
    )
    vals = []
    for r in rows:
        s = r.get("summary")
        vals.append((
            r.get("date"), r.get("mood"), r.get("worry"), r.get("highlight"),
            r.get("tomorrow_plan"), r.get("submitted_at"), r.get("updated_at"),
            (s or {}).get("positive"), (s or {}).get("negative"),
            (s or {}).get("total"), r.get("habits_done"),
        ))
    with conn:
        conn.executemany(sql, vals)


def main():
    conn = sqlite3.connect(DB_PATH)
    try:
        with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
            conn.executescript(f.read())
        rows = fetch_diary()
        upsert(conn, rows)
        print("upserted %d diary row(s) into %s" % (len(rows), DB_PATH))
        return 0
    finally:
        conn.close()


if __name__ == "__main__":
    sys.exit(main())
