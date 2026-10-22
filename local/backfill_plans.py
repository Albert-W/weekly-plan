#!/usr/bin/env python3
"""Backfill `plan_day` for the current week from the Weekly grid.

Repairs the gap left by days when the hourly sync didn't run (Mac asleep, cron
skipped, etc.). Only the CURRENT week can be recovered: the grid is archived and
cleared at week rollover, so older days have nothing left to read.

    python3 backfill_plans.py             # Monday .. today
    python3 backfill_plans.py --dry-run   # show what would be written

Idempotent: plan_day rows are upserted by date.
"""

import datetime
import json
import os
import sqlite3
import sys
import urllib.parse
import urllib.request

try:
    import dotenv  # local/dotenv.py — tiny .env loader
except Exception:  # pragma: no cover
    dotenv = None

from dbschema import ensure_columns  # noqa: E402 — shared schema upgrades
import plan_snapshot  # noqa: E402 — same folder

HERE = os.path.dirname(os.path.abspath(__file__))
if dotenv:
    dotenv.load_dotenv(os.path.join(HERE, ".env"))

WEBAPP_URL = os.environ.get("WEBAPP_URL", "")
AUTH_KEY = os.environ.get("SYNC_AUTH_KEY", "")
DB_PATH = os.path.join(HERE, os.environ.get("DB_PATH", "bot.db"))


def fetch_schedule(date_str):
    qs = urllib.parse.urlencode({"view": "schedule", "k": AUTH_KEY, "date": date_str})
    with urllib.request.urlopen(WEBAPP_URL + "?" + qs, timeout=30) as resp:
        return json.load(resp)


def main(argv):
    dry = "--dry-run" in argv
    today = datetime.date.today()
    monday = today - datetime.timedelta(days=today.weekday())

    conn = sqlite3.connect(DB_PATH)
    try:
        ensure_columns(conn)
        day = monday
        wrote = 0
        while day <= today:
            date_str = day.isoformat()
            try:
                sched = fetch_schedule(date_str)
            except Exception as e:
                print("%s  fetch failed: %s" % (date_str, e))
                day += datetime.timedelta(days=1)
                continue
            slots = sched.get("slots") or []
            label = "%s  %d slot(s)" % (date_str, len(slots))
            if dry:
                print(label + "  (dry-run)")
            else:
                ok = plan_snapshot.capture(
                    conn, date_str, sched,
                    datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"),
                )
                wrote += 1 if ok else 0
                print(label + ("  saved" if ok else "  skipped"))
            day += datetime.timedelta(days=1)
        if not dry:
            print("backfilled %d day(s) into %s" % (wrote, DB_PATH))
        return 0
    finally:
        conn.close()


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
