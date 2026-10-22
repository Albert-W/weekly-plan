#!/usr/bin/env python3
"""Schema upgrades for the diary DB, shared by pull_diary.py / enrich_diary.py /
serve.py.

schema.sql only shapes a brand-new DB (`CREATE TABLE IF NOT EXISTS`), so a
database created by an earlier version needs explicit ALTERs. Everything here is
idempotent and safe to run on every request / every pass.
"""

# Columns added after the first release: (name, SQL declaration). Order does not
# matter — SQLite appends them at the end of the table.
ADDED_COLUMNS = (
    ("improve", "TEXT DEFAULT ''"),        # 有什么可以改进的地方吗？
    ("precip_prob", "INTEGER"),            # % chance of rain, daily max
    ("feels_min", "REAL"),                 # apparent temperature min, °C
    ("feels_max", "REAL"),                 # apparent temperature max, °C
    ("sunrise", "TEXT"),                   # 'HH:MM' local
    ("sunset", "TEXT"),                    # 'HH:MM' local
    ("aqi", "INTEGER"),                    # European AQI, worst hour of the day
    ("pm25", "REAL"),                      # µg/m³, worst hour of the day
    ("pollen", "TEXT"),                    # JSON {"birch":n,"grass":n,"alder":n}
)

# Columns removed after the first release. `plan_json`/`plan_captured_at` were
# the first (row-attached) attempt at storing a day's plan; plans now live in
# their own `plan_day` table so a day can have a plan without a diary entry.
DROPPED_COLUMNS = ("tags", "worry_controllable", "plan_json", "plan_captured_at")

# Per-day snapshots of the Weekly grid, keyed by date rather than by diary row.
PLAN_DAY_DDL = """
CREATE TABLE IF NOT EXISTS plan_day (
  date        TEXT PRIMARY KEY,   -- 'YYYY-MM-DD'
  slots_json  TEXT NOT NULL,      -- JSON [{"time","task","score"}]
  captured_at TEXT NOT NULL       -- ISO8601 UTC of the last capture
)"""


def ensure_columns(conn):
    """Bring a pre-existing DB up to the current schema (columns + tables)."""
    cols = {row[1] for row in conn.execute("PRAGMA table_info(diary)")}
    with conn:
        for name, decl in ADDED_COLUMNS:
            if name not in cols:
                conn.execute("ALTER TABLE diary ADD COLUMN %s %s" % (name, decl))

        conn.execute(PLAN_DAY_DDL)

        # One-time move of row-attached snapshots into plan_day before the
        # columns are dropped below.
        if "plan_json" in cols:
            moved = conn.execute(
                "SELECT date, plan_json, plan_captured_at FROM diary"
                " WHERE plan_json IS NOT NULL"
            ).fetchall()
            for date, slots_json, captured_at in moved:
                conn.execute(
                    "INSERT INTO plan_day(date, slots_json, captured_at) VALUES(?,?,?)"
                    " ON CONFLICT(date) DO NOTHING",
                    (date, slots_json, captured_at or ""),
                )

        for dropped in DROPPED_COLUMNS:
            if dropped in cols:
                conn.execute("ALTER TABLE diary DROP COLUMN " + dropped)
