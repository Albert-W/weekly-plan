#!/usr/bin/env python3
"""Diary table column migrations, shared by pull_diary.py / enrich_diary.py /
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

# Columns removed after the first release (LLM fields that were never acted on).
DROPPED_COLUMNS = ("tags", "worry_controllable")


def ensure_columns(conn):
    """Bring a pre-existing `diary` table up to the current column set."""
    cols = {row[1] for row in conn.execute("PRAGMA table_info(diary)")}
    with conn:
        for name, decl in ADDED_COLUMNS:
            if name not in cols:
                conn.execute("ALTER TABLE diary ADD COLUMN %s %s" % (name, decl))
        for dropped in DROPPED_COLUMNS:
            if dropped in cols:
                conn.execute("ALTER TABLE diary DROP COLUMN " + dropped)
