#!/usr/bin/env python3
"""Per-day snapshots of the Weekly grid — independent of diary rows.

Why this exists: the Weekly grid lives in the Sheet and is *transient* —
"Start new week…" archives it (Archive sheet + Drive CSV) and clears it. Without
a local copy, "what did I plan on 2026-10-20" is only answerable by digging
through an archived week-shaped CSV.

Why a separate table (it used to be columns on `diary`): a day can have a plan
without a diary entry. The first version attached the snapshot to the diary row,
so a day you didn't journal silently kept nothing (2026-10-20/21 were lost that
way). Plans now stand on their own in `plan_day`, keyed by date.

The Web App's ?view=schedule returns one day's column of the grid (including the
calendar events Calendar.js synced into those same task cells). pull_diary.py
(hourly) and serve.py (right after a save) hand its payload here.

Contract: best-effort. Callers wrap this in try/except — a snapshot failure must
never break a pull or a save.
"""

import json


def capture(conn, date_str, schedule, captured_at):
    """Store `schedule`'s slots for `date_str`.

    Returns True when written. Refuses to write when the payload is for another
    day: the endpoint only knows the current week's grid, so a mismatch means
    "that day's cells are gone", and mislabelling yesterday's timetable as
    today's would quietly corrupt the record.
    """
    if not schedule or not schedule.get("ok"):
        return False
    if not date_str or schedule.get("day") != date_str:
        return False

    slots = schedule.get("slots") or []
    with conn:
        conn.execute(
            "INSERT INTO plan_day(date, slots_json, captured_at) VALUES(?,?,?)"
            " ON CONFLICT(date) DO UPDATE SET"
            " slots_json=excluded.slots_json, captured_at=excluded.captured_at",
            (date_str, json.dumps(slots, ensure_ascii=False), captured_at),
        )
    return True


def load(conn, date_str):
    """Stored slots for a date, or []."""
    row = conn.execute("SELECT slots_json FROM plan_day WHERE date=?", (date_str,)).fetchone()
    if not row or not row[0]:
        return []
    try:
        return json.loads(row[0])
    except Exception:
        return []


def load_all(conn):
    """{date: {'slots': [...], 'captured_at': iso}} for every stored day."""
    out = {}
    try:
        rows = conn.execute("SELECT date, slots_json, captured_at FROM plan_day").fetchall()
    except Exception:  # table not created yet (fresh DB, before ensure_columns)
        return out
    for date, slots_json, captured_at in rows:
        try:
            slots = json.loads(slots_json or "[]")
        except Exception:
            slots = []
        out[date] = {"slots": slots, "captured_at": captured_at}
    return out
