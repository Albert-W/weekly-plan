-- Daily Diary — local SQLite schema (the canonical rich store on the Mac).
--
-- Filled in two passes:
--   1. pull_diary.py   — form inputs + system snapshot, immediately (upsert by date).
--   2. enrich_diary.py — environment (weather / weekday / lunar / moon phase) + LLM
--                        fields, once the diary's day has ended. `enriched_at`
--                        marks completion so a re-run skips finished rows.
CREATE TABLE IF NOT EXISTS diary (
  date          TEXT PRIMARY KEY,               -- 'YYYY-MM-DD'
  -- Form inputs (first pull)
  mood          INTEGER,                        -- 1..5
  worry         TEXT DEFAULT '',
  highlight     TEXT DEFAULT '',
  tomorrow_plan TEXT DEFAULT '',
  submitted_at  TEXT,
  updated_at    TEXT,
  -- System snapshot (from the GAS ?view=diary payload join)
  summary_positive REAL,
  summary_negative REAL,
  summary_total   REAL,
  habits_done     INTEGER,
  -- Environment (Mac-computed for the diary's date)
  weekday       TEXT,                           -- e.g. '星期四'
  iso_week      TEXT,                           -- e.g. '2026-W32'
  season        TEXT,                           -- '春' | '夏' | '秋' | '冬'
  lunar_date    TEXT,                           -- e.g. '六月廿三'
  moon_phase    TEXT,                           -- e.g. '盈凸月'
  weather       TEXT,                           -- human summary, e.g. '晴 22°C'
  weather_raw   TEXT,                           -- optional detailed JSON from Open-Meteo
  -- LLM-derived (Gemini)
  tags          TEXT,                           -- JSON array of 1-3 Chinese tags
  worry_controllable TEXT,                      -- 'actionable' | 'rumination' | NULL
  summary_line  TEXT,                           -- one-line Chinese summary
  -- Bookkeeping
  enriched_at   TEXT                            -- ISO8601 when env+LLM were written
);
