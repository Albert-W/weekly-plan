-- Daily Diary — local SQLite schema (the canonical rich store on the Mac).
--
-- Filled in two passes:
--   1. pull_diary.py   — form inputs + system snapshot, immediately (upsert by date).
--   2. enrich_diary.py — environment (weather / weekday / lunar / moon phase) + LLM
--                        fields, once the diary's day has ended. `enriched_at`
--                        marks completion so a re-run skips finished rows.
--
-- NOTE: `CREATE TABLE IF NOT EXISTS` only shapes a brand-new DB. Column changes
-- (added `improve`, the weather/air extras; dropped `tags`,
-- `worry_controllable`) are applied to existing DBs by the ensure_columns()
-- helper in pull_diary.py / enrich_diary.py / serve.py.
CREATE TABLE IF NOT EXISTS diary (
  date          TEXT PRIMARY KEY,               -- 'YYYY-MM-DD'
  -- Form inputs (first pull)
  mood          INTEGER,                        -- 1..5
  worry         TEXT DEFAULT '',                -- shown as "📌 今日要点"
  highlight     TEXT DEFAULT '',
  tomorrow_plan TEXT DEFAULT '',
  improve       TEXT DEFAULT '',                -- 有什么可以改进的地方吗？
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
  precip_prob   INTEGER,                        -- % chance of rain (daily max); NULL in the archive feed
  feels_min     REAL,                           -- apparent temperature min, °C
  feels_max     REAL,                           -- apparent temperature max, °C
  sunrise       TEXT,                           -- 'HH:MM' local
  sunset        TEXT,                           -- 'HH:MM' local
  aqi           INTEGER,                        -- European AQI, worst hour of the day
  pm25          REAL,                           -- µg/m³, worst hour of the day
  pollen        TEXT,                           -- JSON {"birch":n,"grass":n,"alder":n}; '{}' = measured, all zero
  -- Historical: the LLM pass that wrote this was removed 2026-10-17 (cost)
  summary_line  TEXT,                           -- one-line Chinese summary
  -- Bookkeeping
  enriched_at   TEXT                            -- ISO8601 when the env pass completed
);

-- Per-day snapshots of the Weekly grid. Deliberately NOT columns on `diary`:
-- a day can have a planned timetable without a journal entry, and the grid is
-- archived + cleared by "Start new week…", so this is the only durable copy of
-- the plan side of a day.
CREATE TABLE IF NOT EXISTS plan_day (
  date        TEXT PRIMARY KEY,                 -- 'YYYY-MM-DD'
  slots_json  TEXT NOT NULL,                    -- JSON [{"time","task","score"}]
  captured_at TEXT NOT NULL                     -- ISO8601 UTC of the last capture
);
