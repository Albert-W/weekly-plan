#!/usr/bin/env bash
#
# Daily diary sync: Google Sheet → local SQLite (pull + enrich).
#
# Why this wrapper exists: the cron entry ran `python3 pull_diary.py` with the
# home directory as its working directory, so it failed 57 nights in a row with
#     can't open file '/Users/yichangwu/pull_diary.py': [Errno 2] ...
# while the diary silently stopped reaching SQLite (local bot.db stalled at
# 2026-08-17 while the Sheet had entries through 2026-09-23).
#
# This script cd's to its own directory, so the caller's working directory can
# never matter. Point cron at it with absolute paths:
#
#     15 0 * * * /Users/yichangwu/Documents/weekly-plan/local/sync_diary.sh \
#         >> /Users/yichangwu/Documents/weekly-plan/local/diary_cron.log 2>&1
#
# Safe to re-run by hand at any time: pull_diary.py upserts by date, and
# enrich_diary.py skips rows it has already enriched (enriched_at guard).

set -uo pipefail

cd "$(dirname "$0")" || exit 1

# cron runs with a minimal PATH ("/usr/bin:/bin"), so resolve the interpreter
# explicitly rather than trusting `python3` to be found — and prefer Homebrew's,
# which is what the old working enrich line used.
PY=""
for cand in /opt/homebrew/bin/python3 "$(command -v python3 2>/dev/null || true)" /usr/bin/python3; do
  if [ -n "$cand" ] && [ -x "$cand" ]; then PY="$cand"; break; fi
done
if [ -z "$PY" ]; then
  echo "sync_diary.sh: no python3 found" >&2
  exit 127
fi

status=0

echo "=== $(date '+%Y-%m-%d %H:%M:%S') sync start (py=$PY cwd=$(pwd)) ==="

# pull first (form inputs + system snapshot), then enrich (weather + LLM).
# No `set -e`: a failed pull must not stop enrich from retrying other rows,
# but the script still exits non-zero so cron mail/log shows the failure.
"$PY" pull_diary.py || status=$?
"$PY" enrich_diary.py || status=$?

echo "=== $(date '+%Y-%m-%d %H:%M:%S') sync done (status=$status) ==="
exit "$status"
