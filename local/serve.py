#!/usr/bin/env python3
"""Minimal local server for the diary reader.

Serves DiaryReader.html at / and /api/diary from the local SQLite DB — the
same JSON shape the GAS Web App returns, plus the enriched fields. Pure
Python stdlib (http.server + sqlite3). This is also a reference for how your
Mac bot should expose /api/diary; swap in the bot's framework when ready.

    python3 local/serve.py        # then open http://localhost:18000
"""

import json
import os
import sqlite3
import urllib.parse
from http.server import BaseHTTPRequestHandler, HTTPServer

try:
    import dotenv  # local/dotenv.py — tiny .env loader
except Exception:  # pragma: no cover - import path differs in some runners
    dotenv = None

HERE = os.path.dirname(os.path.abspath(__file__))
if dotenv:
    dotenv.load_dotenv(os.path.join(HERE, ".env"))
DB_PATH = os.path.join(HERE, os.environ.get("DB_PATH", "bot.db"))
PORT = int(os.environ.get("PORT", "18000"))
READER = os.path.join(HERE, "DiaryReader.html")


def load_diary(limit):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        sql = "SELECT * FROM diary ORDER BY date DESC"
        args = ()
        if limit:
            sql += " LIMIT ?"
            args = (limit,)
        rows = conn.execute(sql, args).fetchall()
        out = []
        for r in rows:
            tags = None
            if r["tags"]:
                try:
                    tags = json.loads(r["tags"])
                except Exception:
                    tags = None
            has_summary = any(r[k] is not None for k in ("summary_positive", "summary_negative", "summary_total"))
            out.append({
                "date": r["date"],
                "mood": r["mood"],
                "worry": r["worry"] or "",
                "highlight": r["highlight"] or "",
                "tomorrow_plan": r["tomorrow_plan"] or "",
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
                "tags": tags,
                "worry_controllable": r["worry_controllable"],
                "summary_line": r["summary_line"],
            })
        return out
    finally:
        conn.close()


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
            self._json({"ok": True, "diary": load_diary(limit or None)})
        elif parsed.path in ("/", "/index.html"):
            with open(READER, "rb") as f:
                self._send(200, "text/html; charset=utf-8", f.read())
        else:
            self.send_error(404)


if __name__ == "__main__":
    print("Diary reader → http://localhost:%d  (Ctrl+C to stop)" % PORT)
    HTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
