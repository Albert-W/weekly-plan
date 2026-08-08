"""Tiny .env loader (stdlib-only).

Reads KEY=VALUE lines from a .env file into os.environ without overriding
variables that are already set (so a shell env var always wins). Comments
start with '#', and values may be wrapped in single/double quotes.
"""

import os


def load_dotenv(path):
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip().strip("'\"")
            if key and key not in os.environ:
                os.environ[key] = value
