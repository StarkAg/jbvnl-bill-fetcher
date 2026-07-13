#!/usr/bin/env python3
"""
Telegram notifier for the JBVNL bill-fetcher.

Reads TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID from the environment, or falls back
to the existing shared config at ~/.config/whatsapp-mcp/bridge.env — so no token is
duplicated into this repo or committed anywhere.

Standalone test:
    python notify.py "hello from bill-fetcher"
"""

import os
import json
import urllib.request
import urllib.parse
from pathlib import Path

BRIDGE_ENV = Path.home() / ".config" / "whatsapp-mcp" / "bridge.env"


def _load_creds():
    tok = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat = os.environ.get("TELEGRAM_CHAT_ID")
    if (not tok or not chat) and BRIDGE_ENV.exists():
        for line in BRIDGE_ENV.read_text().splitlines():
            line = line.strip()
            if line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            k, v = k.strip(), v.strip().strip('"').strip("'")
            if k == "TELEGRAM_BOT_TOKEN" and not tok:
                tok = v
            elif k == "TELEGRAM_CHAT_ID" and not chat:
                chat = v
    return tok, chat


def send(text: str) -> bool:
    """Send an HTML message to the configured Telegram chat. Returns True on success."""
    tok, chat = _load_creds()
    if not tok or not chat:
        print("[notify] Telegram creds not found (env or bridge.env) — skipping notification")
        return False
    url = f"https://api.telegram.org/bot{tok}/sendMessage"
    payload = urllib.parse.urlencode(
        {"chat_id": chat, "text": text, "parse_mode": "HTML", "disable_web_page_preview": "true"}
    ).encode()
    try:
        req = urllib.request.Request(url, data=payload)
        with urllib.request.urlopen(req, timeout=15) as r:
            return bool(json.loads(r.read().decode()).get("ok"))
    except Exception as e:
        print(f"[notify] Telegram send failed: {e}")
        return False


if __name__ == "__main__":
    import sys
    msg = sys.argv[1] if len(sys.argv) > 1 else "✅ JBVNL bill-fetcher: Telegram notifications are wired up."
    print("sent ✅" if send(msg) else "failed ❌")
