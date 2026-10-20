"""Send the PNG snapshots to a Telegram group.

The demo uses the Telegram Bot API (just a token + chat id). In production the same
step ran through Telethon with a user session, because the target group did not allow
bots - the flow is identical: one photo per table, each with its own caption.
"""
from __future__ import annotations

from pathlib import Path

import requests

import config


def send_photo(png: Path, caption: str) -> None:
    if config.DRY_RUN:
        print(f"[telegram:dry-run] would send {png.name} | caption: {caption}")
        return
    if not (config.TELEGRAM_BOT_TOKEN and config.TELEGRAM_CHAT_ID):
        raise RuntimeError("Set TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID, or keep DRY_RUN=1")

    url = f"https://api.telegram.org/bot{config.TELEGRAM_BOT_TOKEN}/sendDocument"
    # sendDocument instead of sendPhoto: Telegram keeps the full resolution.
    with png.open("rb") as f:
        r = requests.post(url, data={"chat_id": config.TELEGRAM_CHAT_ID, "caption": caption},
                          files={"document": (png.name, f, "image/png")}, timeout=60)
    r.raise_for_status()
    print(f"[telegram] sent {png.name}")


def send_all(items: list[tuple[str, Path, Path]], stamp: str) -> None:
    for key, png, _pdf in items:
        send_photo(png, f"{key} Long Outage RCA - {stamp}")
