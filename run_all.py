"""Hourly entry point: render -> Telegram (PNG) -> WhatsApp (PDF).

A lock file makes sure an hourly run is skipped if the previous one is still going,
so a stuck Excel or browser never stacks up multiple runs.
"""
from __future__ import annotations

import os
import sys
import time
from datetime import datetime

import config
import render_snapshot
import send_telegram
import send_whatsapp

LOCK = config.BASE_DIR / "run.lock"
STALE_AFTER_SEC = 30 * 60


def acquire_lock() -> bool:
    if LOCK.exists() and time.time() - LOCK.stat().st_mtime < STALE_AFTER_SEC:
        return False
    LOCK.write_text(str(os.getpid()))
    return True


def main() -> int:
    if not acquire_lock():
        print("Previous run still active - skipping this one.")
        return 0
    try:
        if not config.WORKBOOK.exists():
            import generate_sample_data
            generate_sample_data.main()

        stamp = datetime.now().strftime("%d %b %Y, %H:%M")
        items = render_snapshot.render_all()

        for step, fn in (("telegram", send_telegram.send_all),
                         ("whatsapp", send_whatsapp.send_all)):
            try:
                fn(items, stamp)
            except Exception as exc:          # one channel failing must not block the other
                print(f"[{step}] FAILED: {exc}", file=sys.stderr)
        print("Done.")
        return 0
    finally:
        LOCK.unlink(missing_ok=True)


if __name__ == "__main__":
    sys.exit(main())
