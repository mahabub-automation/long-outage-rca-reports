"""Send the PDF reports to a WhatsApp group through WhatsApp Web (Playwright).

Why PDF and not PNG: WhatsApp compresses photos so hard that a wide table becomes
unreadable. A PDF arrives untouched.

Lessons baked in:
  * Attaching through the paperclip menu is unreliable - the page only exposes an
    image-only <input type=file>, which silently drops PDFs. Instead the file is put on
    the OS clipboard and pasted into the chat with Ctrl+V (Windows: PowerShell
    Set-Clipboard).
  * After pasting, a preview opens with its own caption box. Both boxes are labelled
    "Type a message"; the chat box adds " to <name>". That suffix is the only reliable
    way to tell them apart.
  * A persistent browser profile keeps the QR login between runs.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import config


def copy_file_to_clipboard(path: Path) -> None:
    if sys.platform != "win32":
        raise RuntimeError("File-to-clipboard paste is implemented for Windows only")
    subprocess.run(
        ["powershell", "-NoProfile", "-Command", f"Set-Clipboard -Path '{path}'"],
        check=True,
    )


def send_all(items: list[tuple[str, Path, Path]], stamp: str) -> None:
    if config.DRY_RUN:
        for key, _png, pdf in items:
            print(f"[whatsapp:dry-run] would send {pdf.name} to '{config.WHATSAPP_GROUP}' "
                  f"| caption: {key} Long Outage RCA - {stamp}")
        return

    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(str(config.WHATSAPP_PROFILE_DIR),
                                                   headless=False)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto("https://web.whatsapp.com")
        page.wait_for_selector("#pane-side", timeout=120_000)   # logged in

        # Open the group chat through the search box.
        search = page.get_by_role("textbox").first
        search.fill(config.WHATSAPP_GROUP)
        page.get_by_title(config.WHATSAPP_GROUP, exact=True).first.click()
        chat_box = page.locator('div[contenteditable="true"][aria-label^="Type a message to"]')
        chat_box.wait_for()

        for key, _png, pdf in items:
            copy_file_to_clipboard(pdf)
            chat_box.click()
            page.keyboard.press("Control+V")
            # The caption box is the "Type a message" box WITHOUT the " to <name>" suffix.
            caption = page.locator('div[contenteditable="true"][aria-label="Type a message"]')
            caption.wait_for(timeout=30_000)
            caption.fill(f"{key} Long Outage RCA - {stamp}")
            page.keyboard.press("Enter")
            page.wait_for_timeout(4_000)
            print(f"[whatsapp] sent {pdf.name}")

        ctx.close()
