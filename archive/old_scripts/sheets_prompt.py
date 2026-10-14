"""
Google Sheets -> paste Missing Keywords + JD -> copy the ChatGPT prompt (C3:C9)
===============================================================================
Steps:
  1. Open the Google Sheet (reusing the running Chrome / Google session)
  2. Paste the Missing Keywords into A9
  3. Paste the Job Description (copied earlier) into B3
  4. Select C3:C9, copy it, and read the generated ChatGPT prompt

Reads : missing_keywords.txt , job_description.txt
Writes: chatgpt_prompt.txt

Assumes Chrome is already running with --remote-debugging-port=9222
(launched by the earlier steps). If not, run fetch_jd.py / pipeline first,
or this script will launch Chrome itself.

Run:
    python sheets_prompt.py
"""

import json
import os
import subprocess
import time
import urllib.request
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ── CONFIG ──────────────────────────────────────────────────────────────────
SHEETS_URL   = "https://docs.google.com/spreadsheets/d/1BsXaJmFELvvB8s-bRogF9jNhiKgxvx-eUVcpBW-B6Ao/edit?gid=0#gid=0"
CHROME_EXE   = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
DEBUG_PORT   = 9222
HOME         = Path(os.path.expanduser("~"))
TEMP_DIR     = str(HOME / "AppData" / "Local" / "Temp" / "ChromeDebugSession")
SESSION_FILE = Path(__file__).parent / "session.json"
JD_INFILE    = Path(__file__).parent / "job_description.txt"
KW_INFILE    = Path(__file__).parent / "missing_keywords.txt"
PROMPT_OUT   = Path(__file__).parent / "chatgpt_prompt.txt"

KEYWORDS_CELL = "A9"
JD_CELL       = "B3"
PROMPT_RANGE  = "C3:C9"
# ─────────────────────────────────────────────────────────────────────────────


def debug_port_up():
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{DEBUG_PORT}/json", timeout=2)
        return True
    except Exception:
        return False


def launch_chrome():
    print("[*] Launching Chrome with remote debugging...")
    Path(TEMP_DIR).mkdir(parents=True, exist_ok=True)
    proc = subprocess.Popen([
        CHROME_EXE,
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--user-data-dir={TEMP_DIR}",
        "--no-first-run",
        "--no-default-browser-check",
        "about:blank",
    ])
    for _ in range(30):
        if debug_port_up():
            print(f"[+] Chrome ready (PID {proc.pid}).")
            return proc
        time.sleep(1)
    raise RuntimeError("Chrome debug port did not come up.")


def load_all_useful_cookies(context):
    """Inject Google + FirstStep cookies so Sheets is authenticated."""
    if not SESSION_FILE.exists():
        return
    try:
        with open(SESSION_FILE, "r") as f:
            storage = json.load(f)
        all_cookies = storage.get("cookies", [])
        skip = ("cawallpaper.com", "doubleclick.net")
        cookies = [c for c in all_cookies
                   if not any(d in c.get("domain", "") for d in skip)]
        if cookies:
            context.add_cookies(cookies)
            print(f"[+] Loaded {len(cookies)} cookie(s) for Google/FirstStep.")
    except Exception as e:
        print(f"[!] Could not load cookies: {e}")


def set_clipboard(page, text):
    """Put text onto the browser clipboard via the async clipboard API."""
    page.evaluate("async (t) => { await navigator.clipboard.writeText(t); }", text)


def go_to_cell(page, ref):
    """Use the Name Box to jump to a cell/range reference."""
    nb = page.locator('input.waffle-name-box, input[aria-label="Name box"], input[aria-label="Name Box"]').first
    nb.click()
    nb.press("Control+a")
    nb.type(ref)
    page.keyboard.press("Enter")
    time.sleep(0.5)


def paste_into_cell(page, ref, value):
    """
    Put `value` on the clipboard, jump to `ref`, enter edit mode (F2) so the
    whole value (including newlines) lands inside the single cell, then paste.
    """
    set_clipboard(page, value)
    go_to_cell(page, ref)
    # Enter edit mode so Ctrl+V pastes as cell content, not as a grid of cells.
    page.keyboard.press("F2")
    time.sleep(0.3)
    page.keyboard.press("Control+v")
    time.sleep(0.6)
    page.keyboard.press("Enter")
    time.sleep(0.5)
    print(f"  [+] Pasted into {ref} ({len(value)} chars).")


def main():
    print("=" * 64)
    print("  Google Sheets — paste Keywords + JD, copy the ChatGPT prompt")
    print("=" * 64)

    if not KW_INFILE.exists() or not JD_INFILE.exists():
        print("[!] missing_keywords.txt or job_description.txt not found.")
        return
    keywords = KW_INFILE.read_text(encoding="utf-8").strip()
    jd = JD_INFILE.read_text(encoding="utf-8").strip()
    print(f"[+] Keywords ({len(keywords)} chars), JD ({len(jd)} chars) loaded.")

    chrome_proc = None
    if not debug_port_up():
        chrome_proc = launch_chrome()
    else:
        print("[+] Reusing the already-running Chrome debug session.")

    with sync_playwright() as p:
        print("[*] Connecting to Chrome over CDP...")
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{DEBUG_PORT}")
        context = browser.contexts[0]

        try:
            context.grant_permissions(
                ["clipboard-read", "clipboard-write"],
                origin="https://docs.google.com",
            )
            print("[+] Clipboard permissions granted.")
        except Exception as e:
            print(f"[~] Could not grant clipboard permissions: {str(e)[:80]}")

        # Reuse an already-open spreadsheet tab if present (it's authenticated),
        # otherwise open the sheet in a new tab.
        page = None
        for pg in context.pages:
            if "docs.google.com/spreadsheets" in pg.url and "/edit" in pg.url:
                page = pg
                print("[+] Reusing the already-open spreadsheet tab.")
                break

        if page is None:
            load_all_useful_cookies(context)
            page = context.new_page()
            print("[*] Opening the Google Sheet...")
            page.goto(SHEETS_URL, wait_until="domcontentloaded", timeout=60000)

        # Wait until the real spreadsheet editor is loaded (URL contains /edit).
        for _ in range(20):
            try:
                page.wait_for_load_state("networkidle", timeout=5000)
            except PlaywrightTimeoutError:
                pass
            if "/edit" in page.url and "accounts.google.com" not in page.url:
                break
            time.sleep(1)
        page.bring_to_front()
        time.sleep(2)

        if "accounts.google.com" in page.url:
            print("[!] Google is on a sign-in page. Sign in in the Chrome window,")
            print("[!] open the sheet, then re-run this script.")
            browser.close()
            if chrome_proc:
                chrome_proc.terminate()
            return

        # Make sure the grid is focused
        try:
            page.wait_for_selector('input.waffle-name-box, [aria-label="Name box"], [aria-label="Name Box"]', timeout=20000)
            print("[+] Sheet loaded and ready.")
        except PlaywrightTimeoutError:
            print("[!] Name Box not found — sheet may not have loaded. URL:", page.url)
            browser.close()
            if chrome_proc:
                chrome_proc.terminate()
            return

        # ── Step 2: Missing Keywords -> A9 ───────────────────────────────
        paste_into_cell(page, KEYWORDS_CELL, keywords)

        # ── Step 3: Job Description -> B3 ────────────────────────────────
        paste_into_cell(page, JD_CELL, jd)

        # ── Step 4: copy C3:C9 (the generated ChatGPT prompt) ────────────
        time.sleep(1.5)  # let the sheet formulas recalc
        go_to_cell(page, PROMPT_RANGE)
        page.keyboard.press("Control+c")
        time.sleep(0.8)
        try:
            prompt = page.evaluate("navigator.clipboard.readText()")
        except Exception:
            prompt = ""
        prompt = (prompt or "").strip()

        if prompt:
            PROMPT_OUT.write_text(prompt, encoding="utf-8")
            print(f"\n[+] ChatGPT prompt copied ({len(prompt)} chars).")
            print(f"[+] Saved to: {PROMPT_OUT}")
            print("\n" + "─" * 64)
            print("CHATGPT PROMPT (C3:C9)")
            print("─" * 64)
            print(prompt)
            print("─" * 64)
        else:
            print("\n[!] Nothing was copied from C3:C9 — the cells may be empty "
                  "or the clipboard read was blocked.")

        time.sleep(2)
        browser.close()
        if chrome_proc:
            chrome_proc.terminate()
        print("\n[+] Done.")


if __name__ == "__main__":
    main()
