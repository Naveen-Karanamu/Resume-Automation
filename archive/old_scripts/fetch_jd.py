"""
Apply Jobs -> Fetch -> Open Resume -> Copy Job Description
==========================================================
Focused automation for exactly these steps:

  1. Launch Chrome (reusing saved login session from session.json)
  2. Open the Apply Jobs page
  3. Enter the filter details:
       - Created By = "Job Collector"
       - Date       = yesterday
       - Category    = "Data Scientist"
  4. Click "Fetch Jobs"
  5. Click the "Resume" (Open Resume) button on the first job row
  6. Copy the Job Description from the modal

The copied Job Description is printed to the console and written to
job_description.txt in this folder.

Run:
    python fetch_jd.py
"""

import json
import os
import subprocess
import time
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ── CONFIG ──────────────────────────────────────────────────────────────────
URL          = "https://firststepjob.com/application-incharge/apply-jobs?userId=google-oauth2|110590764416739071400"
CHROME_EXE   = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
DEBUG_PORT   = 9222
HOME         = Path(os.path.expanduser("~"))
TEMP_DIR     = str(HOME / "AppData" / "Local" / "Temp" / "ChromeDebugSession")
SESSION_FILE = Path(__file__).parent / "session.json"
JD_OUTFILE   = Path(__file__).parent / "job_description.txt"

CREATED_BY = "Job Collector"
CATEGORY   = "Data Scientist"
# ─────────────────────────────────────────────────────────────────────────────


def launch_chrome():
    """Launch system Chrome with remote debugging and wait for the debug port."""
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
        try:
            urllib.request.urlopen(f"http://127.0.0.1:{DEBUG_PORT}/json", timeout=2)
            print(f"[+] Chrome ready (PID {proc.pid}).")
            return proc
        except Exception:
            time.sleep(1)
    raise RuntimeError("Chrome debug port did not come up.")


def load_saved_cookies(context):
    """Inject cookies from session.json so we are already logged in."""
    if not SESSION_FILE.exists():
        print("[!] session.json not found — you may need to log in manually.")
        return
    try:
        with open(SESSION_FILE, "r") as f:
            storage = json.load(f)
        all_cookies = storage.get("cookies", [])

        # Only inject cookies for the FirstStep domains. Injecting Google's
        # auth cookies (SID/LSID/etc.) makes the browser perform a cross-domain
        # SetSID session-sync redirect that hijacks the tab mid-interaction.
        keep_domains = ("firststepjob.com",)
        cookies = [
            c for c in all_cookies
            if any(d in c.get("domain", "") for d in keep_domains)
        ]
        if cookies:
            context.add_cookies(cookies)
            print(f"[+] Loaded {len(cookies)} FirstStep cookie(s) "
                  f"(filtered from {len(all_cookies)}).")
        else:
            print("[!] No FirstStep cookies found in session.json — "
                  "may need manual login.")
    except Exception as e:
        print(f"[!] Could not load session cookies: {e}")


def ensure_on_apply_page(page, settle=2.5):
    """
    Make sure we are on the Apply Jobs page with the filter controls present.
    Auth cookies can trigger a Google SetSID redirect that bounces the tab away;
    if that happened, navigate back and wait for things to settle.
    """
    for attempt in range(6):
        url = page.url
        if "apply-jobs" not in url:
            # An auth redirect (e.g. accounts.google SetSID) pulled us away.
            print(f"  [~] Off the Apply Jobs page ({url[:60]}...). Returning...")
            page.goto(URL, wait_until="domcontentloaded", timeout=60000)
        try:
            page.wait_for_load_state("networkidle", timeout=8000)
        except PlaywrightTimeoutError:
            pass
        try:
            page.wait_for_selector('[role="combobox"]', timeout=8000)
            time.sleep(settle)
            return True
        except PlaywrightTimeoutError:
            time.sleep(1)
    return False


def pick_dropdown(page, open_locator, option_text, label=""):
    """
    Robustly select an option from a MUI Select dropdown:
      1. ensure the page is settled (no pending auth redirect)
      2. open the combobox
      3. wait for the option listbox to render
      4. click the matching option
    Retries once if a navigation interrupts the interaction.
    """
    for attempt in range(2):
        try:
            ensure_on_apply_page(page, settle=0.5)
            open_locator().click()
            # Wait for the MUI listbox popover to appear
            page.wait_for_selector('[role="listbox"], [role="option"]', timeout=6000)
            time.sleep(0.3)
            option = page.locator(f'[role="option"]:has-text("{option_text}")').first
            option.wait_for(state="visible", timeout=6000)
            option.click()
            page.keyboard.press("Escape")
            page.locator("body").click(position={"x": 10, "y": 10}, force=True)
            time.sleep(0.3)
            print(f"  [+] {label or 'Dropdown'} = {option_text}.")
            return True
        except Exception as e:
            print(f"  [~] {label or 'Dropdown'} attempt {attempt + 1} failed: {str(e)[:80]}")
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            time.sleep(1)
    print(f"  [!] {label or 'Dropdown'} could not be set to '{option_text}'.")
    return False


def setup_filters(page):
    """Fill Created By, Date, Category and click Fetch Jobs."""
    yesterday = datetime.today() - timedelta(days=1)
    yesterday_str = yesterday.strftime("%m/%d/%Y")
    mm   = yesterday.strftime("%m")
    dd   = yesterday.strftime("%d")
    yyyy = yesterday.strftime("%Y")
    print(f"[*] Entering filter details (date: {yesterday_str})...")

    # ── Created By (first combobox) ──────────────────────────────────────
    pick_dropdown(
        page,
        lambda: page.locator('[role="combobox"]').nth(0),
        CREATED_BY,
        label="Created By",
    )

    # ── Date (MUI date picker with month/day/year spinbuttons) ───────────
    try:
        ensure_on_apply_page(page, settle=0.5)
        month_span = page.locator('[role="spinbutton"][aria-label="Month"]').first
        month_span.click()
        time.sleep(0.2)
        page.keyboard.type(mm)
        time.sleep(0.2)
        page.keyboard.type(dd)
        time.sleep(0.2)
        page.keyboard.type(yyyy)
        time.sleep(0.3)
        page.keyboard.press("Escape")
        print(f"  [+] Date = {yesterday_str}.")
    except Exception as e:
        print(f"  [!] Date failed: {str(e)[:80]}")

    # ── Category (combobox next to the "Category" label) ─────────────────
    pick_dropdown(
        page,
        lambda: page.locator(
            '.MuiFormControl-root:has(label:has-text("Category")) [role="combobox"]'
        ).first,
        CATEGORY,
        label="Category",
    )

    # ── Fetch Jobs ───────────────────────────────────────────────────────
    try:
        btn = page.locator('button:has-text("Fetch Jobs")').first
        for _ in range(20):
            if btn.is_enabled():
                break
            time.sleep(0.5)
        btn.scroll_into_view_if_needed()
        btn.click()
        print("  [+] Fetch Jobs clicked. Waiting for results...")
        time.sleep(4)
    except Exception as e:
        print(f"  [!] Fetch Jobs failed: {e}")

    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    time.sleep(1)


def open_resume_and_copy_jd(page, row_index=0):
    """
    Click the "Resume" (Open Resume) button on the given job row, then copy
    the Job Description text from the modal that opens.
    Returns the JD string (or "" on failure).
    """
    try:
        ensure_on_apply_page(page, settle=0.5)
        rows = page.locator('table tbody tr').all()
        if not rows:
            print("  [!] No job rows found in the results table.")
            return ""
        if row_index >= len(rows):
            row_index = 0
        row = rows[row_index]
        print(f"  [*] Using job row {row_index + 1} of {len(rows)}.")

        # Scroll the table horizontally to reveal the action buttons
        page.evaluate("document.querySelector('table')?.parentElement?.scrollBy(600, 0)")
        time.sleep(0.3)

        # Click the exact "Resume" button (NOT "Tailor Resume")
        resume_btn = row.locator('button:text-is("Resume")').first
        if resume_btn.count() == 0:
            for btn in row.locator('button').all():
                if btn.inner_text().strip() == "Resume":
                    resume_btn = btn
                    break
        resume_btn.scroll_into_view_if_needed()
        resume_btn.click()
        print("  [*] Clicked Resume — waiting for the modal...")
        time.sleep(2)

        # Wait for the modal dialog to be present
        try:
            page.wait_for_selector('[role="dialog"]', timeout=8000)
        except PlaywrightTimeoutError:
            pass
        time.sleep(0.5)

        # Click the "Copy job description" icon button next to the
        # "Job Description" heading, then read what it put on the clipboard.
        jd = ""
        try:
            copy_jd_btn = page.locator(
                '[aria-label="Copy job description"], '
                '[title="Copy job description"], '
                '[aria-label*="Copy job description" i], '
                '[title*="Copy job description" i]'
            ).first
            copy_jd_btn.wait_for(state="visible", timeout=8000)
            copy_jd_btn.scroll_into_view_if_needed()
            copy_jd_btn.click()
            print("  [+] Clicked the 'Copy job description' icon.")
            time.sleep(0.6)
            try:
                jd = page.evaluate("navigator.clipboard.readText()")
            except Exception:
                # Clipboard read was denied — paste into a temporary textarea
                # to recover what the copy icon placed on the clipboard.
                print("  [~] Direct clipboard read denied; recovering via paste...")
                jd = page.evaluate("""async () => {
                    const ta = document.createElement('textarea');
                    document.body.appendChild(ta);
                    ta.focus();
                    document.execCommand('paste');
                    const val = ta.value;
                    ta.remove();
                    return val;
                }""")
        except Exception as e:
            print(f"  [!] Could not click the copy icon: {str(e)[:120]}")

        if jd:
            print(f"  [+] Job Description copied to clipboard ({len(jd)} chars).")
        else:
            print("  [!] Clipboard was empty after clicking the copy icon.")
        return jd or ""

    except Exception as e:
        print(f"  [!] open_resume_and_copy_jd failed: {e}")
        return ""


def main():
    print("=" * 60)
    print("  FirstStep — Fetch Jobs & Copy Job Description")
    print("=" * 60)

    chrome_proc = launch_chrome()

    with sync_playwright() as p:
        print("[*] Connecting to Chrome over CDP...")
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{DEBUG_PORT}")
        context = browser.contexts[0]

        # Grant clipboard read/write so we can read back what the
        # "Copy job description" icon puts on the clipboard.
        try:
            context.grant_permissions(
                ["clipboard-read", "clipboard-write"],
                origin="https://firststepjob.com",
            )
            print("[+] Clipboard permissions granted.")
        except Exception as e:
            print(f"[~] Could not grant clipboard permissions: {str(e)[:80]}")

        # Inject saved login cookies before navigating
        load_saved_cookies(context)

        page = context.pages[0] if context.pages else context.new_page()

        print("[*] Opening the Apply Jobs page...")
        page.goto(URL, wait_until="domcontentloaded", timeout=60000)
        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except PlaywrightTimeoutError:
            pass

        # If we got bounced to a login page, the saved session may be stale
        if "login" in page.url.lower():
            print("\n[!] A login page appeared — the saved session may be expired.")
            print("[!] Log in with Google in the Chrome window.")
            input("[*] Press Enter here once you reach the Apply Jobs page...")

        if ensure_on_apply_page(page, settle=2.5):
            print("[+] Apply Jobs page is ready.")
        else:
            print("[!] Filters didn't appear in time. Current URL:", page.url)
            input("[*] Press Enter to close...")
            browser.close()
            chrome_proc.terminate()
            return

        # Steps 3-4: enter details + Fetch Jobs
        setup_filters(page)

        # Steps 5-6: open resume + copy JD (first job)
        jd = open_resume_and_copy_jd(page, row_index=0)

        if jd:
            JD_OUTFILE.write_text(jd, encoding="utf-8")
            print(f"\n[+] Job Description saved to: {JD_OUTFILE}")
            print("\n" + "─" * 60)
            print("JOB DESCRIPTION")
            print("─" * 60)
            print(jd)
            print("─" * 60)
        else:
            print("\n[!] No Job Description was captured.")

        print("\n[*] Press Enter to close the browser...")
        input()
        browser.close()
        chrome_proc.terminate()


if __name__ == "__main__":
    main()
