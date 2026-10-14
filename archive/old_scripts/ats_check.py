"""
Resume Builder -> ATS Check -> Paste JD -> FirstStep ATS Check -> Copy Missing Keywords
=======================================================================================
Steps:
  1. Open https://builder.firststepjob.com/resume-builder (reusing saved login)
  2. Click the "ATS Check" button at the bottom
  3. Paste the Job Description (from job_description.txt) into the textarea
  4. Click "FirstStep ATS Check" and wait for the analysis
  5. Scroll down and click the copy icon beside "Missing Keywords"
  6. Read the copied keywords from the clipboard and save them

Reads : job_description.txt   (produced by fetch_jd.py)
Writes: missing_keywords.txt

Run:
    python ats_check.py
"""

import json
import os
import subprocess
import time
import urllib.request
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ── CONFIG ──────────────────────────────────────────────────────────────────
BUILDER_URL  = "https://builder.firststepjob.com/resume-builder"
CHROME_EXE   = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
DEBUG_PORT   = 9222
HOME         = Path(os.path.expanduser("~"))
TEMP_DIR     = str(HOME / "AppData" / "Local" / "Temp" / "ChromeDebugSession")
SESSION_FILE = Path(__file__).parent / "session.json"
JD_INFILE    = Path(__file__).parent / "job_description.txt"
KW_OUTFILE   = Path(__file__).parent / "missing_keywords.txt"
# ─────────────────────────────────────────────────────────────────────────────


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
        try:
            urllib.request.urlopen(f"http://127.0.0.1:{DEBUG_PORT}/json", timeout=2)
            print(f"[+] Chrome ready (PID {proc.pid}).")
            return proc
        except Exception:
            time.sleep(1)
    raise RuntimeError("Chrome debug port did not come up.")


def load_saved_cookies(context):
    """Inject only the FirstStep cookies (avoids Google SetSID redirects)."""
    if not SESSION_FILE.exists():
        print("[!] session.json not found — may need manual login.")
        return
    try:
        with open(SESSION_FILE, "r") as f:
            storage = json.load(f)
        all_cookies = storage.get("cookies", [])
        # The builder SPA authenticates via Auth0 (auth.firststepjob.com) which
        # relies on the Google OAuth session for silent auth, so we need the
        # FirstStep/Auth0 cookies AND the Google cookies here (unlike the
        # apply-jobs dropdown flow). We skip only unrelated third-party noise.
        skip_domains = ("cawallpaper.com", "doubleclick.net", "youtube.com")
        cookies = [
            c for c in all_cookies
            if not any(d in c.get("domain", "") for d in skip_domains)
        ]
        if cookies:
            context.add_cookies(cookies)
            print(f"[+] Loaded {len(cookies)} cookie(s) "
                  f"(filtered from {len(all_cookies)}).")
        else:
            print("[!] No usable cookies found — may need manual login.")
    except Exception as e:
        print(f"[!] Could not load session cookies: {e}")


def read_clipboard(page):
    """Read clipboard text; fall back to a paste-into-textarea trick if denied."""
    try:
        return page.evaluate("navigator.clipboard.readText()")
    except Exception:
        try:
            return page.evaluate("""async () => {
                const ta = document.createElement('textarea');
                document.body.appendChild(ta);
                ta.focus();
                document.execCommand('paste');
                const val = ta.value;
                ta.remove();
                return val;
            }""")
        except Exception:
            return ""


def main():
    print("=" * 60)
    print("  FirstStep — ATS Check & Copy Missing Keywords")
    print("=" * 60)

    # Load the Job Description captured earlier
    if not JD_INFILE.exists():
        print(f"[!] {JD_INFILE.name} not found. Run fetch_jd.py first.")
        return
    jd = JD_INFILE.read_text(encoding="utf-8").strip()
    if not jd:
        print("[!] job_description.txt is empty.")
        return
    print(f"[+] Loaded Job Description ({len(jd)} chars) to paste.")

    chrome_proc = launch_chrome()

    with sync_playwright() as p:
        print("[*] Connecting to Chrome over CDP...")
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{DEBUG_PORT}")
        context = browser.contexts[0]

        try:
            context.grant_permissions(
                ["clipboard-read", "clipboard-write"],
                origin="https://builder.firststepjob.com",
            )
            print("[+] Clipboard permissions granted.")
        except Exception as e:
            print(f"[~] Could not grant clipboard permissions: {str(e)[:80]}")

        load_saved_cookies(context)

        page = context.pages[0] if context.pages else context.new_page()

        print("[*] Opening the Resume Builder page...")
        page.goto(BUILDER_URL, wait_until="domcontentloaded", timeout=60000)
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except PlaywrightTimeoutError:
            pass
        time.sleep(2)

        if "login" in page.url.lower():
            print("\n[!] A login page appeared — the saved session may be expired.")
            print("[!] Log in in the Chrome window, reach the Resume Builder page.")
            input("[*] Press Enter here once the builder is loaded...")

        # ── Step 2: click "ATS Check" at the bottom ──────────────────────
        try:
            ats_btn = page.locator('button:has-text("ATS Check")').first
            ats_btn.wait_for(state="visible", timeout=20000)
            ats_btn.scroll_into_view_if_needed()
            ats_btn.click()
            print("[+] Clicked 'ATS Check' — modal opening...")
            time.sleep(1.5)
        except Exception as e:
            print(f"[!] Could not click ATS Check: {str(e)[:120]}")
            input("[*] Press Enter to close...")
            browser.close(); chrome_proc.terminate(); return

        # ── Step 3: paste the Job Description ─────────────────────────────
        try:
            ta = page.locator('textarea').first
            ta.wait_for(state="visible", timeout=10000)
            ta.click()
            ta.fill(jd)
            print(f"[+] Pasted the Job Description ({len(jd)} chars).")
            time.sleep(0.5)
        except Exception as e:
            print(f"[!] Could not paste JD: {str(e)[:120]}")
            input("[*] Press Enter to close...")
            browser.close(); chrome_proc.terminate(); return

        # ── Step 4: click "FirstStep ATS Check" ───────────────────────────
        try:
            run_btn = page.locator('button:has-text("FirstStep ATS Check")').first
            run_btn.wait_for(state="visible", timeout=10000)
            run_btn.click()
            print("[+] Clicked 'FirstStep ATS Check' — analyzing...")
        except Exception as e:
            print(f"[!] Could not click FirstStep ATS Check: {str(e)[:120]}")
            input("[*] Press Enter to close...")
            browser.close(); chrome_proc.terminate(); return

        # Wait for the "Missing Keywords" section to render
        try:
            page.wait_for_selector('text=/Missing Keywords/i', timeout=60000)
            print("[+] Analysis complete — 'Missing Keywords' section found.")
            time.sleep(1)
        except PlaywrightTimeoutError:
            print("[!] 'Missing Keywords' did not appear within 60s.")

        # ── Step 5: scroll down & click the copy icon beside Missing Keywords
        keywords = ""
        try:
            heading = page.locator('text=/Missing Keywords/i').first
            heading.scroll_into_view_if_needed()
            time.sleep(0.5)

            # The copy icon sits right next to the "Missing Keywords (NN)" heading.
            copy_btn = page.locator(
                '[aria-label*="Copy missing" i], '
                '[title*="Copy missing" i], '
                '[aria-label*="missing keyword" i], '
                '[title*="missing keyword" i]'
            ).first

            if copy_btn.count() == 0:
                # Fall back to the first button/icon that follows the heading.
                copy_btn = page.locator(
                    'xpath=//*[contains(translate(text(),"MISSING","missing"),"missing keywords")]'
                    '/following::button[1]'
                ).first

            copy_btn.wait_for(state="visible", timeout=10000)
            copy_btn.scroll_into_view_if_needed()
            copy_btn.click()
            print("[+] Clicked the copy icon beside 'Missing Keywords'.")
            time.sleep(0.6)
            keywords = read_clipboard(page)
        except Exception as e:
            print(f"[!] Could not click the Missing Keywords copy icon: {str(e)[:120]}")

        # ── Step 6: save & show the copied keywords ───────────────────────
        if keywords:
            KW_OUTFILE.write_text(keywords, encoding="utf-8")
            print(f"\n[+] Missing keywords copied ({len(keywords)} chars).")
            print(f"[+] Saved to: {KW_OUTFILE}")
            print("\n" + "─" * 60)
            print("MISSING KEYWORDS")
            print("─" * 60)
            print(keywords)
            print("─" * 60)
        else:
            print("\n[!] No missing keywords were captured from the clipboard.")

        print("\n[*] Press Enter to close the browser...")
        input()
        browser.close()
        chrome_proc.terminate()


if __name__ == "__main__":
    main()
