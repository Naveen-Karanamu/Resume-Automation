"""
End-to-end: Copy Job Description  ->  ATS Check  ->  Copy Missing Keywords
==========================================================================
Part A — Apply Jobs (firststepjob.com):
  1. Open Apply Jobs (reusing saved FirstStep session cookies)
  2. Enter filters: Created By = Job Collector, Date = yesterday,
     Category = Data Scientist
  3. Fetch Jobs
  4. Open the Resume modal on the first job
  5. Click the "Copy job description" icon -> capture the JD

Part B — Resume Builder (builder.firststepjob.com):
  6. Open the Resume Builder and LOG IN with credentials.txt
  7. Click "ATS Check" at the bottom
  8. Paste the Job Description into the textarea
  9. Click "FirstStep ATS Check" and wait for the analysis
 10. Scroll down and click the copy icon beside "Missing Keywords"
 11. Capture the keywords from the clipboard

Writes: job_description.txt , missing_keywords.txt

Run:
    python pipeline_jd_ats.py
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
APPLY_URL    = "https://firststepjob.com/application-incharge/apply-jobs?userId=google-oauth2|110590764416739071400"
BUILDER_URL  = "https://builder.firststepjob.com/resume-builder"
CHROME_EXE   = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
DEBUG_PORT   = 9222
HOME         = Path(os.path.expanduser("~"))
TEMP_DIR     = str(HOME / "AppData" / "Local" / "Temp" / "ChromeDebugSession")
SESSION_FILE = Path(__file__).parent / "session.json"
CREDS_FILE   = Path(__file__).parent / "credentials.txt"
JD_OUTFILE   = Path(__file__).parent / "job_description.txt"
KW_OUTFILE   = Path(__file__).parent / "missing_keywords.txt"

CREATED_BY = "Job Collector"
CATEGORY   = "Data Scientist"
# ─────────────────────────────────────────────────────────────────────────────


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  Infrastructure                                                            ║
# ╚══════════════════════════════════════════════════════════════════════════╝

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


def load_firststep_cookies(context):
    """Inject only the firststepjob.com cookies (avoids Google SetSID redirect)."""
    if not SESSION_FILE.exists():
        print("[!] session.json not found.")
        return
    try:
        with open(SESSION_FILE, "r") as f:
            storage = json.load(f)
        all_cookies = storage.get("cookies", [])
        cookies = [c for c in all_cookies if "firststepjob.com" in c.get("domain", "")]
        if cookies:
            context.add_cookies(cookies)
            print(f"[+] Loaded {len(cookies)} FirstStep cookie(s) "
                  f"(filtered from {len(all_cookies)}).")
    except Exception as e:
        print(f"[!] Could not load session cookies: {e}")


def read_credentials():
    creds = {}
    with open(CREDS_FILE, "r") as f:
        for line in f:
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                k, _, v = line.partition("=")
                creds[k.strip()] = v.strip()
    return creds.get("email", ""), creds.get("password", "")


def read_clipboard(page):
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


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  Part A — Apply Jobs: copy the Job Description                              ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def ensure_on_apply_page(page, settle=2.5):
    for _ in range(6):
        if "apply-jobs" not in page.url:
            print(f"  [~] Off Apply Jobs ({page.url[:60]}...). Returning...")
            page.goto(APPLY_URL, wait_until="domcontentloaded", timeout=60000)
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
    for attempt in range(2):
        try:
            ensure_on_apply_page(page, settle=0.5)
            open_locator().click()
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
            print(f"  [~] {label} attempt {attempt + 1} failed: {str(e)[:80]}")
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            time.sleep(1)
    print(f"  [!] {label} could not be set to '{option_text}'.")
    return False


def setup_filters(page):
    yesterday = datetime.today() - timedelta(days=1)
    mm, dd, yyyy = yesterday.strftime("%m"), yesterday.strftime("%d"), yesterday.strftime("%Y")
    print(f"[*] Entering filter details (date: {yesterday.strftime('%m/%d/%Y')})...")

    pick_dropdown(page, lambda: page.locator('[role="combobox"]').nth(0),
                  CREATED_BY, label="Created By")

    try:
        ensure_on_apply_page(page, settle=0.5)
        page.locator('[role="spinbutton"][aria-label="Month"]').first.click()
        time.sleep(0.2)
        page.keyboard.type(mm); time.sleep(0.2)
        page.keyboard.type(dd); time.sleep(0.2)
        page.keyboard.type(yyyy); time.sleep(0.3)
        page.keyboard.press("Escape")
        print(f"  [+] Date = {yesterday.strftime('%m/%d/%Y')}.")
    except Exception as e:
        print(f"  [!] Date failed: {str(e)[:80]}")

    pick_dropdown(
        page,
        lambda: page.locator(
            '.MuiFormControl-root:has(label:has-text("Category")) [role="combobox"]'
        ).first,
        CATEGORY, label="Category",
    )

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
    try:
        ensure_on_apply_page(page, settle=0.5)
        rows = page.locator('table tbody tr').all()
        if not rows:
            print("  [!] No job rows found.")
            return ""
        if row_index >= len(rows):
            row_index = 0
        row = rows[row_index]
        print(f"  [*] Using job row {row_index + 1} of {len(rows)}.")

        page.evaluate("document.querySelector('table')?.parentElement?.scrollBy(600, 0)")
        time.sleep(0.3)

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

        try:
            page.wait_for_selector('[role="dialog"]', timeout=8000)
        except PlaywrightTimeoutError:
            pass
        time.sleep(0.5)

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
            jd = read_clipboard(page)
        except Exception as e:
            print(f"  [!] Could not click the copy icon: {str(e)[:120]}")

        if jd:
            print(f"  [+] Job Description copied ({len(jd)} chars).")
        else:
            print("  [!] Clipboard empty after clicking copy icon.")
        return jd or ""
    except Exception as e:
        print(f"  [!] open_resume_and_copy_jd failed: {e}")
        return ""


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  Part B — Resume Builder: ATS Check & copy Missing Keywords                ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def builder_login_if_needed(page, email, password):
    """Open the builder; if the login form shows, sign in with credentials."""
    page.goto(BUILDER_URL, wait_until="domcontentloaded", timeout=60000)
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PlaywrightTimeoutError:
        pass
    time.sleep(2)

    # Detect the login form (Auth0-style #email / #password or generic inputs)
    email_field = page.locator('#email, input[type="email"], input[name="email"]').first
    pwd_field   = page.locator('#password, input[type="password"], input[name="password"]').first

    if email_field.count() > 0 and pwd_field.count() > 0 and email_field.is_visible():
        print("  [*] Login form detected — signing in with credentials.txt...")
        email_field.fill(email)
        pwd_field.fill(password)
        page.locator(
            'button[type="submit"], button:has-text("Log in"), '
            'button:has-text("Login"), button:has-text("Sign in"), '
            'button:has-text("Continue")'
        ).first.click()
        print("  [*] Submitted login...")
        try:
            page.wait_for_load_state("networkidle", timeout=25000)
        except PlaywrightTimeoutError:
            pass
        time.sleep(4)
    else:
        print("  [+] Already logged in to the builder (no login form).")

    print(f"  [+] Builder URL: {page.url}")


def ats_check_and_copy_keywords(page, jd):
    """Click ATS Check, paste JD, run FirstStep ATS Check, copy Missing Keywords."""
    # ── Step 7: ATS Check button at the bottom ───────────────────────────
    # Use an exact-text match to avoid also matching "FirstStep ATS Check".
    try:
        ats_btn = page.locator('button:text-is("ATS Check")').first
        if ats_btn.count() == 0:
            ats_btn = page.locator('button:has-text("ATS Check")').first
        ats_btn.wait_for(state="visible", timeout=25000)
        ats_btn.scroll_into_view_if_needed()
        ats_btn.click()
        print("  [+] Clicked 'ATS Check' — modal opening...")
    except Exception as e:
        print(f"  [!] Could not click ATS Check: {str(e)[:120]}")
        return ""

    # Wait for the "ATS Score Checker" modal, identified by its heading.
    try:
        page.wait_for_selector(
            'text=/Paste Job Description/i, [role="dialog"], .MuiDialog-root',
            timeout=15000,
        )
    except PlaywrightTimeoutError:
        pass
    time.sleep(1.0)

    # ── Step 8: paste the JD into the modal's "Paste Job Description" box ──
    # Target the textarea by its placeholder so we don't grab the builder's
    # own hidden form textareas (e.g. the resume "summary" field).
    try:
        ta = page.locator(
            'textarea[placeholder*="Paste the complete job description" i], '
            'textarea[placeholder*="job description" i]'
        ).first
        if ta.count() == 0:
            # Fall back to a visible textarea inside the dialog.
            ta = page.locator('[role="dialog"] textarea, .MuiDialog-root textarea').first
        ta.wait_for(state="visible", timeout=10000)
        ta.click()
        ta.fill(jd)
        print(f"  [+] Pasted the Job Description ({len(jd)} chars).")
        time.sleep(0.5)
    except Exception as e:
        print(f"  [!] Could not paste JD: {str(e)[:120]}")
        return ""

    # ── Step 9: click "FirstStep ATS Check" ──────────────────────────────
    try:
        run_btn = page.locator('button:has-text("FirstStep ATS Check")').first
        run_btn.wait_for(state="visible", timeout=10000)
        run_btn.click()
        print("  [+] Clicked 'FirstStep ATS Check' — analyzing...")
    except Exception as e:
        print(f"  [!] Could not click FirstStep ATS Check: {str(e)[:120]}")
        return ""

    try:
        page.wait_for_selector('text=/Missing Keywords/i', timeout=90000)
        print("  [+] Analysis complete — 'Missing Keywords' found.")
        time.sleep(1)
    except PlaywrightTimeoutError:
        print("  [!] 'Missing Keywords' did not appear within 90s.")

    # ── Step 10/11: scroll to Missing Keywords & click its copy icon ─────
    keywords = ""
    try:
        heading = page.locator('text=/Missing Keywords/i').first
        heading.scroll_into_view_if_needed()
        time.sleep(0.5)

        copy_btn = page.locator(
            '[aria-label*="Copy missing" i], '
            '[title*="Copy missing" i], '
            '[aria-label*="missing keyword" i], '
            '[title*="missing keyword" i]'
        ).first

        if copy_btn.count() == 0:
            # Fall back to the first button right after the heading text.
            copy_btn = page.locator(
                'xpath=//*[contains(translate(text(),"MISSING","missing"),"missing keywords")]'
                '/following::button[1]'
            ).first

        copy_btn.wait_for(state="visible", timeout=10000)
        copy_btn.scroll_into_view_if_needed()
        copy_btn.click()
        print("  [+] Clicked the copy icon beside 'Missing Keywords'.")
        time.sleep(0.6)
        keywords = read_clipboard(page)
    except Exception as e:
        print(f"  [!] Could not click Missing Keywords copy icon: {str(e)[:120]}")

    return keywords or ""


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  Main                                                                      ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def main():
    print("=" * 64)
    print("  FirstStep — Copy JD  ->  ATS Check  ->  Copy Missing Keywords")
    print("=" * 64)

    email, password = read_credentials()
    print(f"[*] Builder credentials loaded for: {email}")

    chrome_proc = launch_chrome()

    with sync_playwright() as p:
        print("[*] Connecting to Chrome over CDP...")
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{DEBUG_PORT}")
        context = browser.contexts[0]

        for origin in ("https://firststepjob.com", "https://builder.firststepjob.com"):
            try:
                context.grant_permissions(
                    ["clipboard-read", "clipboard-write"], origin=origin
                )
            except Exception:
                pass
        print("[+] Clipboard permissions granted.")

        load_firststep_cookies(context)

        # ── PART A ────────────────────────────────────────────────────────
        print("\n--- PART A: Copy Job Description ---")
        apply_page = context.pages[0] if context.pages else context.new_page()
        apply_page.goto(APPLY_URL, wait_until="domcontentloaded", timeout=60000)
        try:
            apply_page.wait_for_load_state("networkidle", timeout=10000)
        except PlaywrightTimeoutError:
            pass

        if not ensure_on_apply_page(apply_page, settle=2.5):
            print("[!] Apply Jobs filters didn't load. URL:", apply_page.url)
            browser.close(); chrome_proc.terminate(); return
        print("[+] Apply Jobs page ready.")

        setup_filters(apply_page)
        jd = open_resume_and_copy_jd(apply_page, row_index=0)

        if not jd:
            print("[!] No Job Description captured — stopping before ATS check.")
            browser.close(); chrome_proc.terminate(); return

        JD_OUTFILE.write_text(jd, encoding="utf-8")
        print(f"[+] Job Description saved to: {JD_OUTFILE.name} ({len(jd)} chars).")

        # ── PART B ────────────────────────────────────────────────────────
        print("\n--- PART B: ATS Check & Missing Keywords ---")
        builder_page = context.new_page()
        builder_login_if_needed(builder_page, email, password)
        keywords = ats_check_and_copy_keywords(builder_page, jd)

        if keywords:
            KW_OUTFILE.write_text(keywords, encoding="utf-8")
            print(f"\n[+] Missing keywords copied ({len(keywords)} chars).")
            print(f"[+] Saved to: {KW_OUTFILE}")
            print("\n" + "─" * 64)
            print("MISSING KEYWORDS")
            print("─" * 64)
            print(keywords)
            print("─" * 64)
        else:
            print("\n[!] No missing keywords were captured.")

        time.sleep(2)
        browser.close()
        chrome_proc.terminate()
        print("\n[+] Done.")


if __name__ == "__main__":
    main()
