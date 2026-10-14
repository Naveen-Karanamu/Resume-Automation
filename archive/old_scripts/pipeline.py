"""
Full Job Application Pipeline
==============================
BEFORE RUNNING:
  1. Open Chrome manually and go to:
     https://firststepjob.com/application-incharge/apply-jobs?userId=google-oauth2|110590764416739071400
  2. Make sure you are logged in
  3. Run this script — it will connect to your open Chrome and do everything

Run:
    python3.10 pipeline.py
"""

import subprocess
import time
import re
from pathlib import Path
from datetime import datetime, timedelta

import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ── CONFIG ────────────────────────────────────────────────────────────────────
APPLY_JOBS_URL = "https://firststepjob.com/application-incharge/apply-jobs?userId=google-oauth2|110590764416739071400"
BUILDER_URL    = "https://builder.firststepjob.com/resume-builder"
SHEETS_URL     = "https://docs.google.com/spreadsheets/d/1BsXaJmFELvvB8s-bRogF9jNhiKgxvx-eUVcpBW-B6Ao/edit?gid=0#gid=0"
CHATGPT_URL    = "https://chat.openai.com"
CHROME_EXE     = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
DEBUG_PORT     = 9222
EXCEL_FILE     = Path(__file__).parent / "resume_data.xlsx"
RESUME_AUTO    = Path(__file__).parent / "resume_automation.py"
PYTHON_EXE     = r"C:\Users\karan\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0\python3.10.exe"
MAX_YOE        = 5
# ─────────────────────────────────────────────────────────────────────────────


def get_yesterday():
    return (datetime.today() - timedelta(days=1)).strftime("%m/%d/%Y")


def restart_chrome_with_debug():
    """Kill Chrome and relaunch with remote debugging so we can connect."""
    import urllib.request
    print("[*] Restarting Chrome with remote debugging enabled...")
    subprocess.run('taskkill /F /IM chrome.exe /T', shell=True, capture_output=True)
    time.sleep(2)
    # Remove lock files
    for f in [
        r"C:\Users\karan\AppData\Local\Google\Chrome\User Data\SingletonLock",
        r"C:\Users\karan\AppData\Local\Google\Chrome\User Data\SingletonSocket",
    ]:
        Path(f).unlink(missing_ok=True)

    subprocess.Popen([
        CHROME_EXE,
        f"--remote-debugging-port={DEBUG_PORT}",
        "--user-data-dir=C:\\Users\\karan\\AppData\\Local\\Google\\Chrome\\User Data",
        "--profile-directory=Profile 24",
        "--no-first-run",
        "--no-default-browser-check",
        APPLY_JOBS_URL,
    ])

    # Wait for debug port
    for _ in range(30):
        try:
            urllib.request.urlopen(f"http://127.0.0.1:{DEBUG_PORT}/json", timeout=2)
            print("[+] Chrome ready.")
            return
        except Exception:
            time.sleep(1)
    raise Exception("Chrome debug port not available after restart.")


def setup_filters(page):
    """Fill Created By, Date, Category and click Fetch Jobs."""
    yesterday = (datetime.today() - timedelta(days=1))
    yesterday_str = yesterday.strftime("%m/%d/%Y")
    mm  = yesterday.strftime("%m")  # e.g. "10"
    dd  = yesterday.strftime("%d")  # e.g. "12"
    yyyy = yesterday.strftime("%Y") # e.g. "2026"
    print(f"[*] Setting filters (date: {yesterday_str})...")

    def pick_dropdown(nth, option):
        dd_el = page.locator('[role="combobox"]').nth(nth)
        dd_el.click()
        time.sleep(0.5)
        page.locator(f'[role="option"]:has-text("{option}")').first.click()
        page.keyboard.press("Escape")
        page.locator("body").click(position={"x": 10, "y": 10}, force=True)
        time.sleep(0.3)

    # Step 1: Created By (combobox index 0)
    pick_dropdown(0, "Job Collector")
    print("  [+] Created By = Job Collector.")

    # Step 2: Date — MUI date picker with separate month/day/year spinbuttons
    # Click the month section, type month, then day, then year
    try:
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
        print(f"  [+] Date set to {yesterday_str}.")
    except Exception as e:
        print(f"  [!] Date failed: {e}")

    # Step 3: Category — target specifically by its label
    try:
        # Click the combobox that is next to the "Category" label
        page.locator('.MuiFormControl-root:has(label:has-text("Category")) [role="combobox"]').first.click()
        time.sleep(0.5)
        page.locator('[role="option"]:has-text("Data Scientist")').first.click()
        page.keyboard.press("Escape")
        page.locator("body").click(position={"x": 10, "y": 10}, force=True)
        time.sleep(0.3)
        print("  [+] Category = Data Scientist.")
    except Exception as e:
        print(f"  [!] Category failed: {e}")

    # Step 4: Fetch Jobs
    btn = page.locator('button:has-text("Fetch Jobs")').first
    for _ in range(20):
        if btn.is_enabled():
            break
        time.sleep(0.5)
    btn.click()
    print("  [+] Fetch Jobs clicked.")
    time.sleep(3)
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    time.sleep(1)


def get_qualifying_jobs(page):
    """Return all job rows with YOE <= MAX_YOE."""
    rows = page.locator('table tbody tr').all()
    jobs = []
    for i, row in enumerate(rows):
        try:
            text = row.inner_text()
            match = re.search(r'(\d+)\s*yr', text, re.IGNORECASE)
            if match and int(match.group(1)) <= MAX_YOE:
                jobs.append({"index": i, "text": text[:80], "yoe": int(match.group(1))})
        except Exception:
            pass
    print(f"[*] Found {len(jobs)} job(s) with YOE ≤ {MAX_YOE}.")
    return jobs


def process_job_row(page, job_index):
    """
    Click Resume button → modal opens with Job Description →
    click copy icon next to Job Description → get JD →
    close modal → click ATS Check on the same page.
    """
    try:
        # Scroll table right to reveal Resume button
        page.evaluate("document.querySelector('table')?.parentElement?.scrollBy(500, 0)")
        time.sleep(0.3)

        rows = page.locator('table tbody tr').all()
        row  = rows[job_index]

        # Click the dark filled "Resume" button — NOT "Tailor Resume"
        # Use exact text match to avoid matching "Tailor Resume"
        resume_btn = row.locator('button:text-is("Resume")').first
        if resume_btn.count() == 0:
            # Fallback: get all buttons with "Resume" text and pick the one without "Tailor"
            all_btns = row.locator('button').all()
            for btn in all_btns:
                txt = btn.inner_text().strip()
                if txt == "Resume":
                    resume_btn = btn
                    break
        resume_btn.scroll_into_view_if_needed()
        resume_btn.click()
        print("  [*] Clicked Resume — waiting for modal...")
        time.sleep(2)

        # Modal opened — click copy icon next to "Job Description" heading
        # From screenshot: small copy icon button to the right of "Job Description" text
        copy_jd_btn = page.locator(
            'h2:has-text("Job Description") ~ button, '
            'h3:has-text("Job Description") ~ button, '
            'div:has-text("Job Description") > button, '
            '[aria-label*="Copy job" i], '
            '[title*="Copy job" i]'
        ).first
        copy_jd_btn.wait_for(state="visible", timeout=8000)
        copy_jd_btn.click()
        time.sleep(0.5)
        jd = page.evaluate("navigator.clipboard.readText()")
        print(f"  [+] JD copied ({len(jd)} chars).")

        # Close modal
        page.locator('button:has-text("Cancel")').first.click()
        time.sleep(0.5)

        # Now click ATS Check in the bottom bar of the builder
        # But we are on apply-jobs page — navigate to builder for ATS check
        builder_page = page.context.new_page()
        builder_page.goto(BUILDER_URL, wait_until="networkidle", timeout=30000)
        time.sleep(2)

        builder_page.locator('button:has-text("ATS Check")').first.click()
        print("  [*] ATS Check modal opened.")
        time.sleep(1)

        ta = builder_page.locator('textarea').first
        ta.wait_for(state="visible", timeout=5000)
        ta.fill(jd)
        time.sleep(0.3)

        builder_page.locator('button:has-text("FirstStep ATS Check")').first.click()
        print("  [*] Running ATS check...")
        time.sleep(8)

        # Copy missing keywords
        kw_btn = builder_page.locator(
            'button[title*="missing" i], '
            '[aria-label*="missing" i], '
            '[title*="Copy missing" i]'
        ).first
        kw_btn.wait_for(state="visible", timeout=10000)
        kw_btn.click()
        time.sleep(0.5)
        keywords = builder_page.evaluate("navigator.clipboard.readText()")
        print(f"  [+] Missing keywords: {keywords[:80]}...")

        builder_page.close()
        return jd, keywords

    except Exception as e:
        print(f"  [!] process_job_row failed: {e}")
        try:
            page.keyboard.press("Escape")
            # Close builder tab if open
        except Exception:
            pass
        return None, None


def update_sheets_and_get_prompt(context, jd, keywords):
    """Paste data into Google Sheets, copy C3:C9 prompt."""
    tab = context.new_page()
    try:
        tab.goto(SHEETS_URL, wait_until="networkidle", timeout=30000)
        time.sleep(3)

        def go_to_cell(ref, value):
            nb = tab.locator('input[aria-label="Name Box"]').first
            nb.click()
            nb.select_text()
            nb.type(ref)
            tab.keyboard.press("Enter")
            time.sleep(0.3)
            tab.keyboard.press("Delete")
            tab.keyboard.type(value, delay=10)
            tab.keyboard.press("Enter")
            time.sleep(0.3)

        go_to_cell("A9", keywords)
        print("  [+] A9 updated.")
        go_to_cell("B3", jd)
        print("  [+] B3 updated.")

        # Copy C3:C9
        nb = tab.locator('input[aria-label="Name Box"]').first
        nb.click()
        nb.select_text()
        nb.type("C3:C9")
        tab.keyboard.press("Enter")
        time.sleep(0.5)
        tab.keyboard.press("Control+c")
        time.sleep(0.5)
        prompt = tab.evaluate("navigator.clipboard.readText()")
        print(f"  [+] Prompt copied ({len(prompt)} chars).")
        return prompt
    except Exception as e:
        print(f"  [!] Sheets update failed: {e}")
        return ""
    finally:
        tab.close()


def get_chatgpt_response(context, prompt):
    """Send prompt to ChatGPT, extract Summary + Ex Desc 1 + Ex Desc 2."""
    tab = context.new_page()
    try:
        tab.goto(CHATGPT_URL, wait_until="networkidle", timeout=30000)
        time.sleep(3)

        box = tab.locator('#prompt-textarea, [contenteditable="true"]').first
        box.wait_for(state="visible", timeout=10000)
        box.click()
        box.fill(prompt)
        tab.keyboard.press("Enter")
        print("  [*] Prompt sent to ChatGPT. Waiting for response...")

        # Wait for response to finish
        for _ in range(90):
            if tab.locator('button[aria-label*="Stop" i]').count() == 0:
                break
            time.sleep(1)
        time.sleep(2)

        # Get last response
        msgs = tab.locator('[data-message-author-role="assistant"]').all()
        if not msgs:
            return "", "", ""
        text = msgs[-1].inner_text()

        def extract(headings):
            for h in headings:
                m = re.search(rf'{re.escape(h)}[:\-\s]*(.*?)(?=\n[A-Z]{{2,}}|\Z)', text, re.IGNORECASE | re.DOTALL)
                if m:
                    return m.group(1).strip()
            return ""

        summary   = extract(["PROFESSIONAL SUMMARY", "SUMMARY"])
        ex_desc_1 = extract(["Responsibilities 1", "RESPONSIBILITIES 1"])
        ex_desc_2 = extract(["Responsibilities 2", "RESPONSIBILITIES 2"])
        print(f"  [+] Got Summary, Ex Desc 1, Ex Desc 2 from ChatGPT.")
        return summary, ex_desc_1, ex_desc_2
    except Exception as e:
        print(f"  [!] ChatGPT failed: {e}")
        return "", "", ""
    finally:
        tab.close()


def append_to_excel(summary, ex_desc_1, ex_desc_2):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append(["", "", summary, ex_desc_1, ex_desc_2, ""])
    wb.save(EXCEL_FILE)
    print("  [+] Row appended to resume_data.xlsx.")


def main():
    print("=" * 60)
    print("  FirstStep — Full Pipeline")
    print("=" * 60)
    print("\n[!] Make sure Chrome is open at the Apply Jobs page.")
    print("[!] If Chrome is NOT open, the script will open it for you.\n")

    with sync_playwright() as p:
        # Connect to Chrome that was opened with --remote-debugging-port=9222
        try:
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            context = browser.contexts[0]
            page = context.pages[0] if context.pages else context.new_page()
            print("[+] Connected to Chrome.")
        except Exception:
            print("\n[!] Could not connect to Chrome.")
            print("[!] Please open Chrome using this command first:")
            print()
            print('    Stop-Process -Name chrome -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 2; & "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\\Users\\karan\\AppData\\Local\\Google\\Chrome\\User Data" --profile-directory="Profile 24" "https://firststepjob.com/application-incharge/apply-jobs?userId=google-oauth2|110590764416739071400"')
            print()
            return

        # Navigate to Apply Jobs if not already there
        if "apply-jobs" not in page.url:
            print("[*] Navigating to Apply Jobs...")
            page.goto(APPLY_JOBS_URL, wait_until="domcontentloaded", timeout=60000)
            time.sleep(5)

        page.wait_for_selector('[role="combobox"]', timeout=30000)
        print(f"[+] On Apply Jobs page.")

        # Fill filters and fetch jobs
        setup_filters(page)

        # Get qualifying jobs
        jobs = get_qualifying_jobs(page)
        if not jobs:
            print("[!] No qualifying jobs found. Exiting.")
            return

        print(f"\n[*] Processing {len(jobs)} job(s)...\n")

        processed = 0
        for job in jobs:
            print(f"\n{'─'*50}")
            print(f"  Job {job['index']+1} | YOE: {job['yoe']} | {job['text'][:60]}")
            print(f"{'─'*50}")

            # Step 1: Copy JD + ATS check on generate-resume page
            jd, keywords = process_job_row(page, job["index"])
            if not jd:
                page.goto(APPLY_JOBS_URL, wait_until="domcontentloaded", timeout=30000)
                time.sleep(3)
                continue

            # Step 2: Google Sheets → get ChatGPT prompt
            prompt = update_sheets_and_get_prompt(context, jd, keywords)
            if not prompt:
                page.goto(APPLY_JOBS_URL, wait_until="domcontentloaded", timeout=30000)
                time.sleep(3)
                continue

            # Step 3: ChatGPT → Summary, Ex Desc 1, Ex Desc 2
            summary, ex1, ex2 = get_chatgpt_response(context, prompt)

            # Step 4: Write to Excel (1 row)
            append_to_excel(summary, ex1, ex2)

            # Step 5: Run resume_automation.py to build + download this resume
            print("  [*] Building and downloading resume...")
            result = subprocess.run([PYTHON_EXE, str(RESUME_AUTO)], capture_output=True, text=True)
            if result.returncode == 0:
                print("  [+] Resume built and downloaded.")
            else:
                print(f"  [!] Resume automation error: {result.stderr[:100]}")

            processed += 1

            # Step 6: Go back to Apply Jobs for next job
            print("  [*] Going back to Apply Jobs...")
            page.goto(APPLY_JOBS_URL, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_selector('[role="combobox"]', timeout=15000)
            setup_filters(page)
            time.sleep(2)

        print(f"\n[+] Processed {processed} job(s).")

        if processed > 0:
            print("\n[*] Running resume automation...")
            result = subprocess.run([PYTHON_EXE, str(RESUME_AUTO)], capture_output=True, text=True)
            print(result.stdout)
            if result.returncode != 0:
                print(f"[!] Resume automation errors: {result.stderr[:200]}")

        input("\n[*] Done! Press Enter to close...")


if __name__ == "__main__":
    main()
