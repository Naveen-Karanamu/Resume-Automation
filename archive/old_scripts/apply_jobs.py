"""
Apply Jobs Automation
Website: https://firststepjob.com/application-incharge/apply-jobs

Steps automated:
  1. Opens Chrome, you log in with Google once
  2. Sets Created By = "Job Collector"
  3. Sets Date = yesterday
  4. Sets Category = "Data Scientist"
  5. Clicks Fetch Jobs
  6. Scrolls to job listings

Run:
    python3.10 apply_jobs.py
"""

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from datetime import datetime, timedelta
from pathlib import Path
import subprocess
import time

URL        = "https://firststepjob.com/application-incharge/apply-jobs?userId=google-oauth2|110590764416739071400"
CHROME_EXE = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
DEBUG_PORT = 9222
TEMP_DIR   = r"C:\Users\karan\AppData\Local\Temp\ChromeDebugSession"
SLOW_MO    = 50


def get_yesterday():
    yesterday = datetime.today() - timedelta(days=1)
    return yesterday.strftime("%m/%d/%Y")


def select_mui_dropdown(page, label_text, option_text):
    """
    Select an option from a MUI Select dropdown.
    After selecting, press Escape and click the page title to close the dropdown.
    """
    print(f"  [*] Selecting '{option_text}' in '{label_text}' dropdown...")
    try:
        dropdown = page.locator(f'[role="combobox"]:near(:text("{label_text}"))').first
        if dropdown.count() == 0:
            dropdown = page.locator(f'[placeholder*="{label_text}" i]').first
        if dropdown.count() == 0:
            raise Exception(f"Could not find dropdown for '{label_text}'")

        dropdown.scroll_into_view_if_needed()
        dropdown.click()
        time.sleep(0.3)

        # Click the matching option
        option = page.locator(f'[role="option"]:has-text("{option_text}")').first
        option.wait_for(state="visible", timeout=5000)
        option.click()
        print(f"  [+] Selected '{option_text}'.")

        # Close the dropdown — press Escape then click body
        page.keyboard.press("Escape")
        time.sleep(0.2)
        page.locator("body").click(position={"x": 10, "y": 10}, force=True)
        time.sleep(0.2)

    except Exception as e:
        print(f"  [!] Dropdown '{label_text}' failed: {e}")


def set_date(page, date_str):
    """
    Set the Date field using the MUI calendar picker.
    Clicks the calendar icon, then clicks the day before today, then clicks OK.
    """
    print(f"  [*] Setting date to {date_str} (yesterday)...")
    try:
        yesterday = datetime.today() - timedelta(days=1)
        yesterday_day = str(yesterday.day)  # e.g. "12"

        # Click the date input / calendar icon to open the picker
        date_input = page.locator('input[type="date"], input[placeholder*="date" i]').first
        if date_input.count() == 0:
            # MUI date field — click the input area
            date_input = page.locator('input').filter(
                has_text=""
            ).nth(1)

        # Click the calendar icon button to open the picker
        cal_btn = page.locator('button[aria-label*="calendar" i], button[aria-label*="date" i], [class*="calendar"] button').first
        if cal_btn.count() > 0:
            cal_btn.click()
        else:
            # Click the input itself to open the picker
            page.locator('input').nth(1).click()

        time.sleep(1)

        # Click the correct day number in the calendar grid
        # Target a button with exact day text (not part of a larger number)
        day_btn = page.locator(
            f'button[type="button"]:has-text("{yesterday_day}")'
        ).filter(has_text=f"{yesterday_day}")

        # Make sure we click the exact day (e.g. "12" not "12" inside "12x")
        for btn in day_btn.all():
            if btn.inner_text().strip() == yesterday_day and btn.is_visible():
                btn.click()
                print(f"  [+] Clicked day {yesterday_day}.")
                break
        time.sleep(0.5)

        # Click OK to confirm
        ok_btn = page.locator('button:has-text("OK")').first
        if ok_btn.count() > 0:
            ok_btn.click()
            print("  [+] Clicked OK — date confirmed.")
        time.sleep(0.3)

    except Exception as e:
        print(f"  [!] Date error: {e}")


def click_fetch_jobs(page):
    """Wait for Fetch Jobs to be enabled and click it."""
    print("  [*] Waiting for Fetch Jobs button to be enabled...")
    try:
        btn = page.locator('button:has-text("Fetch Jobs")').first

        # Wait up to 10s for it to become enabled
        for _ in range(20):
            if btn.is_enabled():
                break
            time.sleep(0.5)

        if not btn.is_enabled():
            print("  [!] Fetch Jobs button still disabled — check if all dropdowns are filled.")
            return

        btn.scroll_into_view_if_needed()
        btn.click()
        print("  [+] Fetch Jobs clicked. Waiting for results...")
        time.sleep(3)
    except Exception as e:
        print(f"  [!] Could not click Fetch Jobs: {e}")


def scroll_to_jobs(page):
    """Scroll down to show job listings."""
    print("  [*] Scrolling to job listings...")
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    time.sleep(2)
    print("  [+] Scrolled to job listings.")


def main():
    print("=" * 60)
    print("  FirstStep — Apply Jobs Automation")
    print("=" * 60)

    yesterday = get_yesterday()
    print(f"\n[*] Date will be set to: {yesterday}")

    # Launch Chrome with remote debugging
    print("\n[*] Launching Chrome...")
    Path(TEMP_DIR).mkdir(parents=True, exist_ok=True)
    chrome_proc = subprocess.Popen([
        CHROME_EXE,
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--user-data-dir={TEMP_DIR}",
        "--no-first-run",
        "--no-default-browser-check",
        URL,
    ])
    print(f"[+] Chrome launched (PID {chrome_proc.pid})")
    time.sleep(2)

    print("\n[!] Log in with Google in the Chrome window.")
    print("[!] Script will auto-detect when you reach the Apply Jobs page...\n")

    # Connect Playwright to the running Chrome
    with sync_playwright() as p:
        print("[*] Connecting to Chrome...")
        browser = p.chromium.connect_over_cdp(f"http://localhost:{DEBUG_PORT}")
        context = browser.contexts[0]
        page = context.pages[0] if context.pages else context.new_page()

        # Auto-wait until the Apply Jobs page is loaded (poll every 2s, up to 3 mins)
        print("[*] Waiting for Apply Jobs page to load...")
        for _ in range(90):
            try:
                url = page.url
                if "apply-jobs" in url:
                    # Also check the page has rendered (Filters & Controls heading)
                    content = page.content()
                    if "Filters" in content or "Fetch Jobs" in content:
                        print(f"[+] Apply Jobs page detected!")
                        break
            except Exception:
                pass
            time.sleep(2)
        else:
            print("[!] Timed out waiting for Apply Jobs page.")
            chrome_proc.terminate()
            return

        time.sleep(1)  # let React fully settle

        try:
            # Fill all filters in order: Created By → Date → Category
            select_mui_dropdown(page, "Created By", "Job Collector")
            set_date(page, yesterday)
            select_mui_dropdown(page, "Category", "Data Scientist")

            # Click Fetch Jobs
            click_fetch_jobs(page)

            # Scroll to results
            scroll_to_jobs(page)

            print("\n[+] Done! Job listings should be visible.")
            print("[*] Press Enter to close the browser...")
            input()

        except Exception as e:
            print(f"\n[ERROR] {e}")
            input("Press Enter to close...")
        finally:
            browser.close()
            chrome_proc.terminate()


if __name__ == "__main__":
    main()
