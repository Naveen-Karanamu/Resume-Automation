"""
Resume Builder Automation
Website: https://builder.firststepjob.com/resume-builder

Workflow:
  1. Reads login credentials from credentials.txt
  2. Reads resume data rows from resume_data.xlsx
  3. For each row:
       - Opens the builder and logs in
       - Clicks "My Resumes" in the navbar and loads the last (RAW) resume
       - Fills fields based on exact selectors discovered from the live page:
           • Objective field          → textarea[name="summary"]   ← Job Title
           • Location (profile)       → input[name="location"][0]  ← Job Location
           • Job Title (1st Work Ex)  → input[name="jobTitle"][0]
           • Job Title (3rd Work Ex)  → input[name="jobTitle"][2]
           • Ex Desc 1                → div[contenteditable][1]    (1st work ex)
           • Ex Desc 2                → div[contenteditable][3]    (3rd work ex)
       - Downloads the PDF to C:\\Users\\karan\\Downloads\\Resumes\\<ResumeName>.pdf

Run:
    python resume_automation.py
"""

import os
import re
import time
from datetime import datetime
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────

WEBSITE_URL  = "https://builder.firststepjob.com/resume-builder"
CREDS_FILE   = Path(__file__).parent / "credentials.txt"
EXCEL_FILE   = Path(__file__).parent / "resume_data.xlsx"
DOWNLOAD_DIR = Path(r"C:\Users\617404611\Downloads\Temp res")

HEADLESS = False   # False = watch the browser; True = run silently
SLOW_MO  = 200     # ms between actions

# ─────────────────────────────────────────────────────────────────────────────


def read_credentials():
    creds = {}
    with open(CREDS_FILE, "r") as f:
        for line in f:
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                key, _, value = line.partition("=")
                creds[key.strip()] = value.strip()
    if "email" not in creds or "password" not in creds:
        raise ValueError("credentials.txt must have 'email' and 'password' lines.")
    return creds["email"], creds["password"]


def read_excel():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(cell is None or str(cell).strip() == "" for cell in row):
            continue
        rows.append({
            "job_title":    str(row[0]).strip() if row[0] else "",
            "job_location": str(row[1]).strip() if row[1] else "",
            "summary":      str(row[2]).strip() if row[2] else "",
            "ex_desc_1":    str(row[3]).strip() if row[3] else "",
            "ex_desc_2":    str(row[4]).strip() if row[4] else "",
            "resume_name":  str(row[5]).strip() if len(row) > 5 and row[5] else "",
        })
    return rows


def safe_filename(text):
    return re.sub(r'[\\/*?:"<>|]', "_", text).strip() or "resume"


def login(page, email, password):
    print("  [*] Loading login page...")
    page.goto(WEBSITE_URL, wait_until="networkidle", timeout=30000)
    time.sleep(2)

    page.fill("#email", email)
    page.fill("#password", password)
    page.click('button[type="submit"]')
    print("  [*] Submitted login...")
    page.wait_for_load_state("networkidle", timeout=20000)
    time.sleep(3)
    print(f"  [+] Logged in. URL: {page.url}")


def load_last_resume(page):
    """
    Click 'My Resumes' in the navbar, then click the title of the LAST resume
    card (the RAW resume at the bottom). Each card has a div.flex-1.min-w-0.cursor-pointer
    wrapping the h3 title — clicking that loads the resume without duplicating it.
    """
    print("  [*] Opening 'My Resumes' panel...")
    page.click('button[title="My Saved Resumes"]')

    # Wait for the panel to slide in (it transitions from translate-x-full)
    page.wait_for_function("""() => {
        const el = document.querySelector('[class*="translate"]');
        return el && !el.className.includes('translate-x-full');
    }""", timeout=10000)
    time.sleep(1)

    # Each resume card row has: div.flex-1.min-w-0.cursor-pointer > h3 (title)
    # Clicking that div loads the resume. The buttons inside are Duplicate + Delete.
    card_title_divs = page.locator('[class*="translate"] .flex-1.min-w-0.cursor-pointer').all()
    print(f"  [*] Found {len(card_title_divs)} resume card(s).")

    if not card_title_divs:
        print("  [!] No resume cards found in panel.")
        return

    # Click the LAST card's title div (RAW resume is at the bottom)
    last_card = card_title_divs[-1]
    title = last_card.locator('h3').inner_text()
    print(f"  [*] Clicking resume: '{title}'")
    last_card.click()
    # SPA — no full page reload, just wait for React to re-render the form
    time.sleep(4)
    print(f"  [+] Loaded resume: '{title}'")


def clear_and_type_input(el, value):
    """Click, select-all, then type — replaces any existing content."""
    el.scroll_into_view_if_needed()
    el.click()
    el.press("Control+a")
    el.fill(value)


_BOLD_SEG = re.compile(r'\*\*(.+?)\*\*', re.DOTALL)


def _bold_segments(value):
    """Split a string with **bold** markers into [(text, is_bold), ...]."""
    segs, pos = [], 0
    for m in _BOLD_SEG.finditer(value):
        if m.start() > pos:
            segs.append((value[pos:m.start()], False))
        segs.append((m.group(1), True))
        pos = m.end()
    if pos < len(value):
        segs.append((value[pos:], False))
    return segs


def clear_and_type_contenteditable(el, value):
    """Click, clear, then type `value` into the contenteditable, converting
    markdown **bold** markers into REAL bold text.

    The resume builder does not render markdown, so literal '**' would appear
    in the resume. Instead we toggle rich-text bold (document.execCommand) around
    each bold segment, producing <b>...</b> that the builder renders as bold and
    that shows up bold in the downloaded PDF.
    """
    page = el.page
    el.scroll_into_view_if_needed()
    el.click()
    el.press("Control+a")
    el.press("Delete")

    if "**" not in value:
        # no bold markers — fast path: insert the whole string at once
        page.keyboard.insert_text(value)
    else:
        bold_on = False
        for text, is_bold in _bold_segments(value):
            if is_bold != bold_on:
                page.evaluate("document.execCommand('bold', false, null)")
                bold_on = is_bold
            if text:
                page.keyboard.insert_text(text)
        if bold_on:                      # turn bold back off at the end
            page.evaluate("document.execCommand('bold', false, null)")

    # notify React that the field changed so the preview/state updates
    try:
        el.evaluate("e => e.dispatchEvent(new InputEvent('input', {bubbles:true}))")
    except Exception:
        pass


def update_resume(page, row):
    print("  [*] Filling resume fields...")

    # ── Objective (textarea name="summary") → Job Title ───────────────────
    if row["job_title"]:
        try:
            el = page.locator('textarea[name="summary"]').first
            clear_and_type_input(el, row["job_title"])
            print("  [+] Objective (Job Title) filled.")
        except Exception as e:
            print(f"  [!] Objective failed: {e}")

    # ── Location (profile header, index 0) → Job Location ────────────────
    if row["job_location"]:
        try:
            location_els = page.locator('input[name="location"]').all()
            if location_els:
                clear_and_type_input(location_els[0], row["job_location"])
                print("  [+] Location filled.")
            else:
                print("  [!] No location input found.")
        except Exception as e:
            print(f"  [!] Location failed: {e}")

    # ── Job Title fields ──────────────────────────────────────────────────
    # jobTitle[0] = 1st work ex, jobTitle[2] = 3rd work ex
    if row["job_title"]:
        job_title_els = page.locator('input[name="jobTitle"]').all()
        print(f"  [*] Found {len(job_title_els)} jobTitle field(s).")

        if len(job_title_els) >= 1:
            try:
                clear_and_type_input(job_title_els[0], row["job_title"])
                print("  [+] Job Title (1st work ex) filled.")
            except Exception as e:
                print(f"  [!] Job Title 1st failed: {e}")

        if len(job_title_els) >= 3:
            try:
                clear_and_type_input(job_title_els[2], row["job_title"])
                print("  [+] Job Title (3rd work ex) filled.")
            except Exception as e:
                print(f"  [!] Job Title 3rd failed: {e}")
        else:
            print(f"  [~] Only {len(job_title_els)} jobTitle field(s) — skipping 3rd.")

    # ── Experience Description contenteditable divs ───────────────────────
    # On a loaded resume these divs have NO data-placeholder — select by class only.
    # Index 0 = Summary/profile description
    # Index 1 = 1st work experience description  ← Ex Desc 1
    # Index 2 = 2nd work experience description
    # Index 3 = 3rd work experience description  ← Ex Desc 2
    desc_els = page.locator(
        'div.mt-1.px-3.py-2.block.w-full.rounded-md.border.border-gray-300[contenteditable]'
    ).all()
    print(f"  [*] Found {len(desc_els)} contenteditable description field(s).")

    if row["ex_desc_1"] and len(desc_els) >= 2:
        try:
            clear_and_type_contenteditable(desc_els[1], row["ex_desc_1"])
            print("  [+] Ex Desc 1 (1st work ex) filled.")
        except Exception as e:
            print(f"  [!] Ex Desc 1 failed: {e}")

    if row["ex_desc_2"] and len(desc_els) >= 4:
        try:
            clear_and_type_contenteditable(desc_els[3], row["ex_desc_2"])
            print("  [+] Ex Desc 2 (3rd work ex) filled.")
        except Exception as e:
            print(f"  [!] Ex Desc 2 failed: {e}")
    elif row["ex_desc_2"]:
        print(f"  [~] Only {len(desc_els)} description field(s) — skipping Ex Desc 2.")

    time.sleep(1)


def download_resume(page, resume_name):
    """
    Click the Download button in the bottom bar → a modal appears asking for
    the resume filename → clear the pre-filled name, type resume_name → click
    the Download button inside the modal to save the PDF.
    """
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    print(f"  [*] Starting download flow for '{resume_name}'...")

    # ── Step 1: Click the Download button in the bottom toolbar ──────────
    # From the screenshot it's the last button in the bottom bar labelled "Download"
    try:
        dl_bar_btn = page.locator('button:has-text("Download")').last
        dl_bar_btn.scroll_into_view_if_needed()
        dl_bar_btn.click()
        print("  [*] Clicked bottom toolbar Download button.")
        time.sleep(2)
    except Exception as e:
        print(f"  [!] Could not click bottom Download button: {e}")
        return False

    # ── Step 2: Modal appears — clear the filename input and type our name ─
    try:
        # The input inside the modal (pre-filled with existing resume name)
        name_input = page.locator('dialog input, [role="dialog"] input, div[class*="modal"] input').first
        if name_input.count() == 0:
            # Fallback: any visible input that appeared after clicking Download
            name_input = page.locator('input[type="text"]').last
        name_input.click()
        name_input.press("Control+a")
        name_input.fill(resume_name)
        print(f"  [+] Entered resume name: '{resume_name}'")
        time.sleep(1)
    except Exception as e:
        print(f"  [!] Could not fill resume name in modal: {e}")
        return False

    # ── Step 3: Click the Download button inside the modal ────────────────
    try:
        with page.expect_download(timeout=30000) as dl_info:
            # The modal has a blue "Download" button — it's distinct from Cancel
            modal_dl_btn = page.locator(
                'dialog button:has-text("Download"), '
                '[role="dialog"] button:has-text("Download"), '
                'div[class*="fixed"] button:has-text("Download")'
            ).last
            modal_dl_btn.click()
            print("  [*] Clicked modal Download button.")

        download = dl_info.value
        save_path = DOWNLOAD_DIR / f"{safe_filename(resume_name)}.pdf"
        # If file already exists and is locked, add a timestamp suffix
        if save_path.exists():
            ts = datetime.now().strftime("%H%M%S")
            save_path = DOWNLOAD_DIR / f"{safe_filename(resume_name)}_{ts}.pdf"
        download.save_as(str(save_path))
        print(f"  [+] Saved: {save_path}")
        return True

    except PlaywrightTimeoutError:
        print("  [!] Download timed out — modal button selector may need updating.")
    except Exception as e:
        print(f"  [!] Download error: {e}")
    return False


def run_for_row(browser, email, password, row, index):
    print(f"\n{'='*60}")
    print(f"  Row {index} | Job Title: {row['job_title']}")
    print(f"{'='*60}")

    context = browser.new_context(
        accept_downloads=True,
        viewport={"width": 1280, "height": 900},
    )
    page = context.new_page()

    try:
        login(page, email, password)
        load_last_resume(page)
        update_resume(page, row)
        download_resume(page, row["resume_name"] or row["job_title"] or f"resume_{index}")
    except Exception as e:
        print(f"  [ERROR] Row {index}: {e}")
    finally:
        time.sleep(2)
        context.close()


def main():
    print("=" * 60)
    print("  FirstStep Resume Builder — Batch Automation")
    print("=" * 60)

    email, password = read_credentials()
    print(f"\n[*] Credentials loaded for: {email}")

    rows = read_excel()
    if not rows:
        print("[!] No data rows found in resume_data.xlsx. Fill in the sheet and re-run.")
        return

    print(f"[*] Found {len(rows)} resume(s) to generate.\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)
        for i, row in enumerate(rows, start=1):
            run_for_row(browser, email, password, row, i)
        browser.close()

    print(f"\n[+] All done! PDFs saved in: {DOWNLOAD_DIR}")


if __name__ == "__main__":
    main()
