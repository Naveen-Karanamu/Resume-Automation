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
CREDS_FILE   = Path(__file__).parent.parent / "credentials.txt"
EXCEL_FILE   = Path(__file__).parent.parent / "resume_data.xlsx"
DOWNLOAD_DIR = Path(r"C:\Users\karan\Downloads\Resumes")

HEADLESS = False   # False = watch the browser; True = run silently
SLOW_MO  = 200     # ms between actions

# Page fitting: if the resume renders on MORE than MAX_PAGES, drop bullets from
# the experience descriptions until it fits (or until MIN_BULLETS remain).
# MAX_PAGES is read from credentials.txt ('max_pages='); this is a fallback.
FIT_PAGES   = True
MAX_PAGES   = 1       # default page cap if credentials.txt has no 'max_pages'
MIN_BULLETS        = 3   # never trim Ex Desc 1 & 2 below this many bullet points
MIN_BULLETS_EX1    = 4   # floor for Ex Desc 1 specifically
MIN_BULLETS_EX2    = 3   # floor for Ex Desc 2 specifically
MIN_BULLETS_EXTRA  = 1   # Ex Desc 3 & 4 can go down to 1 bullet if needed
MAX_TRIM_ITERS = 40   # generous safety cap (enough to trim both lists to MIN)

# ─────────────────────────────────────────────────────────────────────────────


def read_credentials():
    """Return (email, password, max_pages, download_dir, desc_indices) from credentials.txt."""
    creds = {}
    with open(CREDS_FILE, "r") as f:
        for line in f:
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                key, _, value = line.partition("=")
                creds[key.strip()] = value.strip()
    if "email" not in creds or "password" not in creds:
        raise ValueError("credentials.txt must have 'email' and 'password' lines.")

    # max_pages: positive integer; falls back to MAX_PAGES if absent/invalid.
    max_pages = MAX_PAGES
    raw = (creds.get("max_pages", "") or "").strip()
    if raw:
        m = re.search(r'\d+', raw)
        if m and int(m.group()) >= 1:
            max_pages = int(m.group())

    # download_dir: falls back to hardcoded default if absent.
    download_dir = Path(creds["download_dir"]) if creds.get("download_dir") else DOWNLOAD_DIR

    # desc field indices — fixed per resume template, set in credentials.txt
    def _idx(key, default):
        try:
            return int(creds[key])
        except (KeyError, ValueError):
            return default

    desc_indices = {
        "summary": _idx("summary_idx", 0),
        "ex1":     _idx("ex1_idx", 1),
        "ex2":     _idx("ex2_idx", 2),
    }

    return creds["email"], creds["password"], max_pages, download_dir, desc_indices


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


RAW_RESUME_HINT = "RAW"   # the RAW template card contains this in its title


def _get_resume_cards(page):
    """Return the list of resume-card title locators currently in the panel."""
    return page.locator('[class*="translate"] .flex-1.min-w-0.cursor-pointer').all()


def load_last_resume(page):
    """
    Open 'My Resumes' and click the RAW resume card (preferred by its name,
    e.g. 'Ajay ... RAW Resume'); fall back to the LAST card if no RAW match.

    This is made robust against the intermittent 'Found 0 cards' failure by:
      - retrying the panel-open up to 3 times,
      - polling for cards to render (instead of a single fixed-wait read),
      - verifying the resume actually loaded (form fields populate), retrying.
    """
    for attempt in range(1, 4):
        # If cards are already visible (panel already open), skip re-opening.
        cards = _get_resume_cards(page)
        if not cards:
            print(f"  [*] Opening 'My Resumes' panel (attempt {attempt})...")
            opened = False
            for sel in ('button[title="My Saved Resumes"]',
                        'button:has-text("My Resumes")',
                        'button:has-text("My Saved Resumes")'):
                try:
                    btn = page.locator(sel).first
                    if btn.count() > 0:
                        btn.click(timeout=5000)
                        opened = True
                        break
                except Exception:
                    continue
            if not opened:
                print("    [!] My Resumes button not clickable — retrying...")
                time.sleep(2)
                continue

            # Wait for the panel to slide in.
            try:
                page.wait_for_function("""() => {
                    const el = document.querySelector('[class*="translate"]');
                    return el && !el.className.includes('translate-x-full');
                }""", timeout=8000)
            except Exception:
                pass

        # Poll for cards to actually render (up to ~8s) instead of reading once.
        for _ in range(16):
            cards = _get_resume_cards(page)
            if cards:
                break
            time.sleep(0.5)

        if not cards:
            print("    [!] No resume cards yet — closing & retrying panel...")
            # press Escape / re-click the button to toggle the panel and retry
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            time.sleep(1.5)
            continue

        print(f"    [+] Found {len(cards)} resume card(s).")

        # Pick the RAW resume card by name; fall back to the last card.
        target, target_title = None, ""
        for c in cards:
            try:
                t = c.locator('h3').inner_text().strip()
            except Exception:
                t = ""
            if RAW_RESUME_HINT.lower() in t.lower():
                target, target_title = c, t
                break
        if target is None:
            target = cards[-1]
            try:
                target_title = target.locator('h3').inner_text().strip()
            except Exception:
                target_title = "(last card)"

        print(f"    [*] Clicking resume: '{target_title}'")
        try:
            target.scroll_into_view_if_needed()
            target.click()
        except Exception:
            try:
                target.evaluate("el => el.click()")
            except Exception as e:
                print(f"    [!] click failed: {str(e)[:70]}")
                continue

        # Verify the form actually populated (editable description fields appear).
        loaded = False
        for _ in range(16):          # up to ~8s for React to render the form
            n = page.locator(
                'div.mt-1.px-3.py-2.block.w-full.rounded-md.border.border-gray-300[contenteditable]'
            ).count()
            if n >= 2:
                loaded = True
                break
            time.sleep(0.5)

        if loaded:
            print(f"  [+] Loaded resume: '{target_title}'")
            time.sleep(1.5)
            return True
        print("    [!] Resume didn't load (no form fields) — retrying...")
        time.sleep(1.5)

    print("  [!] Could not load the RAW resume after 3 attempts.")
    return False


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


def _html_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _markdown_to_html(value, paragraph=False):
    """Convert a string with **bold** markers and newlines into builder HTML:
    each line becomes a <div ...>...</div> and **word** becomes <b>word</b>.
    If paragraph=True, the whole value is rendered as a single inline block
    (sentences separated by a space, not block divs) — used for the summary."""
    html_lines = []
    for ln in value.split("\n"):
        parts = []
        for text, is_bold in _bold_segments(ln):
            if not text:
                continue
            esc = _html_escape(text)
            parts.append(f"<b>{esc}</b>" if is_bold else esc)
        inner = "".join(parts)
        if inner:
            html_lines.append(inner)

    if paragraph:
        combined = " ".join(html_lines)
        return f'<p style="line-height: 1.5; text-align: justify; margin: 0;">{combined}</p>'

    return "".join(
        f'<div style="line-height: 1.5; text-align: justify;">{inner or "<br>"}</div>'
        for inner in html_lines
    )


def clear_and_type_contenteditable(el, value, paragraph=False):
    """Fill a contenteditable, converting markdown **bold** into REAL bold.

    The resume builder does not render markdown, so literal '**' would show in
    the resume. We set the element's innerHTML directly with <b>...</b> tags
    (deterministic — no reliance on cursor focus / execCommand), then dispatch
    an input event so the builder's React state and PDF preview update.

    For plain text (no ** markers) we keep the fast keyboard-insert path.
    If paragraph=True, all lines are merged into a single block (for summary).
    """
    page = el.page
    el.scroll_into_view_if_needed()
    el.click()
    el.press("Control+a")
    el.press("Delete")

    if "**" not in value and not paragraph:
        page.keyboard.insert_text(value)
    else:
        html = _markdown_to_html(value, paragraph=paragraph)
        el.evaluate(
            "(node, html) => { node.focus(); node.innerHTML = html; "
            "node.dispatchEvent(new InputEvent('input', {bubbles: true})); }",
            html,
        )

    try:
        el.evaluate("e => e.dispatchEvent(new InputEvent('input', {bubbles:true}))")
    except Exception:
        pass


def split_bullets(text):
    """Split an experience description into a list of bullet lines."""
    if not text:
        return []
    bullets = []
    for ln in str(text).split("\n"):
        ln = ln.strip()
        if not ln:
            continue
        # Remove any leading punctuation characters that shouldn't start a line
        ln = ln.lstrip(".,;")
        ln = ln.strip()
        # Remove all commas
        ln = ln.replace(",", "")
        if ln:
            bullets.append(ln)
    return bullets


def join_bullets(bullets):
    return "\n".join(bullets)


def _raw_canvas_count(page):
    """Count <canvas> page renders across every frame (incl. iframes)."""
    best = 0
    for fr in page.frames:
        try:
            n = fr.locator('canvas').count()
            if n > best:
                best = n
        except Exception:
            continue
    if best > 0:
        return best
    # Fallback: react-pdf / rpv page layers in any frame.
    for fr in page.frames:
        try:
            for s in ('.react-pdf__Page', '[data-page-number]',
                      '.rpv-core__page-layer'):
                n = fr.locator(s).count()
                if n > 0:
                    return n
        except Exception:
            continue
    return 0


def get_page_count(page):
    """
    Return how many pages the resume preview currently renders to.

    The PDF preview re-renders a few seconds AFTER an edit (observed ~4s), with
    ONE <canvas> per page. Reading too early returns a stale count, so we POLL
    until the count stays stable across consecutive samples before returning.
    Returns an int (>=1) or None if nothing renders.
    """
    stable_needed = 3          # consecutive equal samples => settled
    poll_gap      = 1.0        # seconds between samples
    max_polls     = 15         # up to ~15s for the preview to repaginate

    last, stable, seen_nonzero = -1, 0, False
    for _ in range(max_polls):
        time.sleep(poll_gap)
        cnt = _raw_canvas_count(page)
        if cnt > 0:
            seen_nonzero = True
        if cnt == last and cnt > 0:
            stable += 1
            if stable >= stable_needed:
                return cnt
        else:
            stable = 0
        last = cnt
    # Return whatever we last saw (best effort) if it never fully stabilised.
    return last if seen_nonzero and last > 0 else None


def fill_experience_descs(page, ex1_bullets, ex2_bullets, ex3_bullets=None, ex4_bullets=None,
                          ex1_idx=None, ex2_idx=None, ex3_idx=None, ex4_idx=None):
    """(Re)write experience contenteditable fields from bullet lists using detected indices."""
    desc_els = page.locator(
        'div.mt-1.px-3.py-2.block.w-full.rounded-md.border.border-gray-300[contenteditable]'
    ).all()
    for idx, bullets, label in [
        (ex1_idx, ex1_bullets, "Ex Desc 1"),
        (ex2_idx, ex2_bullets, "Ex Desc 2"),
        (ex3_idx, ex3_bullets, "Ex Desc 3"),
        (ex4_idx, ex4_bullets, "Ex Desc 4"),
    ]:
        if idx is not None and bullets is not None and len(desc_els) > idx:
            try:
                clear_and_type_contenteditable(desc_els[idx], join_bullets(bullets))
            except Exception as e:
                print(f"      [!] re-fill {label} failed: {str(e)[:80]}")
    time.sleep(0.8)


def fit_to_pages(page, ex1_bullets, ex2_bullets, ex3_bullets=None, ex4_bullets=None,
                 ex1_idx=None, ex2_idx=None, ex3_idx=None, ex4_idx=None, max_pages=1):
    """
    If the resume spills beyond `max_pages`, drop the LAST bullet from the
    longest experience list, re-fill, and re-check — repeating until it fits
    within max_pages or all lists reach MIN_BULLETS.
    """
    pages = get_page_count(page)
    if pages is None:
        print("  [~] Could not read page count on first check — will attempt trimming anyway.")
        pages = max_pages + 1  # treat unknown as over the limit so trimming runs
    print(f"  [*] Resume spans {pages} page(s); cap is {max_pages}.")
    if pages <= max_pages:
        return ex1_bullets, ex2_bullets, ex3_bullets, ex4_bullets

    # Build a list of all bullet lists that exist (non-None)
    all_lists = [
        (ex1_bullets, "Ex Desc 1", MIN_BULLETS_EX1),
        (ex2_bullets, "Ex Desc 2", MIN_BULLETS_EX2),
        (ex3_bullets, "Ex Desc 3", MIN_BULLETS_EXTRA),
        (ex4_bullets, "Ex Desc 4", MIN_BULLETS_EXTRA),
    ]

    for it in range(MAX_TRIM_ITERS):
        # Trim in strict order: Ex1 first, then Ex2, then Ex3, then Ex4.
        # Never move to the next block until the current one hits its floor.
        # This keeps each work experience block intact on one page rather than
        # splitting bullets across a page boundary.
        trimmed = False
        for bl, lbl, floor in all_lists:
            if bl is not None and len(bl) > floor:
                dropped = bl.pop()
                print(f"  [-] Trimming {lbl} -> now "
                      f"{len(ex1_bullets) if ex1_bullets else 0}/"
                      f"{len(ex2_bullets) if ex2_bullets else 0}/"
                      f"{len(ex3_bullets) if ex3_bullets else 0}/"
                      f"{len(ex4_bullets) if ex4_bullets else 0} bullets "
                      f"(removed: '{dropped[:50]}...')")
                trimmed = True
                break

        if not trimmed:
            print(f"  [!] Reached minimum bullets in all lists; "
                  f"still {pages} page(s) (cap {max_pages}). Cannot trim further.")
            break

        fill_experience_descs(page, ex1_bullets, ex2_bullets, ex3_bullets, ex4_bullets,
                              ex1_idx=ex1_idx, ex2_idx=ex2_idx, ex3_idx=ex3_idx, ex4_idx=ex4_idx)
        pages = get_page_count(page)
        if pages is None:
            print(f"      -> page count unreadable, continuing to trim.")
            continue
        print(f"      -> now {pages} page(s).")
        if pages <= max_pages:
            print(f"  [+] Fits within {max_pages} page(s).")
            break

    return ex1_bullets, ex2_bullets, ex3_bullets, ex4_bullets


def update_resume(page, row, max_pages=1, desc_indices=None):
    if desc_indices is None:
        desc_indices = {"summary": 0, "ex1": 1, "ex2": 2}
    print("  [*] Filling resume fields...")

    # ── Inject page-break CSS so work experience blocks never split across pages ──
    try:
        page.evaluate("""() => {
            const style = document.createElement('style');
            style.id = 'kiro-page-break-fix';
            style.textContent = `
                @media print {
                    .work-experience-item,
                    [class*="experience"],
                    [class*="work-ex"],
                    [class*="job-block"],
                    [class*="position"] {
                        page-break-inside: avoid !important;
                        break-inside: avoid !important;
                    }
                }
            `;
            if (!document.getElementById('kiro-page-break-fix')) {
                document.head.appendChild(style);
            }
        }""")
        print("  [+] Page-break CSS injected.")
    except Exception as e:
        print(f"  [~] Page-break CSS injection failed: {e}")

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
    desc_els = page.locator(
        'div.mt-1.px-3.py-2.block.w-full.rounded-md.border.border-gray-300[contenteditable]'
    ).all()
    print(f"  [*] Found {len(desc_els)} contenteditable description field(s).")
    for i, el in enumerate(desc_els):
        try:
            preview = el.inner_text().strip().replace("\n", " ")[:80]
        except Exception:
            preview = "?"
        print(f"      desc_els[{i}]: '{preview}'")

    # Use fixed indices from credentials.txt
    summary_idx = desc_indices["summary"]
    ex1_idx     = desc_indices["ex1"]
    ex2_idx     = desc_indices["ex2"]
    ex3_idx     = ex2_idx + 1 if ex2_idx + 1 < len(desc_els) else None
    ex4_idx     = ex2_idx + 2 if ex2_idx + 2 < len(desc_els) else None
    print(f"  [*] Using indices — summary:{summary_idx} ex1:{ex1_idx} ex2:{ex2_idx}")

    # ── Summary ────────────────────────────────────────────────────────────
    if row["summary"] and summary_idx is not None:
        try:
            summary_text = " ".join(
                ln.strip() for ln in row["summary"].splitlines() if ln.strip()
            )
            summary_text = summary_text.replace(",", "")
            print(f"  [*] Summary text ({len(summary_text)} chars): {summary_text[:120]}...")
            el = desc_els[summary_idx]
            el.scroll_into_view_if_needed()
            el.click()
            el.press("Control+a")
            el.press("Delete")
            html = _markdown_to_html(summary_text, paragraph=True)
            el.evaluate(
                "(node, html) => { node.focus(); node.innerHTML = html; "
                "node.dispatchEvent(new InputEvent('input', {bubbles: true})); }",
                html,
            )
            try:
                el.evaluate("e => e.dispatchEvent(new InputEvent('input', {bubbles:true}))")
            except Exception:
                pass
            print("  [+] Summary Textbox filled.")
        except Exception as e:
            print(f"  [!] Summary Textbox failed: {e}")
    elif row["summary"]:
        print("  [!] Could not detect summary field — skipping.")

    ex1_bullets = split_bullets(row["ex_desc_1"])
    ex2_bullets = split_bullets(row["ex_desc_2"])

    if ex1_bullets and ex1_idx is not None:
        try:
            clear_and_type_contenteditable(desc_els[ex1_idx], join_bullets(ex1_bullets))
            print(f"  [+] Ex Desc 1 filled ({len(ex1_bullets)} bullets) at index {ex1_idx}.")
        except Exception as e:
            print(f"  [!] Ex Desc 1 failed: {e}")
    elif ex1_bullets:
        print("  [!] Could not detect Ex Desc 1 field — skipping.")

    if ex2_bullets and ex2_idx is not None:
        try:
            clear_and_type_contenteditable(desc_els[ex2_idx], join_bullets(ex2_bullets))
            print(f"  [+] Ex Desc 2 filled ({len(ex2_bullets)} bullets) at index {ex2_idx}.")
        except Exception as e:
            print(f"  [!] Ex Desc 2 failed: {e}")
    elif ex2_bullets:
        print("  [!] Could not detect Ex Desc 2 field — skipping.")

    # Read existing bullets from fields after ex2 for trimming
    def read_existing_bullets(el):
        try:
            html = el.evaluate("e => e.innerHTML")
            if not html:
                return None
            # Convert <b>...</b> back to **...**
            import re as _re
            text = _re.sub(r'<b>(.*?)</b>', lambda m: f'**{m.group(1)}**', html, flags=_re.DOTALL)
            # Each <div> is a line; strip all remaining tags
            text = _re.sub(r'<div[^>]*>', '\n', text)
            text = _re.sub(r'<[^>]+>', '', text)
            import html as _html
            text = _html.unescape(text).strip()
            return split_bullets(text) if text else None
        except Exception:
            return None

    ex3_idx = (ex2_idx + 1) if ex2_idx is not None and ex2_idx + 1 < len(desc_els) else None
    ex4_idx = (ex2_idx + 2) if ex2_idx is not None and ex2_idx + 2 < len(desc_els) else None
    ex3_bullets = read_existing_bullets(desc_els[ex3_idx]) if ex3_idx is not None else None
    ex4_bullets = read_existing_bullets(desc_els[ex4_idx]) if ex4_idx is not None else None
    if ex3_bullets:
        print(f"  [*] Ex Desc 3 (2nd work ex) has {len(ex3_bullets)} existing bullets.")
    if ex4_bullets:
        print(f"  [*] Ex Desc 4 (4th work ex) has {len(ex4_bullets)} existing bullets.")

    time.sleep(1)

    # ── Fit the resume within max_pages by trimming bullets if needed ─────
    if FIT_PAGES and (ex1_bullets or ex2_bullets or ex3_bullets or ex4_bullets):
        fit_to_pages(page, ex1_bullets, ex2_bullets, ex3_bullets, ex4_bullets,
                     ex1_idx=ex1_idx, ex2_idx=ex2_idx, ex3_idx=ex3_idx, ex4_idx=ex4_idx,
                     max_pages=max_pages)

    time.sleep(1)


def download_resume(page, resume_name, download_dir=None):
    """
    Click the Download button in the bottom bar → a modal appears asking for
    the resume filename → clear the pre-filled name, type resume_name → click
    the Download button inside the modal to save the PDF.
    """
    if download_dir is None:
        download_dir = DOWNLOAD_DIR
    download_dir.mkdir(parents=True, exist_ok=True)
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
        save_path = download_dir / f"{safe_filename(resume_name)}.pdf"
        # If file already exists and is locked, add a timestamp suffix
        if save_path.exists():
            ts = datetime.now().strftime("%H%M%S")
            save_path = download_dir / f"{safe_filename(resume_name)}_{ts}.pdf"
        download.save_as(str(save_path))
        print(f"  [+] Saved: {save_path}")
        return True

    except PlaywrightTimeoutError:
        print("  [!] Download timed out — modal button selector may need updating.")
    except Exception as e:
        print(f"  [!] Download error: {e}")
    return False


def run_for_row(browser, email, password, row, index, max_pages=1, download_dir=None, desc_indices=None):
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
        update_resume(page, row, max_pages=max_pages, desc_indices=desc_indices)

        # Guard: verify page count before downloading. If still over the limit,
        # skip the download and warn rather than saving an oversized resume.
        final_pages = get_page_count(page)
        if final_pages is not None and final_pages > max_pages:
            print(f"  [!] SKIPPING download — resume is {final_pages} page(s) "
                  f"but max is {max_pages}. Trim could not reduce it further.")
        else:
            if final_pages is None:
                print(f"  [~] Page count unreadable before download — proceeding anyway.")
            download_resume(page, row["resume_name"] or row["job_title"] or f"resume_{index}", download_dir=download_dir)
    except Exception as e:
        print(f"  [ERROR] Row {index}: {e}")
    finally:
        time.sleep(2)
        context.close()


def main():
    print("=" * 60)
    print("  FirstStep Resume Builder — Batch Automation")
    print("=" * 60)

    email, password, max_pages, download_dir, desc_indices = read_credentials()
    print(f"\n[*] Credentials loaded for: {email}")
    print(f"[*] Max pages per resume (from credentials.txt): {max_pages}")
    print(f"[*] Download directory: {download_dir}")
    print(f"[*] Desc indices: summary={desc_indices['summary']} ex1={desc_indices['ex1']} ex2={desc_indices['ex2']}")

    rows = read_excel()
    if not rows:
        print("[!] No data rows found in resume_data.xlsx. Fill in the sheet and re-run.")
        return

    print(f"[*] Found {len(rows)} resume(s) to generate.\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)
        for i, row in enumerate(rows, start=1):
            run_for_row(browser, email, password, row, i, max_pages=max_pages, download_dir=download_dir, desc_indices=desc_indices)
        browser.close()

    print(f"\n[+] All done! PDFs saved in: {download_dir}")


if __name__ == "__main__":
    main()
