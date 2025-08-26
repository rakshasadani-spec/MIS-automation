# bot.py
import os, sys, time, zipfile, smtplib, ssl, re
from pathlib import Path
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from email.message import EmailMessage

import pandas as pd  # keep installed even if not used yet
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ----------------- Config via env / GitHub Secrets -----------------
LOGIN_URL   = os.getenv("LOGIN_URL", "https://eclientreporting.nuvamaassetservices.com/wealthspectrum/app/loginWith")
PORTAL_USER = os.getenv("PORTAL_USER")
PORTAL_PASS = os.getenv("PORTAL_PASS")

EMAIL_FROM  = os.getenv("EMAIL_FROM")
EMAIL_TO    = os.getenv("EMAIL_TO", "")   # comma-separated
SMTP_HOST   = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT   = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER   = os.getenv("SMTP_USER", EMAIL_FROM)
SMTP_PASS   = os.getenv("SMTP_PASS")

DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR", "downloads"))
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

IST = ZoneInfo("Asia/Kolkata")

# ----------------- Helpers -----------------
def yesterday_ist_formats():
    y = datetime.now(IST) - timedelta(days=1)
    return {
        "yyyy_mm_dd": y.strftime("%Y-%m-%d"),
        "dd_mmm_yyyy": y.strftime("%d-%b-%Y"),
        "dd_mm_yyyy": y.strftime("%d/%m/%Y"),
        "dd_mm_yyyy_dash": y.strftime("%d-%m-%Y"),
    }

def send_email(subject: str, body: str, attach_path: Path):
    if not EMAIL_TO.strip():
        print("EMAIL_TO empty; skipping email.")
        return
    msg = EmailMessage()
    msg["From"] = EMAIL_FROM
    msg["To"] = [x.strip() for x in EMAIL_TO.split(",") if x.strip()]
    msg["Subject"] = subject
    msg.set_content(body)

    with open(attach_path, "rb") as f:
        data = f.read()

    # Guess subtype (pdf vs xlsx vs general octet)
    ext = attach_path.suffix.lower()
    if ext == ".pdf":
        maintype, subtype = "application", "pdf"
    elif ext in (".xlsx", ".xls"):
        maintype, subtype = "application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        maintype, subtype = "application", "octet-stream"

    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=attach_path.name)

    ctx = ssl.create_default_context()
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls(context=ctx)
        s.login(SMTP_USER, SMTP_PASS)
        s.send_message(msg)

def save_download(download, dest_dir: Path) -> Path:
    out = dest_dir / download.suggested_filename
    download.save_as(str(out))
    if out.suffix.lower() == ".zip":
        with zipfile.ZipFile(out, "r") as zf:
            members = [m for m in zf.namelist() if not m.endswith("/")]
            if not members:
                raise RuntimeError("ZIP file is empty")
            first = members[0]
            extracted = dest_dir / Path(first).name
            with zf.open(first) as src, open(extracted, "wb") as dst:
                dst.write(src.read())
            print(f"ZIP extracted: {extracted}")
            return extracted
    return out

def pick_status_col_and_click_latest(page):
    """
    On 'Reports Generated', click the latest row's control in the 'Status' column.
    Tries common <a>/<button>/<i>/<span> inside that cell and expects a file download.
    """
    table = page.locator("table").first
    table.wait_for(timeout=60000)

    headers = table.locator("thead tr th")
    count = headers.count()
    status_idx = None
    for i in range(count):
        text = headers.nth(i).inner_text().strip().lower()
        if "status" in text:
            status_idx = i + 1
            break
    if status_idx is None:
        status_idx = count  # fallback: last column
    print(f"Detected Status column index: {status_idx}")

    first_row = table.locator("tbody tr").first
    if first_row.count() == 0:
        raise RuntimeError("No rows found in 'Reports Generated' table.")

    # Try common clickable controls within the Status cell
    candidates = [
        f"td:nth-child({status_idx}) a",
        f"td:nth-child({status_idx}) button",
        f"td:nth-child({status_idx}) i",
        f"td:nth-child({status_idx}) span",
    ]
    for sel in candidates:
        loc = first_row.locator(sel)
        if loc.count() > 0:
            print(f"Clicking Status cell control via selector: {sel}")
            try:
                with page.expect_download(timeout=120000) as dl:
                    loc.first.click()
                return dl.value
            except PWTimeout:
                print("Click did not trigger a download; trying next candidate...")
                continue
    raise RuntimeError("Could not trigger download from the Status cell.")

# -------- Robust report picker --------
TARGET_TEXT = "statement of capital flows"

def try_select_dropdown(page) -> bool:
    """Scan all <select> elements and choose an option whose visible text contains our target (case-insensitive)."""
    selects = page.locator("select")
    n = selects.count()
    print(f"Found {n} <select> element(s).")
    for i in range(n):
        sel = selects.nth(i)
        # Collect option texts/values and print them to logs for debugging
        options = sel.evaluate(
            """s => Array.from(s.options).map(o => ({value:o.value, text:(o.textContent||'').trim()}))"""
        )
        print(f"[Report select #{i}] options: {options}", flush=True)
        # Find a case-insensitive contains match
        match = None
        for opt in options:
            if TARGET_TEXT in opt["text"].lower():
                match = opt
                break
        if match:
            print(f"Matched option on select #{i}: {match}")
            try:
                if match["value"]:
                    sel.select_option(value=match["value"])
                else:
                    sel.select_option(label=match["text"])
                return True
            except Exception as e:
                print(f"Select #{i} matched but selection failed: {e}", flush=True)
    return False

def try_custom_dropdown(page) -> bool:
    """Handle non-<select> dropdowns (comboboxes/menus) using regex matches."""
    openers = [
        '[role="combobox"]',
        'div[aria-haspopup="listbox"]',
        'button[aria-haspopup="listbox"]',
        'button:has-text("Select")',
        'button:has-text("Report")',
        'label:has-text("Report") + *',
        'div[role="button"]',
    ]
    opened = False
    for op in openers:
        loc = page.locator(op).first
        if loc.count() > 0:
            print(f"Trying to open custom dropdown via: {op}")
            loc.click()
            opened = True
            page.wait_for_timeout(600)
            break
    if not opened:
        print("Could not positively identify a dropdown opener; trying to focus via Tab...")
        page.keyboard.press("Tab")
        page.wait_for_timeout(300)

    # Try clicking the item by regex (case-insensitive)
    pattern_main = re.compile(r"Statement\s+of\s+Capital\s+Flows", re.I)
    pattern_alt  = re.compile(r"Capital\s+Flow", re.I)

    item = page.get_by_text(pattern_main).first
    if item.count() == 0:
        item = page.get_by_text(pattern_alt).first

    if item.count() > 0:
        print("Clicking dropdown item by text regex match.")
        item.click()
        return True

    # Role-based attempt
    candidates = page.get_by_role("option", name=pattern_main)
    if candidates.count() == 0:
        candidates = page.get_by_role("option", name=pattern_alt)
    if candidates.count() == 0:
        candidates = page.get_by_role("menuitem", name=pattern_main)
    if candidates.count() == 0:
        candidates = page.get_by_role("menuitem", name=pattern_alt)

    if candidates.count() > 0:
        print("Clicking role-based dropdown candidate.")
        candidates.first.click()
        return True

    print("Custom dropdown selection failed to find the target item.")
    return False

# ----------------- Main flow -----------------
def run_bot():
    if not PORTAL_USER or not PORTAL_PASS:
        raise RuntimeError("Missing PORTAL_USER or PORTAL_PASS. Set GitHub Secrets.")

    dates = yesterday_ist_formats()
    print("Yesterday (IST):", dates, flush=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        # Start tracing for diagnostics
        context.tracing.start(screenshots=True, snapshots=True, sources=True)
        page = context.new_page()

        try:
            # -------- 1) Login --------
            print("Navigating to login page...")
            page.goto(LOGIN_URL, timeout=120000)

            #
