import os, sys, time, zipfile, smtplib, ssl
from pathlib import Path
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from email.message import EmailMessage

import pandas as pd
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

    # Guess subtype (pdf vs xlsx)
    subtype = "pdf" if attach_path.suffix.lower() == ".pdf" else "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    msg.add_attachment(data, maintype="application", subtype=subtype, filename=attach_path.name)

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
            return extracted
    return out

def pick_status_col_and_click_latest(page):
    """
    On 'Reports Generated', click the latest row's control in the 'Status' column.
    Tries common <a>/<button>/<i> inside that cell and expects a file download.
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

    first_row = table.locator("tbody tr").first

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
            try:
                with page.expect_download(timeout=120000) as dl:
                    loc.first.click()
                return dl.value
            except PWTimeout:
                pass
    raise RuntimeError("Could not trigger download from the Status cell.")

def run_bot():
    if not PORTAL_USER or not PORTAL_PASS:
        raise RuntimeError("Missing PORTAL_USER or PORTAL_PASS. Set GitHub Secrets.")

    dates = yesterday_ist_formats()
    print("Yesterday (IST):", dates, flush=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # -------- 1) Login --------
        page.goto(LOGIN_URL, timeout=120000)

        # Fill username
        for sel in ['input[name="username"]', 'input[name="email"]', '#username', '#loginId', 'input[type="text"]']:
            if page.locator(sel).count() > 0:
                page.fill(sel, PORTAL_USER)
                break
        else:
            raise RuntimeError("Username field not found. Update selector.")

        # Fill password
        for sel in ['input[name="password"]', '#password', 'input[type="password"]']:
            if page.locator(sel).count() > 0:
                page.fill(sel, PORTAL_PASS)
                break
        else:
            raise RuntimeError("Password field not found. Update selector.")

        # Click submit
        for sel in ['button[type="submit"]', 'button:has-text("Login")', 'button:has-text("Sign in")', '#login', 'input[type="submit"]']:
            if page.locator(sel).count() > 0:
                page.click(sel)
                break
        else:
            raise RuntimeError("Login button not found. Update selector.")

        page.wait_for_load_state("networkidle")

        # -------- 2) Go to Reports tab --------
        for sel in ['a:has-text("Reports")', 'button:has-text("Reports")', 'li:has-text("Reports") >> a', 'text=Reports']:
            if page.locator(sel).count() > 0:
                page.click(sel)
                break
        else:
            raise RuntimeError("Reports tab not found. Update selector.")
        page.wait_for_load_state("networkidle")

        # -------- 3) Choose 'Statement of Capital Flows' --------
        # If it's a real <select>
        if page.locator('select').count() > 0:
            selects = page.locator("select")
            picked = False
            for i in range(selects.count()):
                sel = selects.nth(i)
                try:
                    sel.select_option(label="Statement of Capital Flows")
                    picked = True
                    break
                except Exception:
                    pass
            if not picked:
                raise RuntimeError("Could not select 'Statement of Capital Flows' in a <select>.")
        else:
            # Custom dropdown (combobox style)
            # Try to open the dropdown
            opened = False
            for opener in ['[role="combobox"]', 'div[aria-haspopup="listbox"]', 'button:has-text("Select")', 'text=Statement of']:
                if page.locator(opener).count() > 0:
                    page.click(opener)
                    opened = True
                    break
            if not opened:
                # Sometimes clicking the label focuses it
                page.keyboard.press("Tab")
            item = page.locator('text="Statement of Capital Flows"').first
            if item.count() == 0:
                item = page.locator('text=Statement of Capital Flows').first
            if item.count() == 0:
                raise RuntimeError("Report option 'Statement of Capital Flows' not found.")
            item.click()

        # -------- 4) Set date = yesterday (IST) --------
        # Try common date input selectors and formats
        date_set = False
        for sel in ['input[type="date"]', 'input[name*="date"]', '#fromDate', '#asOnDate', 'input[placeholder*="Date"]', 'input[placeholder*="date"]']:
            if page.locator(sel).count() > 0:
                for val in [dates["yyyy_mm_dd"], dates["dd_mmm_yyyy"], dates["dd_mm_yyyy"], dates["dd_mm_yyyy_dash"]]:
                    try:
                        page.fill(sel, "")
                        page.fill(sel, val)
                        # If readonly, set via JS
                        readonly = page.locator(sel).evaluate("el => el.readOnly || el.getAttribute('readonly') !== null")
                        if readonly:
                            page.evaluate("(el, v) => { el.value = v; el.dispatchEvent(new Event('input', {bubbles:true})); }", page.locator(sel), val)
                        date_set = True
                        break
                    except Exception:
                        pass
                if date_set:
                    break
        if not date_set:
            print("Date input not set — if it's a calendar widget, replace with explicit picker clicks.", file=sys.stderr)

        # -------- 5) Click Execute / Generate --------
        clicked = False
        for sel in ['button:has-text("Execute")', 'button:has-text("Generate")', 'button:has-text("Submit")', 'input[type="submit"]']:
            if page.locator(sel).count() > 0:
                page.click(sel)
                clicked = True
                break
        if not clicked:
            raise RuntimeError("Execute/Generate button not found. Update selector.")

        # Site redirects to 'Reports Generated'
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(3000)  # give backend a moment

        # -------- 6) On 'Reports Generated' → click latest 'page icon' under Status --------
        download = pick_status_col_and_click_latest(page)
        saved = save_download(download, DOWNLOAD_DIR)

        # -------- 7) Email the downloaded file (as-is) --------
        send_email(
            subject=f"Statement of Capital Flows — {datetime.now(IST).strftime('%d %b %Y')}",
            body="Automated download attached.",
            attach_path=saved
        )
        print(f"Downloaded and emailed: {saved}")

        context.close()
        browser.close()

if __name__ == "__main__":
    try:
        run_bot()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        raise
