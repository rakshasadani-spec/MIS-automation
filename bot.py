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
    y = datetime.now(IST) - time
