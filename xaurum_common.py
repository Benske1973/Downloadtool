# xaurum_common.py
from __future__ import annotations
import os, sys, traceback
from pathlib import Path
from datetime import datetime
from playwright.sync_api import TimeoutError as PWTimeout

# ==========================
# Paden & configuratie
# ==========================
def _appdata_dir() -> Path:
    p = Path(os.environ.get("APPDATA", str(Path.home()))) / "XaurumUploader"
    p.mkdir(parents=True, exist_ok=True)
    return p

APPDATA_DIR = _appdata_dir()
AUTH_STATE = APPDATA_DIR / "xaurum_auth_state.json"   # Playwright sessie per gebruiker

def get_sync_download_dir() -> Path:
    r"""
    Voorkeur:
      1) SharePoint bibliotheek (lokaal gesynchroniseerd):
         C:\Users\<user>\EQUANS\Competentie Management - Documenten\XaurumTools\downloads
         (of eender welke "<Site> - Documenten\XaurumTools\downloads")
      2) OneDrive (bedrijf/persoonlijk):
         <OneDrive>\CompetentieManagement\XaurumTools\downloads
      3) Fallback lokaal: .\downloads_xaurum
    Je kunt altijd forceren met env var XAURUM_DOWNLOAD_DIR.
    """
    override = os.environ.get("XAURUM_DOWNLOAD_DIR")
    if override:
        p = Path(override); p.mkdir(parents=True, exist_ok=True); return p

    userprofile = Path(os.environ.get("USERPROFILE", "C:\\Users\\Default"))

    # 1) SharePoint-sync ("* - Documenten")
    try:
        equans_root = userprofile / "EQUANS"
        if equans_root.exists():
            # Eerst specifiek zoeken naar "Competentie Management - Documenten"
            specific = equans_root / "Competentie Management - Documenten" / "XaurumTools" / "downloads"
            if specific.parent.parent.exists():  # Check of de site map bestaat
                try:
                    specific.mkdir(parents=True, exist_ok=True)
                    return specific
                except Exception:
                    pass
            
            # Anders: zoek alle "* - Documenten" mappen
            sp_choice = None
            for site in equans_root.glob("* - Documenten"):
                target = site / "XaurumTools" / "downloads"
                try:
                    target.mkdir(parents=True, exist_ok=True)
                    # Geef voorkeur aan competentie management site
                    if "competentie management" in site.name.lower():
                        return target
                    # Bewaar eerste werkende optie als fallback
                    if sp_choice is None:
                        sp_choice = target
                except Exception:
                    continue
            if sp_choice is not None:
                return sp_choice
    except Exception:
        pass

    # 2) OneDrive
    od_candidates = []
    # Eerst environment variables proberen
    for k in ("OneDriveCommercial", "OneDrive"):
        v = os.environ.get(k)
        if v: od_candidates.append(Path(v))
    # Dan zoeken in userprofile
    od_candidates += list(userprofile.glob("OneDrive*"))
    od_candidates += list(userprofile.glob("OneDrive - *"))
    
    for base in od_candidates:
        if not base.exists():
            continue
        try:
            p = base / "CompetentieManagement" / "XaurumTools" / "downloads"
            p.mkdir(parents=True, exist_ok=True)
            return p
        except Exception:
            continue

    # 3) fallback lokaal
    local = Path.cwd() / "downloads_xaurum"
    local.mkdir(parents=True, exist_ok=True)
    return local

DL_DIR = get_sync_download_dir()

# ==========================
# Headless logica
# ==========================
def should_run_headless() -> bool:
    """
    Bepaal of we headless kunnen draaien:
    - Als AUTH_STATE niet bestaat: ALTIJD zichtbaar (eerste login met MFA)
    - Als X_HEADLESS=0: zichtbaar (handmatig forceren)
    - Als X_HEADLESS=1 EN auth bestaat: headless
    
    LET OP: Deze functie is NIET gecached, zodat environment changes
    tijdens runtime worden opgepikt.
    """
    if not AUTH_STATE.exists():
        print("⚠️  Eerste keer aanmelden: browser wordt zichtbaar geopend voor Microsoft Authenticator.")
        print(f"    Auth state wordt opgeslagen in: {AUTH_STATE}")
        return False
    
    env_val = os.getenv("X_HEADLESS", "1")
    if env_val == "0":
        print("ℹ️  X_HEADLESS=0 gedetecteerd: browser wordt zichtbaar geopend.")
        return False
    
    return True

# Voor backwards compatibility: HEADLESS als variabele
# Maar deze checkt nu dynamisch via should_run_headless()
HEADLESS = should_run_headless()

# ==========================
# Playwright helpers
# ==========================
def close_popups_everywhere(page, rounds:int=4):
    sels = [
        "button:has-text('Sluiten')","button:has-text('OK')","button:has-text('Ok')",
        "button:has-text('Later')","button:has-text('Misschien later')",
        "text=Misschien later","text=OK","button[aria-label='Close']",
        "button.close",".modal-footer >> text=Sluiten",
    ]
    for _ in range(rounds):
        hit = False
        for s in sels:
            try:
                l = page.locator(s).first
                if l.is_visible(timeout=500):
                    l.click(timeout=500); hit = True
            except Exception:
                pass
        if not hit:
            break

def robust_goto(page, url: str, timeout:int=120_000):
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=timeout)
    except PWTimeout:
        try: close_popups_everywhere(page)
        except Exception: pass
        page.wait_for_timeout(1500)
        page.goto(url, wait_until="domcontentloaded", timeout=timeout + 60_000)

def ensure_logged_in(page, context):
    try:
        page.wait_for_selector("text=Dispatcher", timeout=6_000)
        print("✅ Al ingelogd (Dispatcher menu gevonden)")
    except PWTimeout:
        print("↪️ Wachten op login/MFA (max 180 s)…")
        print(f"   Browser headless: {should_run_headless()}")
        print(f"   X_HEADLESS env var: {os.getenv('X_HEADLESS', 'niet gezet')}")
        page.wait_for_selector("text=Dispatcher", timeout=180_000)
        context.storage_state(path=str(AUTH_STATE))
        print(f"✅ Login opgeslagen: {AUTH_STATE}")

def wait_loading_token(page, token: str="Loading", max_timeout:int=120_000):
    try:
        page.wait_for_selector(f"text={token}", timeout=5_000, state="visible")
        page.wait_for_selector(f"text={token}", timeout=max_timeout, state="detached")
    except Exception:
        pass

def find_export_button(page):
    cands = [
        "text=Download to Excel","text=Download to excel",
        "text=Export to excel","text=Export to Excel","text=Excel",
        "button:has-text('Excel')","a:has-text('Excel')",
    ]
    for sel in cands:
        try:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=800):
                return loc
        except Exception:
            pass
    return None

# ==========================
# E-mail bij falen
# ==========================
def _send_email_outlook(to: str, subject: str, html_body: str, attachments: list[Path]|None=None) -> bool:
    try:
        import win32com.client  # pywin32
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject
        mail.HTMLBody = html_body
        if attachments:
            for a in attachments:
                if a and Path(a).exists():
                    mail.Attachments.Add(str(a))
        mail.Send()
        return True
    except Exception as e:
        print(f"[mail] Outlook COM niet beschikbaar: {e}")
        return False

def _send_email_smtp(to: str, subject: str, text_body: str) -> bool:
    """
    Eenvoudige SMTP fallback — werkt alleen als je deze env vars zet:
      SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM
    """
    try:
        import smtplib
        from email.message import EmailMessage
        host = os.getenv("SMTP_HOST"); port = int(os.getenv("SMTP_PORT", "587"))
        user = os.getenv("SMTP_USER"); pw = os.getenv("SMTP_PASS"); from_addr = os.getenv("SMTP_FROM")
        if not all([host, user, pw, from_addr]):
            return False
        msg = EmailMessage()
        msg["From"] = from_addr
        msg["To"] = to
        msg["Subject"] = subject
        msg.set_content(text_body)
        with smtplib.SMTP(host, port, timeout=20) as s:
            s.starttls()
            s.login(user, pw)
            s.send_message(msg)
        return True
    except Exception as e:
        print(f"[mail] SMTP fallback faalde: {e}")
        return False

def notify_failure(task_name: str, exc: Exception, last_screenshot: Path|None=None, extra_log: str=""):
    to = "benny.ponet@equans.com"
    subject = f"[Xaurum downloads] Mislukt: {task_name} ({datetime.now():%Y-%m-%d %H:%M})"
    tb = traceback.format_exc()
    html = f"""
    <p>Downloadtaak <b>{task_name}</b> is definitief <b>mislukt</b> na herpogingen.</p>
    <p><b>Tijdstip:</b> {datetime.now():%Y-%m-%d %H:%M}<br>
       <b>Machine:</b> {os.environ.get('COMPUTERNAME','?')}</p>
    <p><b>Fout:</b><pre>{str(exc)}</pre></p>
    <p><b>Traceback:</b><pre>{tb}</pre></p>
    <p><b>Log:</b><pre>{extra_log}</pre></p>
    """
    if not _send_email_outlook(to, subject, html, [last_screenshot] if last_screenshot else None):
        # probeer smtp als tekst
        _send_email_smtp(to, subject, f"{task_name} mislukt.\n\n{str(exc)}\n\n{tb}\n\n{extra_log}")