# smart_auth_bootstrap.py  – VERBETERD MET DIRECT AUTH-OPSLAG
# Slaat auth op via de reeds geopende Selenium-sessie
# ---------------------------------------------------------------

from __future__ import annotations
import os, sys, time
from pathlib import Path
from typing import Optional
import json

# --- zorg dat we de Scripts-map kunnen importeren ---
HERE = Path(__file__).parent.resolve()
SCRIPTS = HERE / "Scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

# AUTH_STATE (waar Playwright zijn login vindt)
try:
    from xaurum_common import AUTH_STATE  # type: ignore
except Exception:
    appdata = Path(os.environ.get("APPDATA", str(Path.home()))) / "XaurumUploader"
    appdata.mkdir(parents=True, exist_ok=True)
    AUTH_STATE = appdata / "xaurum_auth_state.json"

# --- Selenium / Edge setup ---
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException, NoSuchElementException

XAURUM_START_URL = "https://equans.xaurum.be/nl/dispatcher/certificates"

def _default_profile_dir() -> Path:
    root = Path(os.environ.get("LOCALAPPDATA", str(Path.home())))
    return root / "XaurumTools" / "auth_browser_profile"

def _build_edge(headless: bool, user_data_dir: Path) -> webdriver.Edge:
    opts = EdgeOptions()
    opts.add_argument(f"--user-data-dir={user_data_dir}")
    opts.add_argument("--profile-directory=Default")
    
    # GROOT VENSTER ZODAT MFA-CODE ZICHTBAAR IS
    opts.add_argument("--start-maximized")
    opts.add_argument("--window-size=1920,1440")
    
    opts.add_argument("--disable-infobars")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--no-first-run")
    opts.add_argument("--disable-features=PrivacySandboxSettings4")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    
    # BELANGRIJK: Bij headless=False, zorg dat --headless NIET wordt toegevoegd
    if headless:
        opts.add_argument("--headless=new")
    
    return webdriver.Edge(options=opts)

def _is_on_microsoft_login(url: str) -> bool:
    u = (url or "").lower()
    return ("login.microsoftonline" in u) or ("aad" in u and "microsoft" in u)

def _logged_in_marker_present(drv: webdriver.Edge) -> bool:
    try:
        if "xaurum.be" not in (drv.current_url or "").lower():
            return False
        candidates = [
            (By.XPATH, "//*[contains(., 'Dispatcher')]"),
            (By.XPATH, "//*[contains(., 'Certificaten')]"),
            (By.XPATH, "//*[contains(., 'Certificates')]"),
            (By.XPATH, "//*[contains(., 'Mijn opleidingen')]"),
            (By.XPATH, "//*[contains(., 'Mes formations')]"),
            (By.XPATH, "//*[contains(., 'My trainings')]"),
            (By.CSS_SELECTOR, "header, nav"),
        ]
        for by, sel in candidates:
            try:
                el = drv.find_element(by, sel)
                if el and el.is_displayed():
                    return True
            except NoSuchElementException:
                pass
    except Exception:
        return False
    return False

def _extract_cookies_and_storage(drv: webdriver.Edge) -> dict:
    """
    Haal cookies en localStorage op via JavaScript en bouw een AUTH_STATE-achtige dict.
    Dit werkt beter dan Playwright bij Edge-profielen.
    """
    auth_state = {
        "cookies": [],
        "origins": []
    }
    
    try:
        # Cookies via Selenium
        for cookie in drv.get_cookies():
            c = {
                "name": cookie.get("name"),
                "value": cookie.get("value"),
                "domain": cookie.get("domain", ""),
                "path": cookie.get("path", "/"),
                "expires": cookie.get("expiry", -1),
                "httpOnly": cookie.get("httpOnly", False),
                "secure": cookie.get("secure", False),
                "sameSite": cookie.get("sameSite", "None")
            }
            auth_state["cookies"].append(c)
    except Exception as e:
        print(f"[auth] Kon cookies niet ophalen: {e}")
    
    try:
        # localStorage via JavaScript
        local_storage = drv.execute_script(
            "return Object.entries(localStorage).reduce((acc, [k, v]) => { acc[k] = v; return acc; }, {});"
        )
        if local_storage:
            auth_state["origins"] = [
                {
                    "origin": drv.current_url.split("/nl/")[0] if "/nl/" in drv.current_url else "https://equans.xaurum.be",
                    "localStorage": [
                        {"name": k, "value": v} for k, v in local_storage.items()
                    ]
                }
            ]
    except Exception as e:
        print(f"[auth] Kon localStorage niet ophalen: {e}")
    
    return auth_state

def _save_auth_state(drv: webdriver.Edge) -> None:
    """
    Bewaar auth via cookies en localStorage in AUTH_STATE.
    """
    try:
        auth_state = _extract_cookies_and_storage(drv)
        AUTH_STATE.parent.mkdir(parents=True, exist_ok=True)
        with open(AUTH_STATE, "w", encoding="utf-8") as f:
            json.dump(auth_state, f, indent=2, ensure_ascii=False)
        print(f"[auth] AUTH_STATE opgeslagen: {AUTH_STATE}")
        print(f"[auth] ({len(auth_state.get('cookies', []))} cookies)")
    except Exception as e:
        print(f"[auth] Kon AUTH_STATE niet opslaan: {e!r}")


def smart_bootstrap(timeout_sec: int = 600, start_url: Optional[str] = None) -> bool:
    """
    Start Edge ZICHTBAAR (grote venster voor MFA-code!) met vaste profielmap.
    Wacht tot Xaurum ingelogd is.
    
    Bij succes:
      - zet X_EDGE_USER_DATA_DIR
      - bewaar AUTH_STATE (cookies + localStorage)
      - zet X_HEADLESS=1
    """
    start_url = start_url or XAURUM_START_URL
    profile_dir = _default_profile_dir()
    profile_dir.parent.mkdir(parents=True, exist_ok=True)

    # Laat andere code dezelfde profielmap gebruiken
    os.environ["X_EDGE_USER_DATA_DIR"] = str(profile_dir)

    print("[auth] ========================================")
    print("[auth] XAURUM LOGIN BOOTSTRAP")
    print("[auth] ========================================")
    print(f"[auth] Profielmap  : {profile_dir}")
    print(f"[auth] Start URL   : {start_url}")
    print(f"[auth] Timeout     : {timeout_sec}s")
    print("[auth] ----------------------------------------")
    print("[auth] Browser opent ZICHTBAAR voor MFA.")
    print("[auth] Volg de inlogstappen en voer de 6-cijferige code in.")
    print("[auth] ========================================")

    drv = None
    try:
        # headless=False betekent ZICHTBAAR
        drv = _build_edge(headless=False, user_data_dir=profile_dir)
    except WebDriverException as e:
        print(f"[auth] FOUT: Kan Edge niet starten: {e}")
        return False

    try:
        drv.get(start_url)
        t0 = time.time()
        mfa_hint = False
        last_url = ""

        # Wachten tot ofwel MS login verdwijnt, of Xaurum geladen is
        while time.time() - t0 < timeout_sec:
            try:
                url = drv.current_url
            except Exception:
                url = ""
            
            # Toon huidige URL (voor debugging)
            if url != last_url:
                print(f"[auth] Huidige URL: {url}")
                last_url = url

            if _is_on_microsoft_login(url):
                if not mfa_hint:
                    print("[auth] ")
                    print("[auth] >>> MICROSOFT LOGIN/MFA SCHERM GEDETECTEERD <<<")
                    print("[auth] Vul je gegevens in en voer de 6-cijferige code in.")
                    print("[auth] Laat het browservenster OPEN tot je terugkeert naar Xaurum.")
                    print("[auth] ")
                    mfa_hint = True
                time.sleep(2.0)
                continue

            if _logged_in_marker_present(drv):
                print("[auth] ")
                print("[auth] ✅ SUCCES: Xaurum geladen en ingelogd!")
                print("[auth] ")
                
                # Bewaar auth-sessie DIRECT vanuit deze Selenium-sessie
                print("[auth] Opslaan auth-sessie voor automatische downloads…")
                _save_auth_state(drv)
                
                # Zet headless aan voor volgende scripts
                os.environ["X_HEADLESS"] = "1"
                print("[auth] X_HEADLESS ingesteld op: 1")
                print("[auth] ========================================")
                return True

            time.sleep(1.5)

        print("[auth] ")
        timeout_msg = f"[auth] ❌ TIME-OUT: Login/MFA kon niet worden afgerond binnen {timeout_sec}s."
        print(timeout_msg)
        print("[auth] Controleer:")
        print("[auth]   - Ziet je het browservenster?")
        print("[auth]   - Kun je de MFA-code zien?")
        print("[auth]   - Heb je genoeg tijd gehad?")
        print("[auth] ========================================")
        return False
        
    except Exception as e:
        print(f"[auth] ONVERWACHTE FOUT: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        try:
            drv.quit()
        except Exception:
            pass