from __future__ import annotations
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright
from xaurum_common import (
    DL_DIR, AUTH_STATE, should_run_headless,
    close_popups_everywhere, robust_goto, ensure_logged_in,
    find_export_button, wait_loading_token, notify_failure
)

BASE_URL = "https://equans.xaurum.be/nl/dispatcher/certificates"
TASK_NAME = "Dispatcher – Certificates"

def attempt_once(pg) -> Path:
    # pagina klaarzetten
    robust_goto(pg, BASE_URL)
    wait_loading_token(pg, "Loading certificates")
    close_popups_everywhere(pg)

    btn = find_export_button(pg)
    if not btn:
        pg.mouse.wheel(0, 900); btn = find_export_button(pg)
    if not btn:
        raise RuntimeError("Excel-knop niet gevonden (Certificates).")

    with pg.expect_download(timeout=120_000) as di:
        btn.click()
    dl = di.value
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = DL_DIR / f"{ts}_certificates_overview.xls"
    dl.save_as(str(out))
    return out

def run():
    last_err = None
    last_shot = None
    logbuf = []

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=should_run_headless())
        ctx = browser.new_context(accept_downloads=True, storage_state=str(AUTH_STATE) if AUTH_STATE.exists() else None)
        pg = ctx.new_page()
        pg.set_default_timeout(120_000); pg.set_default_navigation_timeout(120_000)

        # login
        robust_goto(pg, BASE_URL)
        ensure_logged_in(pg, ctx)

        for attempt in range(1, 4):  # 1 + 2 retries
            try:
                logbuf.append(f"Attempt {attempt}")
                out = attempt_once(pg)
                print(f"✅ Bestand opgeslagen in {out}")
                # optionele screenshot (alleen zichtbaar)
                if not should_run_headless():
                    try:
                        shot = DL_DIR / f"certificates_page_{datetime.now():%Y%m%d_%H%M%S}.png"
                        pg.screenshot(path=str(shot), full_page=False)
                        print(f"📸 Screenshot: {shot}")
                    except Exception:
                        pass
                ctx.storage_state(path=str(AUTH_STATE))
                ctx.close(); browser.close()
                return
            except Exception as e:
                last_err = e
                print(f"⚠️  Mislukt (poging {attempt}): {e}")
                try:
                    close_popups_everywhere(pg)
                    pg.wait_for_timeout(1500)
                    robust_goto(pg, BASE_URL)  # refresh voor volgende poging
                except Exception:
                    pass
                if attempt == 3:
                    # laatste poging: screenshot voor mail
                    try:
                        last_shot = DL_DIR / f"certificates_error_{datetime.now():%Y%m%d_%H%M%S}.png"
                        pg.screenshot(path=str(last_shot), full_page=True)
                    except Exception:
                        last_shot = None

        # alle pogingen mislukt → mail
        notify_failure(TASK_NAME, last_err or RuntimeError("Onbekende fout"), last_shot, "\n".join(logbuf))
        ctx.close(); browser.close()
        raise SystemExit(1)

if __name__ == "__main__":
    run()
