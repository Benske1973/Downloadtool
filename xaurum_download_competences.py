from __future__ import annotations
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright
from xaurum_common import (
    DL_DIR, AUTH_STATE, should_run_headless,
    close_popups_everywhere, robust_goto, ensure_logged_in,
    find_export_button, notify_failure
)

BASE_URL = "https://equans.xaurum.be/nl/dispatcher/competences"
TASK_NAME = "Dispatcher – Competences"

def attempt_once(pg) -> Path:
    robust_goto(pg, BASE_URL)
    close_popups_everywhere(pg)
    try:
        pg.wait_for_selector("table, [role='grid'], .table", timeout=120_000)
    except Exception:
        pass
    btn = find_export_button(pg)
    if not btn:
        pg.mouse.wheel(0, 900); btn = find_export_button(pg)
    if not btn:
        raise RuntimeError("Excel-knop niet gevonden (Competences).")
    with pg.expect_download(timeout=120_000) as di:
        btn.click()
    dl = di.value
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = DL_DIR / f"{ts}_competences_overview.xls"
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

        robust_goto(pg, BASE_URL)
        ensure_logged_in(pg, ctx)

        for attempt in range(1, 4):
            try:
                logbuf.append(f"Attempt {attempt}")
                out = attempt_once(pg)
                print(f"✅ Bestand opgeslagen in {out}")
                ctx.storage_state(path=str(AUTH_STATE))
                ctx.close(); browser.close()
                return
            except Exception as e:
                last_err = e
                print(f"⚠️  Mislukt (poging {attempt}): {e}")
                try:
                    close_popups_everywhere(pg)
                    pg.wait_for_timeout(1500)
                    robust_goto(pg, BASE_URL)
                except Exception:
                    pass
                if attempt == 3:
                    try:
                        last_shot = DL_DIR / f"competences_error_{datetime.now():%Y%m%d_%H%M%S}.png"
                        pg.screenshot(path=str(last_shot), full_page=True)
                    except Exception:
                        last_shot = None

        notify_failure(TASK_NAME, last_err or RuntimeError("Onbekende fout"), last_shot, "\n".join(logbuf))
        ctx.close(); browser.close()
        raise SystemExit(1)

if __name__ == "__main__":
    run()
