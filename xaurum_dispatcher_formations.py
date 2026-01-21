from __future__ import annotations
from datetime import datetime, timedelta
from pathlib import Path
from playwright.sync_api import sync_playwright
from xaurum_common import (
    DL_DIR, AUTH_STATE, should_run_headless,
    close_popups_everywhere, robust_goto, ensure_logged_in,
    find_export_button, wait_loading_token, notify_failure
)

BASE_URL = "https://equans.xaurum.be/nl/dispatcher/formations"
TASK_NAME = "Dispatcher – Team Opleidingen"

def set_period_and_filters(pg):
    today = datetime.today()
    end = today + timedelta(days=730)
    # datumvelden (simpelste aanpak: eerste 2 inputvelden met '/')
    try:
        inputs = pg.locator("input[placeholder*='/']").all()
        visible = [i for i in inputs if i.is_visible(timeout=800)]
        if len(visible) >= 2:
            visible[0].fill(f"{today.day}/{today.month}/{today.year}")
            visible[1].fill(f"{end.day}/{end.month}/{end.year}")
    except Exception:
        pass
    # checkbox ter goedkeuring
    for sel in [
        "label:has-text('Ter goedkeuring') input",
        "text=Ter goedkeuring >> xpath=..//input",
        "xpath=//*[contains(., 'Ter goedkeuring')]/ancestor-or-self::*[1]//input[@type='checkbox']",
    ]:
        try:
            cb = pg.locator(sel).first
            if cb.is_visible(timeout=1200) and not cb.is_checked():
                cb.check(timeout=1200); break
        except Exception:
            pass

def attempt_once(pg) -> Path:
    robust_goto(pg, BASE_URL)
    close_popups_everywhere(pg)
    set_period_and_filters(pg)

    # Toon rapport en wacht tot laden klaar is
    clicked = False
    for sel in ["button:has-text('Toon rapport')", "text=Toon rapport >> ..//button"]:
        try:
            btn = pg.locator(sel).first
            if btn.is_visible(timeout=1500):
                btn.click(timeout=2000); clicked = True; break
        except Exception:
            pass
    if not clicked:
        pg.keyboard.press("Enter")

    wait_loading_token(pg, "Loading")
    close_popups_everywhere(pg)
    try:
        pg.wait_for_selector("table, [role='grid'], .table", timeout=120_000)
    except Exception:
        pass

    btn = find_export_button(pg)
    if not btn:
        pg.mouse.wheel(0, 1000); btn = find_export_button(pg)
    if not btn:
        raise RuntimeError("Excel-knop niet gevonden (Team opleidingen).")

    with pg.expect_download(timeout=180_000) as di:
        btn.click()
    dl = di.value
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = DL_DIR / f"rapport_teamopleidingen_{ts}.xls"
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
                if not should_run_headless():
                    try:
                        shot = DL_DIR / f"rapport_teamopleidingen_{datetime.now():%Y%m%d_%H%M%S}.png"
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
                    robust_goto(pg, BASE_URL)
                except Exception:
                    pass
                if attempt == 3:
                    try:
                        last_shot = DL_DIR / f"formations_error_{datetime.now():%Y%m%d_%H%M%S}.png"
                        pg.screenshot(path=str(last_shot), full_page=True)
                    except Exception:
                        last_shot = None

        notify_failure(TASK_NAME, last_err or RuntimeError("Onbekende fout"), last_shot, "\n".join(logbuf))
        ctx.close(); browser.close()
        raise SystemExit(1)

if __name__ == "__main__":
    run()
