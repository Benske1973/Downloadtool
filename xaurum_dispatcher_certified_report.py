from __future__ import annotations
from datetime import datetime, timedelta
from pathlib import Path
from playwright.sync_api import sync_playwright
from xaurum_common import (
    DL_DIR, AUTH_STATE, should_run_headless,
    close_popups_everywhere, robust_goto, ensure_logged_in,
    find_export_button, notify_failure
)

BASE_URL = "https://equans.xaurum.be/nl/dispatcher/report/certified-notcertified"
TASK_NAME = "Dispatcher – Certified/Not Certified"


def get_date_range_last_2_months() -> tuple[datetime, datetime]:
    """
    Bereken datumbereik:  laatste 2 maanden tot vandaag.
    """
    today = datetime.now()
    start_date = today - timedelta(days=60)
    return start_date, today


def set_date_filters(pg, start_date:  datetime, end_date: datetime):
    """
    Vul de datumfilters in en trigger filter() voor het laden van data.
    Xaurum gebruikt formaat dd/mm/yyyy.
    """
    start_str = start_date.strftime("%d/%m/%Y")
    end_str = end_date.strftime("%d/%m/%Y")
    
    print(f"📅 Datumfilters instellen:  {start_str} - {end_str}")
    
    # Wacht tot de datumvelden zichtbaar zijn
    pg.wait_for_selector("#startDate", timeout=30_000)
    pg.wait_for_selector("#endDate", timeout=30_000)
    
    # Methode 1: Direct via JavaScript - meest betrouwbaar! 
    pg.evaluate(f"""
        document.getElementById('startDate').value = '{start_str}';
        document.getElementById('endDate').value = '{end_str}';
    """)
    
    # Controleer of waarden correct zijn ingesteld
    actual_start = pg.evaluate("document.getElementById('startDate').value")
    actual_end = pg.evaluate("document.getElementById('endDate').value")
    print(f"📋 Ingestelde waarden: startDate='{actual_start}', endDate='{actual_end}'")
    
    if actual_start != start_str or actual_end != end_str:
        print(f"⚠️ Waarden niet correct ingesteld, probeer opnieuw via fill()...")
        # Fallback: via Playwright fill
        start_input = pg.locator("#startDate")
        start_input.click()
        start_input.press("Control+a")
        start_input.type(start_str, delay=50)
        
        end_input = pg.locator("#endDate")
        end_input.click()
        end_input.press("Control+a")
        end_input.type(end_str, delay=50)
    
    # Kleine pauze
    pg.wait_for_timeout(500)
    
    # Trigger filter() functie via JavaScript
    print("🔄 Filter() aanroepen...")
    pg.evaluate("filter()")
    
    # Wacht op netwerk idle (AJAX calls klaar)
    try:
        pg.wait_for_load_state("networkidle", timeout=15_000)
    except Exception:
        pass
    
    # Extra wachttijd voor tabel refresh
    pg.wait_for_timeout(3000)
    
    # Verifieer nogmaals dat datums nog steeds correct zijn NA filter
    actual_start = pg.evaluate("document.getElementById('startDate').value")
    actual_end = pg.evaluate("document.getElementById('endDate').value")
    print(f"✅ Na filter:  startDate='{actual_start}', endDate='{actual_end}'")


def attempt_once(pg, start_date: datetime, end_date: datetime) -> Path:
    robust_goto(pg, BASE_URL)
    close_popups_everywhere(pg)
    
    # Wacht tot pagina volledig geladen is
    try:
        pg.wait_for_load_state("networkidle", timeout=30_000)
    except Exception:
        pass
    
    # Wacht tot searchresults div aanwezig is
    try:
        pg.wait_for_selector("#searchresults", timeout=30_000)
    except Exception:
        pass
    
    # BELANGRIJK: Eerst datumfilters instellen VOORDAT we iets anders doen
    set_date_filters(pg, start_date, end_date)
    
    # Wacht op tabel met resultaten
    try:
        pg.wait_for_selector("#resultstable_wrapper table", timeout=30_000)
    except Exception:
        try:
            pg.wait_for_selector("table", timeout=30_000)
        except Exception:
            pass
    
    # Zoek export knop - specifiek voor deze pagina
    btn = None
    
    # Eerst proberen:  exacte Xaurum export knop
    try: 
        xaurum_btn = pg.locator("button[onclick*='exportExcel']")
        if xaurum_btn.count() > 0 and xaurum_btn.first.is_visible():
            btn = xaurum_btn.first
            print("🔘 Gevonden:  Xaurum exportExcel knop")
    except Exception:
        pass
    
    # Fallback: generieke find_export_button
    if not btn:
        btn = find_export_button(pg)
    
    if not btn:
        pg.mouse.wheel(0, 900)
        pg.wait_for_timeout(500)
        btn = find_export_button(pg)
    
    if not btn:
        raise RuntimeError("Excel-knop niet gevonden (Certified report).")
    
    print("📥 Starten met download...")
    with pg.expect_download(timeout=120_000) as di:
        btn.click()
    dl = di.value
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = DL_DIR / f"{ts}_Report_certification.xls"
    dl.save_as(str(out))
    return out


def run():
    last_err = None
    last_shot = None
    logbuf = []

    # Laatste 2 maanden
    start_date, end_date = get_date_range_last_2_months()
    print(f"📆 Rapport periode:  {start_date:%d/%m/%Y} - {end_date:%d/%m/%Y}")

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=should_run_headless())
        ctx = browser.new_context(accept_downloads=True, storage_state=str(AUTH_STATE) if AUTH_STATE.exists() else None)
        pg = ctx.new_page()
        pg.set_default_timeout(120_000)
        pg.set_default_navigation_timeout(120_000)

        robust_goto(pg, BASE_URL)
        ensure_logged_in(pg, ctx)

        for attempt in range(1, 4):
            try:
                logbuf.append(f"Attempt {attempt}")
                out = attempt_once(pg, start_date, end_date)
                print(f"✅ Bestand opgeslagen in {out}")
                if not should_run_headless():
                    try:
                        shot = DL_DIR / f"certified_report_{datetime.now():%Y%m%d_%H%M%S}.png"
                        pg.screenshot(path=str(shot), full_page=False)
                        print(f"📸 Screenshot: {shot}")
                    except Exception: 
                        pass
                ctx.storage_state(path=str(AUTH_STATE))
                ctx.close()
                browser.close()
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
                        last_shot = DL_DIR / f"certified_error_{datetime.now():%Y%m%d_%H%M%S}.png"
                        pg.screenshot(path=str(last_shot), full_page=True)
                    except Exception: 
                        last_shot = None

        notify_failure(TASK_NAME, last_err or RuntimeError("Onbekende fout"), last_shot, "\n".join(logbuf))
        ctx.close()
        browser.close()
        raise SystemExit(1)


if __name__ == "__main__": 
    run()