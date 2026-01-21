from __future__ import annotations
import re
import csv
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright
from xaurum_common import (
    DL_DIR, AUTH_STATE, should_run_headless,
    robust_goto, ensure_logged_in, close_popups_everywhere,
    notify_failure,
)

TASK_NAME = "Dispatcher – Training catalogus"
BASE_URL = "https://equans.xaurum.be/nl/offer"


def parse_code_and_name(full: str) -> tuple[str | None, str]:
    """
    'EA-S-012 - Hulpverlener - Refresh'
         -> ('EA-S-012', 'Hulpverlener - Refresh')

    Als het patroon niet klopt, geven we None terug voor code
    en heel de tekst in 'title'.
    """
    full = full.strip()
    m = re.match(r"\s*([^-]+?)\s*-\s*(.+)", full)
    if not m:
        return None, full
    code = m.group(1).strip()
    title = m.group(2).strip()
    return code, title


def collect_programs(pg) -> dict[str, dict]:
    """
    Scrapet de huidige /nl/offer pagina en verzamelt alle
    unieke opleidingen op basis van href '/nl/training/<id>'.

    Return: dict[training_id] = { 'training_id', 'code', 'title', 'raw_text', 'url' }
    """
    programs: dict[str, dict] = {}

    # Alle links naar trainingsdetail
    links = pg.locator("a[href^='/nl/training/']")
    count = links.count()
    print(f"   → gevonden {count} links naar /nl/training/... op deze pagina")

    for i in range(count):
        el = links.nth(i)
        href = el.get_attribute("href") or ""
        txt = el.inner_text().strip()

        m_id = re.search(r"/training/(\d+)", href)
        if not m_id:
            continue
        training_id = m_id.group(1)
        url = f"https://equans.xaurum.be{href}"

        code, title = parse_code_and_name(txt)

        if training_id in programs:
            # al gezien → overslaan (we nemen de eerste)
            continue

        programs[training_id] = {
            "training_id": training_id,
            "code": code or "",
            "title": title,
            "raw_text": txt,
            "url": url,
        }

    return programs


def load_all_pages(pg) -> dict[str, dict]:
    """
    Probeert alle opleidingen in de catalogus op te halen.
    - Werkt met:
      * 1 enkele pagina
      * of 'Meer laden' / 'Volgende' knoppen

    Mogelijk moet je de teksten van de knoppen onderaan wat tweaken
    (afhankelijk van hoe Xaurum het precies toont).
    """
    all_programs: dict[str, dict] = {}

    while True:
        close_popups_everywhere(pg)

        # Verzamel op deze pagina
        page_programs = collect_programs(pg)
        before = len(all_programs)
        all_programs.update(page_programs)
        after = len(all_programs)
        print(f"   → totaal nu {after} unieke opleidingen ( +{after - before} )")

        # Probeer een 'volgende' / 'meer laden' knop te vinden
        # (PAS EVENTUEEL AAN ALS DE TEKST ANDERS IS)
        next_button = None
        for selector in [
            "button:has-text('Meer laden')",
            "button:has-text('Volgende')",
            "a:has-text('Volgende')",
        ]:
            btn = pg.locator(selector)
            if btn.count() > 0 and btn.first.is_enabled():
                next_button = btn.first
                break

        if not next_button:
            print("   → geen 'Meer laden' / 'Volgende' knop gevonden, stoppen.")
            break

        print("   → 'Meer laden/Volgende' klikken…")
        next_button.click()
        # kleine wachttijd; als Xaurum een spinner / loading token heeft,
        # kun je hier eventueel wait_loading_token(...) gebruiken
        pg.wait_for_timeout(2000)

    return all_programs


def attempt_once(pg, csv_path: Path):
    """
    1 poging om alle opleidingen op te halen en weg te schrijven naar CSV.
    """
    robust_goto(pg, BASE_URL)
    close_popups_everywhere(pg)

    all_programs = load_all_pages(pg)

    print(f"📦 Schrijf {len(all_programs)} opleidingen naar {csv_path}")
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=["training_id", "code", "title", "raw_text", "url"]
        )
        writer.writeheader()
        for p in sorted(all_programs.values(), key=lambda x: (x["code"], x["title"])):
            writer.writerow(p)

    print(f"✅ Klaar: {csv_path}")


def run():
    last_err = None
    last_shot = None
    logbuf = []

    out_file = DL_DIR / f"{datetime.now():%Y%m%d_%H%M%S}_training_catalog.csv"

    with sync_playwright() as p:
        browser = p.chromium.launch(
            channel="msedge",
            headless=should_run_headless()
        )
        ctx = browser.new_context(
            accept_downloads=False,
            storage_state=str(AUTH_STATE) if AUTH_STATE.exists() else None
        )
        pg = ctx.new_page()
        pg.set_default_timeout(120_000)
        pg.set_default_navigation_timeout(120_000)

        # login
        robust_goto(pg, BASE_URL)
        ensure_logged_in(pg, ctx)

        for attempt in range(1, 4):  # 1 + 2 retries
            try:
                logbuf.append(f"Attempt {attempt}")
                attempt_once(pg, out_file)
                # auth state bewaren
                ctx.storage_state(path=str(AUTH_STATE))
                ctx.close()
                browser.close()
                return
            except Exception as e:
                last_err = e
                print(f"⚠️  Mislukt (poging {attempt}): {e}")
                close_popups_everywhere(pg)
                try:
                    robust_goto(pg, BASE_URL)
                except Exception:
                    pass

                if attempt == 3:
                    try:
                        last_shot = DL_DIR / f"training_catalog_error_{datetime.now():%Y%m%d_%H%M%S}.png"
                        pg.screenshot(path=str(last_shot), full_page=True)
                    except Exception:
                        last_shot = None

        notify_failure(
            TASK_NAME,
            last_err or RuntimeError("Onbekende fout in catalogus-downloader"),
            last_shot,
            "\n".join(logbuf),
        )
        ctx.close()
        browser.close()
        raise SystemExit(1)


if __name__ == "__main__":
    run()
