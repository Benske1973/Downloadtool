from pathlib import Path
import os

# Basispad van de gesynchroniseerde SharePoint-bibliotheek.
# Geen hardcoded gebruikersnaam: werkt voor elke Windows-gebruiker met EQUANS-sync.
BASE = Path.home() / "EQUANS" / "Operations Support - Trainingapp"

# Map waar alle Xaurum-downloadscripts hun bestanden wegschrijven.
DL_DIR = BASE / "XaurumTools" / "downloads"

# Locatie waar de login/auth-sessie (Playwright storage_state) wordt bijgehouden.
# Deze mag gerust gebruikers-specifiek in APPDATA staan.
AUTH_STATE = Path(
    os.environ.get("APPDATA", str(Path.home()))
) / "XaurumUploader" / "xaurum_auth_state.json"


def ensure_download_dir() -> Path:
    """
    Zorg dat de downloadmap bestaat en geef hem terug.
    Handig als helper voor scripts.
    """
    DL_DIR.mkdir(parents=True, exist_ok=True)
    return DL_DIR

