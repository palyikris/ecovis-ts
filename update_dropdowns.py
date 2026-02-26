# -*- coding: utf-8 -*-
import xlwings as xw
import pandas as pd
import os
import unicodedata
from datetime import datetime
import logging
from pathlib import Path
import time
import sys

# =========================
# Config
# =========================
FOLDER_PATH = "."
ECOVIS_PATH = "Ecovis Compliance Solution számlázási adatok_2025.xlsx"
TS_KODOK_SHEET = "TS kódok"

# =========================
# Logging setup (UTF-8)
# =========================
LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / f"update_dropdowns_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logging.info("▶ update_dropdowns started")
logging.info(f"Log file: {LOG_FILE.resolve()}")


# =========================
# Helpers
# =========================
def remove_accents(s: str) -> str:
    nfkd = unicodedata.normalize("NFKD", str(s))
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower().strip()


# Hónapnevek (ékezet nélkül)
HONAPOK = [
    "januar",
    "februar",
    "marcius",
    "aprilis",
    "majus",
    "junius",
    "julius",
    "augusztus",
    "szeptember",
    "oktober",
    "november",
    "december",
]
cur_idx = datetime.now().month - 1
TARGET_MONTHS = HONAPOK[cur_idx:]
logging.info(f"Csak ezek a hónapok frissülnek: {', '.join(TARGET_MONTHS)}")

# =========================
# Load Ecovis data once
# =========================
try:
    ecovis_df = pd.read_excel(ECOVIS_PATH, sheet_name=TS_KODOK_SHEET)
    ceg = pd.read_excel(ECOVIS_PATH, sheet_name="Cégadatok")
    active_clients = set(
        ceg[ceg["Ügyfél aktív"].astype(str).str.strip().str.lower() == "igen"]["Ügyfélkód"].astype(str)
    )

    ugyfelkodok = sorted(
        [
            x
            for x in ecovis_df["Ügyfélkód"].dropna().astype(str).unique()
            if x in active_clients
        ],
        key=remove_accents,
    )

    projektnevek = sorted(
        ecovis_df["Projekt neve"].dropna().astype(str).unique(), key=remove_accents
    )
    logging.info(
        f"Loaded TS kódok: {len(ugyfelkodok)} ügyfélkód, {len(projektnevek)} projekt"
    )
except Exception as e:
    logging.exception(f"❌ Nem sikerült betölteni a TS kódok adatot: {e}")
    raise

# =========================
# Counters
# =========================
start_time = time.time()
processed = 0
skipped = 0
errors = 0
skipped_workers: list[str] = []

# =========================
# Main
# =========================
# FONTOS: saját App példány kezelése, hogy ne maradjanak üres EXCEL.EXE-k
app = None
try:
    # add_book=False => NEM nyit “Book1”-et; visible=False => nem villog a GUI
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    for file in os.listdir(FOLDER_PATH):
        if not (file.endswith(".xlsx") and "TS" in file and not file.startswith("~$")):
            continue

        file_path = os.path.join(FOLDER_PATH, file)
        logging.info(f"🔧 Feldolgozás: {file}")

        wb = None
        try:
            # Mindig az általunk kezelt app-ban nyissunk!
            wb = app.books.open(file_path, update_links=False, read_only=False)

            for ws in wb.sheets:
                sheet_norm = remove_accents(ws.name)
                if sheet_norm not in TARGET_MONTHS:
                    continue

                logging.info(f"  ➔ Sheet: {ws.name}")

                # 0) Inicializálás: üres cellák kitöltése, hogy Validation ne akadjon fenn
                #    (gyorsabb blokkonként írni, de hagyjuk egyszerűen és stabilan)
                for row in range(2, 302):
                    for col in ("D", "E"):
                        rng = ws.range(f"{col}{row}")
                        if rng.value is None:
                            rng.value = ""

                # 1) Segédoszlopok ürítése + feltöltése (Y: ügyfélkódok, Z: projektek)
                ws.range("Y2:Y1000").clear_contents()
                ws.range("Z2:Z1000").clear_contents()
                ws.range("Y2").options(transpose=True).value = ugyfelkodok
                ws.range("Z2").options(transpose=True).value = projektnevek

                # 2) Tartomány képletek a validációhoz
                client_formula = f"=$Y$2:$Y${1 + len(ugyfelkodok)}"
                project_formula = f"=$Z$2:$Z${1 + len(projektnevek)}"

                # 3) Data validation a D és E oszlopokra (2..301)
                d_block = ws.range("D2:D301").api
                e_block = ws.range("E2:E301").api

                # Töröljük a meglévő validációkat (ha lennének)
                try:
                    d_block.Validation.Delete()
                except Exception:
                    pass
                try:
                    e_block.Validation.Delete()
                except Exception:
                    pass

                # Add: Type=3 (xlValidateList), AlertStyle=1 (Stop), Operator=1 (Between)
                d_block.Validation.Add(3, 1, 1, client_formula)
                e_block.Validation.Add(3, 1, 1, project_formula)

            wb.save()
            wb.close()
            processed += 1
            logging.info(f"✅ Kész: {file}")

        except Exception as e:
            errors += 1
            logging.exception(f"❌ Hiba feldolgozás közben: {file} — {e}")
            try:
                if wb is not None:
                    wb.close()
            except Exception:
                pass
            # ha nyitási hiba volt (pl. megnyitás írásvédetten), számoljuk kihagyásnak is
            if "wb is None" or "Cannot open" in str(e):
                skipped += 1
                skipped_workers.append(file)

    # Skip report
    if skipped_workers:
        try:
            with open("update_dropdowns_logs.txt", "w", encoding="utf-8") as f:
                for worker in skipped_workers:
                    f.write(worker + "\n")
                f.write(f"\nGenerálva: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
            logging.warning(
                "A kihagyott fájlok listája elmentve: update_dropdowns_logs.txt"
            )
        except Exception as e:
            logging.exception(f"Nem sikerült kiírni a kihagyott fájlok listáját: {e}")

except Exception as top_e:
    errors += 1
    logging.exception(f"❌ Váratlan hiba: {top_e}")
finally:
    # Mindig zárjuk le az általunk indított App-ot, különben ott marad az EXCEL.EXE
    try:
        if app is not None:
            # Zárjuk, ha véletlen maradt volna nyitott munkafüzet
            for b in list(app.books):
                try:
                    b.close()
                except Exception:
                    pass
            app.quit()
    except Exception:
        # végső fallback
        pass

    # Summary
    duration = time.time() - start_time
    logging.info("📊 Run summary:")
    logging.info(f"   ✔ {processed} files processed")
    logging.info(f"   ⚠ {skipped} skipped")
    logging.info(f"   ❌ {errors} errors")
    logging.info(f"   ⏱ Duration: {duration:.1f}s")
    logging.info("✅ update_dropdowns finished")
