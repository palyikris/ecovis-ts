import xlwings as xw
import pandas as pd
import logging
from ..utils.paths import ts_root


def sync_dropdown_lists():
    logger = logging.getLogger(__name__)
    logger.info("▶ Legördülők frissítése minden TS fájlban...")

    ts_dir = ts_root()
    master_path = ts_dir / "Ecovis Compliance Solution számlázási adatok_2025.xlsx"

    try:
        # Load master data
        master_df = pd.read_excel(master_path, sheet_name="TS kódok")
        u_list = master_df["Ügyfélkód"].unique().tolist()
        p_list = master_df["TS kód"].unique().tolist()

        app = xw.App(visible=False, add_book=False)
        for ts_file in ts_dir.glob("*.xlsx"):
            if "TS" not in ts_file.name or ts_file.name.startswith("~$"):
                continue
            if ts_file.name == master_path.name:
                continue

            logger.info(f"  - Frissítés: {ts_file.name}")
            wb = app.books.open(ts_file)

            # Update the hidden Lists sheet or static columns
            # Assuming logic from original update_dropdowns.py
            sheet = wb.sheets[0]
            sheet.range("Y2:Y500").clear_contents()
            sheet.range("Z2:Z500").clear_contents()

            sheet.range("Y2").options(transpose=True).value = u_list
            sheet.range("Z2").options(transpose=True).value = p_list

            wb.save()
            wb.close()

        app.quit()
        logger.info("✅ Minden legördülő lista frissítve.")
        return True
    except Exception as e:
        logger.error(f"Hiba a frissítés során: {e}")
        return False
