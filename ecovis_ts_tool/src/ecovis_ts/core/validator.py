import pandas as pd
import logging
from datetime import datetime
from openpyxl import Workbook
from ..utils.paths import ts_root, output_root
from .helpers import norm_header, write_table, add_title_banner, autosize_columns


def validate_client_project_pairs(month: str):
    logger = logging.getLogger(__name__)
    logger.info(f"▶ Indítás: Párellenőrzés - Hónap: {month}")

    ts_dir = ts_root()

    # 1. Load Master Pairs
    try:
        master_path = ts_dir / "Ecovis Compliance Solution számlázási adatok_2025.xlsx"
        master_df = pd.read_excel(master_path, sheet_name="TS kódok")
        master_pairs = set(
            zip(
                master_df["Ügyfélkód"].astype(str).str.strip(),
                master_df["TS kód"].astype(str).str.strip(),
            )
        )
    except Exception as e:
        logger.error(f"Hiba a törzsadatok betöltésekor: {e}")
        return False

    # 2. Check each TS file
    errors = []
    files = [
        p
        for p in ts_dir.glob("*.xlsx")
        if "TS" in p.name and not p.name.startswith("~$")
    ]

    for file_path in files:
        logger.info(f"Ellenőrzés: {file_path.name}")
        try:
            xls = pd.ExcelFile(file_path)
            sheet = next(
                (s for s in xls.sheet_names if norm_header(s) == norm_header(month)),
                None,
            )
            if not sheet:
                continue

            df = pd.read_excel(xls, sheet_name=sheet)
            mask = df["Ügyfélkód"].notna() & df["Projektkód"].notna()

            for idx, row in df[mask].iterrows():
                u = str(row["Ügyfélkód"]).strip()
                p = str(row["Projektkód"]).strip()
                if (u, p) not in master_pairs:
                    errors.append(
                        {
                            "Fájl": file_path.name,
                            "Sor": idx + 2,
                            "Ügyfélkód": u,
                            "Projektkód": p,
                        }
                    )
        except Exception as e:
            logger.error(f"Hiba a(z) {file_path.name} fájlban: {e}")

    if not errors:
        logger.info("✅ Minden párosítás helyes.")
        return True

    # 3. Save error report
    err_df = pd.DataFrame(errors)
    wb = Workbook()
    ws = wb.active
    ws.title = "Hibás párok"
    add_title_banner(ws, "Hibás Projekt-Ügyfél párosítások", month, col_count=4)
    write_table(ws, err_df, start_row=3)
    autosize_columns(ws)

    out_path = (
        output_root() / f"hibas_parok_{month}_{datetime.now().strftime('%H%M')}.xlsx"
    )
    wb.save(out_path)
    logger.warning(f"⚠️ {len(errors)} hiba található. Lista mentve: {out_path.name}")
    return False
