import pandas as pd
import logging
from datetime import datetime
from openpyxl import Workbook
from ..utils.paths import ts_root, output_root
from .helpers import norm_header, write_table, add_title_banner, autosize_columns


def aggregate_timesheets(month: str):
    logger = logging.getLogger(__name__)
    logger.info(f"▶ Indítás: Összesített idők - Hónap: {month}")

    ts_dir = ts_root()
    output_dir = output_root()

    # 1. Load active clients
    try:
        ceg_path = ts_dir / "Ecovis Compliance Solution számlázási adatok_2025.xlsx"
        ceg = pd.read_excel(ceg_path, sheet_name="Cégadatok")
        active_clients = set(
            ceg[ceg["Ügyfél aktív"].astype("str").str.strip().str.lower() == "igen"][
                "Ügyfélkód"
            ].astype("str")
        )
    except Exception as e:
        logger.error(f"Hiba az aktív ügyfelek betöltésekor: {e}")
        return None

    # 2. Process Files
    records = []
    files = [
        p
        for p in ts_dir.glob("*.xlsx")
        if "TS" in p.name and not p.name.startswith("~$")
    ]

    for file_path in files:
        name_part = file_path.stem.replace("TS ", "")
        logger.info(f"Feldolgozás: {file_path.name}")

        try:
            xls = pd.ExcelFile(file_path)
            month_sheet = next(
                (s for s in xls.sheet_names if norm_header(s) == norm_header(month)),
                None,
            )

            if not month_sheet:
                logger.warning(
                    f"  - Nem található '{month}' munkalap a(z) {file_path.name} fájlban."
                )
                continue

            df = pd.read_excel(xls, sheet_name=month_sheet)
            df.columns = [str(c).strip() for c in df.columns]

            # Filter rows with data
            mask = (
                df["Ügyfélkód"].notna()
                & df["Projektkód"].notna()
                & (df["Időtartam (óra)"] > 0)
            )
            valid_df = df[mask]

            for _, row in valid_df.iterrows():
                u_kod = str(row["Ügyfélkód"]).strip()
                if u_kod in active_clients:
                    records.append(
                        {
                            "Munkatárs": name_part,
                            "Ügyfélkód": u_kod,
                            "Projektkód": str(row["Projektkód"]).strip(),
                            "Feladat részletezése": str(
                                row.get("Feladat részletezése", "")
                            ),
                            "Dátum": row["Dátum"],
                            "Óra": row["Időtartam (óra)"],
                        }
                    )
        except Exception as e:
            logger.error(f"  - Hiba a(z) {file_path.name} feldolgozásakor: {e}")

    if not records:
        logger.error("Nem találtam adatot a megadott hónapra.")
        return None

    # 3. Create Result Workbook
    full_df = pd.DataFrame(records)
    summary_df = (
        full_df.groupby(["Ügyfélkód", "Projektkód", "Munkatárs"])["Óra"]
        .sum()
        .reset_index()
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Összesítés"

    add_title_banner(ws, "Havi Összesített Idők", month, col_count=4)
    write_table(ws, summary_df, start_row=3)
    autosize_columns(ws)

    filename = (
        f"timesheet_summary_{month}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    )
    save_path = output_dir / filename
    wb.save(save_path)

    logger.info(f"✅ Kész! Mentve: {save_path.name}")
    return save_path
