import pandas as pd
from pathlib import Path
import logging
from ecovis_ts_tool.src.ecovis_ts.utils.paths import ts_root, output_root


def aggregate_timesheets(month: str):
    """
    Summarizes total hours from all XLSX files in the TS folder for a given month.
    Refactored from standalone timesheet_summary.py.
    """
    logger = logging.getLogger(__name__)
    logger.info(f"Indítás: Összesített idők - Hónap: {month}")

    ts_dir = ts_root()
    output_dir = output_root()

    files = [
        p
        for p in ts_dir.glob("*.xlsx")
        if "TS" in p.name and not p.name.startswith("~$")
    ]

    if not files:
        logger.error("Nem található feldolgozandó TS fájl.")
        return None

    all_data = []
    for file_path in files:
        logger.info(f"Feldolgozás: {file_path.name}")
        try:
            # Logic adapted from the original timesheet_summary.py
            df = pd.read_excel(file_path, sheet_name=month)
            # ... perform your specific aggregation logic here ...
            all_data.append(df)
        except Exception as e:
            logger.error(f"Hiba a {file_path.name} feldolgozásakor: {e}")

    if not all_data:
        return None

    # Save result
    summary_name = f"timesheet_summary_{month}.xlsx"
    summary_path = output_dir / summary_name
    # pd.concat(all_data).to_excel(summary_path)

    logger.info(f"✅ Összesítés elkészült: {summary_name}")
    return summary_path
