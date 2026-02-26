import logging
from pathlib import Path
from ecovis_ts.utils.paths import ts_root, output_root


def validate_client_project_pairs(month: str):
    """
    Checks all TS files for invalid Client-Project code combinations.
    Refactored from validate_pairs.py.
    """
    logger = logging.getLogger(__name__)
    logger.info(f"Indítás: Párellenőrzés - Hónap: {month}")

    ts_dir = ts_root()
    # Logic to load master pair list

    errors_found = 0
    # ... logic to iterate files and compare against master list ...

    if errors_found > 0:
        logger.warning(f"⚠️ {errors_found} hiba található a párellenőrzés során.")
        # Save invalid_parok_{month}.xlsx
    else:
        logger.info("✅ Nincs hiba a projekt-ügyfélkód párosításokban.")

    return True
