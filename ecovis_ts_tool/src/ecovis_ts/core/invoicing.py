import logging
from pathlib import Path
from ecovis_ts.utils.paths import ts_root, output_root


def generate_invoice_annex(month: str, client_codes: list):
    """
    Generates billing annexes for specific clients.
    Refactored from generate_szamlamelleklet.py.
    """
    logger = logging.getLogger(__name__)
    codes_str = ", ".join(client_codes) if client_codes else "alapértelmezett"
    logger.info(f"Indítás: Számlamelléklet generálás ({codes_str})")

    # ... logic to filter aggregated data by client codes and format Excel ...

    output_path = output_root() / f"szamlamelleklet_{month}.xlsx"
    logger.info(f"✅ Számlamelléklet elkészült: {output_path.name}")
    return output_path
