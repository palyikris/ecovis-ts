import sys
from pathlib import Path
from src.ecovis_ts.ui.app import EcovisApp

# Ensure the 'src' directory is in the Python path
src_path = str(Path(__file__).parent / "src")
if src_path not in sys.path:
    sys.path.insert(0, src_path)


if __name__ == "__main__":
    app = EcovisApp()
    app.mainloop()
