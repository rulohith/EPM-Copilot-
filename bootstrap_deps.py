import subprocess
import sys
from pathlib import Path


REQUIREMENTS_FILE = Path(__file__).resolve().parent / "requirements.txt"


def install_dependencies():
    try:
        import openpyxl  # test import
        return
    except ImportError:
        print("[BOOTSTRAP] Installing missing dependencies...")

    if REQUIREMENTS_FILE.exists():
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(REQUIREMENTS_FILE)])
    else:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])


if __name__ == "__main__":
    install_dependencies()