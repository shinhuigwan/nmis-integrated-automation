"""Keep writable user files beside the portable executable, resources in bundle."""
import sys
from pathlib import Path

APP_ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
RESOURCE_ROOT = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
