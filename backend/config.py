import os
import sys
from pathlib import Path


def get_env() -> str:
    return os.getenv("APP_ENV", "dev").lower()


def get_data_dir() -> Path:
    # When compiled with PyInstaller, store data next to the .exe so it persists.
    # sys._MEIPASS is the temp extraction folder — writing there would be lost on exit.
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).resolve().parent.parent
    data_dir = base / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir


def data_file(filename: str) -> Path:
    env = get_env()
    data_dir = get_data_dir()
    stem, dot, suffix = filename.partition(".")
    env_filename = f"{stem}_{env}"
    if suffix:
        env_filename = f"{env_filename}.{suffix}"
    return data_dir / env_filename


def get_template_dir() -> Path:
    """Read-only template files — inside the bundle when frozen, project root otherwise."""
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / "template"
    return Path(__file__).resolve().parent.parent / "template"


def get_eps_template_path() -> Path:
    """Return path to the EPS blank template CSV (read-only reference)."""
    import glob as _glob
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent.parent
    matches = _glob.glob(str(base / "EPS*Blank*.csv"))
    if matches:
        return Path(matches[0])
    return base / "EPS  Blank 2026.csv"


def get_output_dir() -> Path:
    """Writable output directory — next to the exe when frozen, project template/output otherwise."""
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).parent / "output"
    else:
        base = Path(__file__).resolve().parent.parent / "template" / "output"
    base.mkdir(parents=True, exist_ok=True)
    return base
