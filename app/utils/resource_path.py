import sys
from pathlib import Path


def get_base_path() -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parents[2]


def get_resource_path(relative_path: Path) -> Path:
    return get_base_path() / relative_path
