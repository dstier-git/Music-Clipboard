import os
import platform
from pathlib import Path

IS_WINDOWS = os.name == "nt"
IS_MACOS = platform.system() == "Darwin"


def project_root():
    return Path(__file__).resolve().parents[1]


def output_dirs():
    root = project_root()
    return root / "txts", root / "midis"


def default_hotkey():
    return "ctrl+cmd+s" if IS_MACOS else "ctrl+alt+s"


def save_selection_shortcut_label():
    return "Cmd+Shift+S" if IS_MACOS else "Ctrl+Shift+S"
