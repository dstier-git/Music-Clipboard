import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

try:
    import keyboard
except ImportError:
    keyboard = None

try:
    import psutil
except ImportError:
    psutil = None

APP_SCRIPT = Path(__file__).resolve().parent / "musescore_extractor_gui.py"
REQUEST_FILE = Path(tempfile.gettempdir()) / "musescore_hotkey_request.txt"
HOTKEY = "ctrl+alt+s"


def _get_interpreter():
    interpreter = Path(sys.executable)
    pythonw = interpreter.with_name("pythonw.exe")
    if os.name == "nt" and pythonw.exists():
        return pythonw
    return interpreter


def _is_gui_running():
    if psutil is None:
        return False

    for proc in psutil.process_iter(["cmdline"]):
        cmdline = proc.info.get("cmdline") or []
        for part in cmdline:
            if not part:
                continue
            if APP_SCRIPT.name in os.path.basename(part):
                return True
    return False


def _start_gui():
    interpreter = _get_interpreter()
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
    subprocess.Popen(
        [str(interpreter), str(APP_SCRIPT), "--trigger-save-selection", "--disable-global-hotkey"],
        creationflags=creation_flags,
    )


def _signal_gui():
    try:
        REQUEST_FILE.parent.mkdir(parents=True, exist_ok=True)
        REQUEST_FILE.write_text(str(time.time()))
    except Exception:
        pass


def _on_hotkey():
    if _is_gui_running():
        _signal_gui()
    else:
        _start_gui()


def main():
    if keyboard is None:
        raise SystemExit(
            "The 'keyboard' library is required to run hotkey_listener.py. "
            "Install it with: pip install keyboard"
        )

    if psutil is None:
        print("Warning: psutil is not installed. The listener will always restart the GUI instead of signaling the running one.")

    # Touch the marker file so the GUI's watcher has a baseline timestamp.
    try:
        REQUEST_FILE.parent.mkdir(parents=True, exist_ok=True)
        REQUEST_FILE.write_text(str(time.time()))
    except Exception:
        pass

    keyboard.add_hotkey(HOTKEY, _on_hotkey)
    print(f"Listening for Ctrl+Alt+S → launches {APP_SCRIPT.name}")

    keyboard.wait()


if __name__ == "__main__":
    main()

