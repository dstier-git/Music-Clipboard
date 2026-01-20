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
HOTKEY = "ctrl+cmd+s"


def _get_interpreter():
    interpreter = Path(sys.executable)
    # On macOS, we can use the same interpreter
    return interpreter


def _is_gui_running():
    if psutil is None:
        return False

    for proc in psutil.process_iter(["cmdline", "name"]):
        try:
            cmdline = proc.info.get("cmdline") or []
            name = proc.info.get("name") or ""
            # Check both command line and process name
            for part in cmdline:
                if not part:
                    continue
                if APP_SCRIPT.name in os.path.basename(part):
                    return True
            # Also check process name (on macOS, Python processes might show as "Python")
            if "python" in name.lower() and any(APP_SCRIPT.name in str(arg) for arg in cmdline):
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False


def _start_gui():
    interpreter = _get_interpreter()
    # On macOS, we can run in background without showing terminal
    subprocess.Popen(
        [str(interpreter), str(APP_SCRIPT), "--trigger-save-selection", "--disable-global-hotkey"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
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
    print(f"Listening for Ctrl+Cmd+S → launches {APP_SCRIPT.name}")

    keyboard.wait()


if __name__ == "__main__":
    main()


