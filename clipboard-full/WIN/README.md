# MuseScore Pitch Extractor

A desktop application for extracting pitches and metric positions from MuseScore files (.mscx/.mscz).

## Features

- **Manual Mode**: Select and process any MuseScore file
- **Auto Mode**: Automatically process files saved via MuseScore's "Save Selection" feature
- **Automated Save Selection**: Trigger MuseScore's "Save Selection" dialog with a single click or the background hotkey listener (Ctrl+Alt+S) - works even when the GUI is closed!
- **Extracts**: Pitch names (e.g., C4, E5) and metric positions (Measure:Beat format)
- **User-friendly GUI**: Simple interface with real-time output

## Files

- `extract_pitches.py` - Command-line script for extracting pitches only
- `extract_pitches_with_position.py` - Command-line script for extracting pitches with metric positions
- `musescore_extractor_gui.py` - Desktop GUI application

## Installation

1. Ensure you have Python 3.6+ installed.

2. **Core dependencies** (usually included with Python):
   - `tkinter` (GUI)
   - `xml.etree.ElementTree` (XML parsing)
   - `zipfile` (for .mscz files)

3. **Optional dependencies** (for automation features):
   ```bash
   pip install -r requirements.txt
   ```
   
   Or install individually:
   ```bash
   pip install pywinauto pyautogui keyboard
   ```
   
   Note: The automation feature (triggering Save Selection) requires `pywinauto` and `pyautogui`. The global hotkey feature requires `keyboard`. The app will work without them, but the automation features will be disabled.

## Usage

### Desktop GUI Application

1. Run the GUI:
   ```bash
   python musescore_extractor_gui.py
   ```

2. **Manual Mode**:
   - Click "Browse..." to select a MuseScore file
   - **Optional**: Check "Extract only specific measures" and enter a measure range (e.g., measures 5-10)
   - Click "Extract" to process it
   - Results will appear in the output area and be saved to a text file

3. **Auto Mode (Save Selection Workflow)** - Recommended for measure selection:
   - Set the watch folder (default: `Documents/MuseScore4/Scores`)
   - Click "Start Watching"
   - In MuseScore:
     - **Select the measures** you want to extract (click and drag to select)
     - **Option A (Global Hotkey - Recommended)**: Press `Ctrl+Alt+S` from anywhere (even when the GUI is closed and another app is focused)
     - **Option B (Button)**: Click the "Trigger Save Selection in MuseScore" button in the app
     - **Option C (Manual)**: Go to **File → Save Selection** (or press `Ctrl+Shift+S` in MuseScore)
     - Complete the save dialog in MuseScore (choose location and filename)
     - Save the selection to the watched folder
   - The app will automatically detect and process the new file
   - Results appear in real-time in the output area
   - The last watch folder and watching state are remembered so the workflow resumes after restarting the app
   - **Note**: Saved selections already contain only the selected measures, so no need to specify a range

### Background global hotkey listener

1. Install the optional automation requirements (see `requirements.txt`) so `keyboard` and `psutil` are available.
2. Run `python hotkey_listener.py` or double-click `run_hotkey_listener.bat`. Once running it will register `Ctrl+Alt+S` globally.
3. The listener writes to `musescore_hotkey_request.txt` in your OS temporary directory so the GUI can detect the request the next time it runs.
4. When `Ctrl+Alt+S` is pressed:
   - If the GUI is already running, it receives the notification and triggers the Save Selection workflow.
   - If the GUI is not running, the listener launches it with `--trigger-save-selection`, so the app opens and immediately starts the automation.
5. Add a shortcut to `run_hotkey_listener.bat` inside `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup` if you want the listener to start automatically on login.

### Command-Line Scripts

**Extract pitches only:**
```bash
python extract_pitches.py
```

**Extract pitches with positions:**
```bash
python extract_pitches_with_position.py
```

Both scripts will prompt you to enter a file path.

## Output Format

The output text file contains:
- **Pitch**: Note name (e.g., C4, E5, F#3)
- **Position**: Measure and beat (e.g., M1:1.00, M2:2.50)
- **Tick**: Internal timing value (for reference)

Example:
```
C4	M1:1.00	(tick: 0)
E4	M1:1.00	(tick: 0)
G4	M1:2.00	(tick: 480)
```

## Tips

- **For selecting specific measures**: Use MuseScore's "Save Selection" feature - it's the easiest way to extract pitches from specific measures
- **For processing full files**: Use manual mode and optionally specify a measure range
- The watch folder feature is perfect for quickly processing selections from MuseScore
- Output files are saved in the same directory as the input file
- The GUI shows the first 10 extracted notes in the output area
- You can clear the output area at any time
- When you save a selection in MuseScore, the saved file only contains those measures, so the app will extract all measures from that file

## Troubleshooting

- **"Could not import extraction scripts"**: Make sure `extract_pitches_with_position.py` is in the same directory
- **No notes extracted**: Check that the file is a valid MuseScore file (.mscx or .mscz)
- **Watch folder not working**: Ensure MuseScore is saving to the watched folder location
- **"Trigger Save Selection" button is disabled**: Install the required dependencies: `pip install pywinauto pyautogui`
- **Global hotkey not working**: Install the `keyboard` library: `pip install keyboard`. Note: On Windows, you may need to run the app as Administrator for global hotkeys to work properly.
- **Global hotkey doesn't open the GUI when it's closed**: Make sure the background helper (`run_hotkey_listener.bat`) is running (or has been added to your Startup folder) so `Ctrl+Alt+S` can launch the app and trigger Save Selection.
- **"MuseScore Not Found" error**: Make sure MuseScore 4 is running and has at least one score open before using the automation feature
- **Keyboard shortcut not working**: The automation feature requires MuseScore to be running. The global hotkey (Ctrl+Alt+S) will automatically find and activate MuseScore, so you don't need to manually switch to it

