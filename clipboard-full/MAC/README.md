# MuseScore Pitch Extractor (macOS)

A desktop application for extracting pitches and metric positions from MuseScore files (.mscx/.mscz).

## Features

- **Manual Mode**: Select and process any MuseScore file
- **Auto Mode**: Automatically process files saved via MuseScore's "Save Selection" feature
- **Automated Save Selection**: Trigger MuseScore's "Save Selection" dialog with a single click (no manual menu navigation needed!)
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
   pip install pyautogui
   ```
   
   Note: The automation feature (triggering Save Selection) requires `pyautogui`. The app will work without it, but the automation button will be disabled.

## Usage

### Desktop GUI Application

1. Run the GUI:
   ```bash
   python3 musescore_extractor_gui.py
   ```
   
   Or use the provided script:
   ```bash
   chmod +x run_gui.sh
   ./run_gui.sh
   ```

2. **Manual Mode**:
   - Click "Browse..." to select a MuseScore file
   - Click "Extract" to process it
   - Results will appear in the output area and be saved to a text file

3. **Auto Mode (Save Selection Workflow)** - Recommended for measure selection:
   - Set the watch folder (default: `Documents/MuseScore4/Scores`)
   - Click "Start Watching"
   - In MuseScore:
     - **Select the measures** you want to extract (click and drag to select)
     - **Option A (Automated)**: Click the "Trigger Save Selection in MuseScore" button in the app
     - **Option B (Manual)**: Go to **File → Save Selection** (or press `Cmd+Shift+S`)
     - Complete the save dialog in MuseScore (choose location and filename)
     - Save the selection to the watched folder
   - The app will automatically detect and process the new file
   - Results appear in real-time in the output area
   - **Note**: Saved selections already contain only the selected measures, so no need to specify a range

### Command-Line Scripts

**Extract pitches only:**
```bash
python3 extract_pitches.py
```

**Extract pitches with positions:**
```bash
python3 extract_pitches_with_position.py
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

## Output Location

All extracted files are saved to the `txts` folder in the project root directory (one level up from the MAC folder).

## Tips

- **For selecting specific measures**: Use MuseScore's "Save Selection" feature - it's the easiest way to extract pitches from a subset of measures
- **For processing full files**: Use manual mode to extract the entire score
- The watch folder feature is perfect for quickly processing selections from MuseScore
- The GUI shows the first 10 extracted notes in the output area
- You can clear the output area at any time
- When you save a selection in MuseScore, the saved file only contains those measures, so the app will extract all measures from that file

## Troubleshooting

- **"Could not import extraction scripts"**: Make sure `extract_pitches_with_position.py` is in the same directory
- **No notes extracted**: Check that the file is a valid MuseScore file (.mscx or .mscz)
- **Watch folder not working**: Ensure MuseScore is saving to the watched folder location
- **"Trigger Save Selection" button is disabled**: Install the required dependencies: `pip install pyautogui`
- **"MuseScore Not Found" error**: Make sure MuseScore 4 is running and has at least one score open before clicking the automation button
- **Keyboard shortcut not working**: The automation feature requires MuseScore to be the active window. If it doesn't work, try manually activating MuseScore first, then click the button
- **Permission errors with AppleScript**: macOS may require accessibility permissions. Go to System Preferences → Security & Privacy → Privacy → Accessibility and ensure Terminal (or your Python IDE) has permission

## macOS-Specific Notes

- The automation feature uses AppleScript to find and activate MuseScore windows
- Keyboard shortcuts use `Cmd+Shift+S` instead of `Ctrl+Shift+S`
- The app automatically detects macOS and uses the appropriate automation methods
- If you encounter permission issues with AppleScript, you may need to grant accessibility permissions to Terminal or your Python environment

