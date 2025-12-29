# MuseScore Pitch Extractor

A desktop application for extracting pitches and metric positions from MuseScore files (.mscx/.mscz).

## Features

- **Manual Mode**: Select and process any MuseScore file
- **Auto Mode**: Automatically process files saved via MuseScore's "Save Selection" feature
- **Extracts**: Pitch names (e.g., C4, E5) and metric positions (Measure:Beat format)
- **User-friendly GUI**: Simple interface with real-time output

## Files

- `extract_pitches.py` - Command-line script for extracting pitches only
- `extract_pitches_with_position.py` - Command-line script for extracting pitches with metric positions
- `musescore_extractor_gui.py` - Desktop GUI application

## Installation

No installation required! Just ensure you have Python 3.6+ installed.

Required Python libraries (usually included):
- `tkinter` (GUI)
- `xml.etree.ElementTree` (XML parsing)
- `zipfile` (for .mscz files)

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
     - Go to **File → Save Selection** (or press `Ctrl+Shift+S`)
     - Save the selection to the watched folder
   - The app will automatically detect and process the new file
   - Results appear in real-time in the output area
   - **Note**: Saved selections already contain only the selected measures, so no need to specify a range

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

