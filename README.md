# MuseScore 4 Plugin: Note Name Above

This plugin adds a textbox above selected notes displaying the note name (e.g., C, C#, D, Eb).

## Installation

1. Copy `NoteNameAbove.qml` to your MuseScore 4 plugins folder:
   - **Windows**: `C:\Users\[Your User Name]\Documents\MuseScore4\Plugins\`
   - **macOS**: `~/Documents/MuseScore4/Plugins/`
   - **Linux**: `~/Documents/MuseScore4/Plugins/`

2. Enable the plugin:
   - Open MuseScore 4
   - Go to `Plugins` → `Manage Plugins...`
   - Find "Note Name Above" and check the box to enable it

## Usage

1. Open a score in MuseScore 4
2. Select one or more notes
3. Go to `Plugins` → `Note Name Above`
4. The note names will appear above the selected notes

## Debugging

Since MuseScore 4 doesn't show console.log output in the terminal, check the MuseScore log files:

**Windows**: `%LOCALAPPDATA%\MuseScore\MuseScore4\logs\`

Look for the most recent log file and check for any plugin-related errors.

## Testing

If the main plugin doesn't work, try `NoteNameAbove_Test.qml` which processes only the first selected note. This can help identify if the issue is with processing multiple notes.

## Troubleshooting

- **Plugin doesn't appear in menu**: Make sure the plugin file is in the correct Plugins folder and is enabled in Plugin Manager
- **Nothing happens when running**: 
  - Make sure you have notes selected (not just measures or other elements)
  - Check the MuseScore log files for errors
  - Try the test version first

