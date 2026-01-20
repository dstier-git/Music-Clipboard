import argparse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
import time
import subprocess
import sys
from pathlib import Path
import shutil
import platform
import tempfile

# Try to import automation libraries
try:
    import pyautogui
    # Disable pyautogui failsafe (the mouse moving to corner)
    pyautogui.FAILSAFE = False
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False

# Try keyboard library for global hotkeys
try:
    import keyboard
    KEYBOARD_AVAILABLE = True
except ImportError:
    KEYBOARD_AVAILABLE = False

# Try psutil for process searching
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

# Check if we're on macOS
IS_MACOS = platform.system() == 'Darwin'

# Path to persist preferences
CONFIG_FILE = Path(os.path.expanduser("~")) / ".musescore_pitch_extractor_prefs"

# Location for hotkey trigger requests from the background listener
HOTKEY_REQUEST_FILE = Path(tempfile.gettempdir()) / "musescore_hotkey_request.txt"

# Output directory for extracted files - use relative path from script location
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "txts")
MIDI_OUTPUT_DIR = os.path.join(PROJECT_ROOT, "midis")

# Import the extraction function
EXTRACTION_FUNCTION = None
MIDI_EXTRACTION_FUNCTION = None
try:
    from extract_pitches_with_position import extract_pitches_with_position_from_mscx
    EXTRACTION_FUNCTION = extract_pitches_with_position_from_mscx
    EXTRACTION_SCRIPT = "extract_pitches_with_position"
except ImportError:
    try:
        from extract_pitches import extract_pitches_from_mscx
        EXTRACTION_SCRIPT = "extract_pitches"
        # Create a wrapper to match the expected format
        def extract_pitches_with_position_from_mscx(file_path, output_file_path=None, debug=False):
            pitches = extract_pitches_from_mscx(file_path, output_file_path, debug)
            if pitches:
                # Convert to expected format (pitch, position, tick)
                return [(p, "N/A", None) for p in pitches]
            return None
        EXTRACTION_FUNCTION = extract_pitches_with_position_from_mscx
    except ImportError:
        # Will show error when app starts
        EXTRACTION_FUNCTION = None
        EXTRACTION_SCRIPT = None

# Import MIDI extraction function
try:
    from extract_midi import extract_midi_from_mscx
    MIDI_EXTRACTION_FUNCTION = extract_midi_from_mscx
except ImportError:
    MIDI_EXTRACTION_FUNCTION = None


def run_applescript(script):
    """Run an AppleScript command and return the result"""
    try:
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True,
            text=True,
            timeout=10
        )
        # Check both returncode and stderr for errors
        if result.returncode != 0:
            error_msg = result.stderr.strip() or "AppleScript returned non-zero exit code"
            return False, result.stdout.strip(), error_msg
        return True, result.stdout.strip(), result.stderr.strip()
    except subprocess.TimeoutExpired:
        return False, "", "Timeout"
    except Exception as e:
        return False, "", str(e)


def find_musescore_window_macos():
    """Find MuseScore window on macOS using AppleScript with multiple fallback methods"""
    # Try multiple approaches to find MuseScore
    scripts = [
        # Method 1: Try exact process name "mscore" (actual process name)
        '''
        tell application "System Events"
            try
                set museScoreProcess to first process whose name is "mscore"
                return name of museScoreProcess
            on error
                return ""
            end try
        end tell
        ''',
        # Method 2: Try exact process name "MuseScore 4"
        '''
        tell application "System Events"
            try
                set museScoreProcess to first process whose name is "MuseScore 4"
                return name of museScoreProcess
            on error
                return ""
            end try
        end tell
        ''',
        # Method 3: Try process name containing "MuseScore"
        '''
        tell application "System Events"
            try
                set museScoreProcess to first process whose name contains "MuseScore"
                return name of museScoreProcess
            on error
                return ""
            end try
        end tell
        ''',
        # Method 4: Try to list all processes and find MuseScore or mscore
        '''
        tell application "System Events"
            set processList to name of every process
            repeat with procName in processList
                if procName contains "MuseScore" or procName is "mscore" then
                    return procName
                end if
            end repeat
            return ""
        end tell
        ''',
    ]
    
    # Try AppleScript methods first
    for script in scripts:
        success, output, error = run_applescript(script)
        if success and output and output.strip():
            return True, output.strip(), ""
    
    # Fallback: Try using psutil if available
    if PSUTIL_AVAILABLE:
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    proc_name = proc.info.get('name') or ''
                    proc_lower = proc_name.lower()
                    if 'musescore' in proc_lower or proc_lower == 'mscore':
                        return True, proc_name, ""
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
        except Exception:
            pass
    
    return False, "", "MuseScore process not found"


def activate_musescore_window_macos():
    """Activate MuseScore window on macOS using AppleScript with multiple fallback methods"""
    # Try multiple approaches to activate MuseScore
    scripts = [
        # Method 1: Try mscore first (actual process name)
        '''
        tell application "System Events"
            try
                set museScoreProcess to first process whose name is "mscore"
                set frontmost of museScoreProcess to true
            on error
                return false
            end try
        end tell
        ''',
        # Method 2: Try MuseScore 4 application name
        '''
        tell application "MuseScore 4"
            activate
        end tell
        ''',
        # Method 3: Try MuseScore 3
        '''
        tell application "MuseScore 3"
            activate
        end tell
        ''',
        # Method 4: Try generic MuseScore
        '''
        tell application "MuseScore"
            activate
        end tell
        ''',
        # Method 5: Use System Events with exact name "mscore"
        '''
        tell application "System Events"
            try
                set museScoreProcess to first process whose name is "mscore"
                set frontmost of museScoreProcess to true
            on error
                try
                    set museScoreProcess to first process whose name is "MuseScore 4"
                    set frontmost of museScoreProcess to true
                on error
                    set museScoreProcess to first process whose name contains "MuseScore"
                    set frontmost of museScoreProcess to true
                end try
            end try
        end tell
        ''',
        # Method 6: Use System Events with name containing MuseScore
        '''
        tell application "System Events"
            set museScoreProcess to first process whose name contains "MuseScore"
            set frontmost of museScoreProcess to true
        end tell
        '''
    ]
    
    for script in scripts:
        success, output, error = run_applescript(script)
        if success:
            return True, output, error
    
    return False, "", "Could not activate MuseScore"


def send_shortcut_macos():
    """Send Cmd+Shift+S shortcut on macOS with multiple fallback methods"""
    # Try multiple process names and methods
    scripts = [
        # Method 1: Try exact process name "mscore" (actual process name)
        '''
        tell application "System Events"
            try
                tell process "mscore"
                    keystroke "s" using {command down, shift down}
                    return true
                end tell
            on error
                return false
            end try
        end tell
        ''',
        # Method 2: Try exact process name "MuseScore 4"
        '''
        tell application "System Events"
            try
                tell process "MuseScore 4"
                    keystroke "s" using {command down, shift down}
                    return true
                end tell
            on error
                return false
            end try
        end tell
        ''',
        # Method 3: Try "MuseScore 3"
        '''
        tell application "System Events"
            try
                tell process "MuseScore 3"
                    keystroke "s" using {command down, shift down}
                    return true
                end tell
            on error
                return false
            end try
        end tell
        ''',
        # Method 4: Try generic "MuseScore"
        '''
        tell application "System Events"
            try
                tell process "MuseScore"
                    keystroke "s" using {command down, shift down}
                    return true
                end tell
            on error
                return false
            end try
        end tell
        ''',
        # Method 5: Find process dynamically by name "mscore" first
        '''
        tell application "System Events"
            try
                set museScoreProcess to first process whose name is "mscore"
                tell museScoreProcess
                    keystroke "s" using {command down, shift down}
                end tell
                return true
            on error
                return false
            end try
        end tell
        ''',
        # Method 6: Find process dynamically by name containing MuseScore
        '''
        tell application "System Events"
            try
                set museScoreProcess to first process whose name contains "MuseScore"
                tell museScoreProcess
                    keystroke "s" using {command down, shift down}
                end tell
                return true
            on error
                return false
            end try
        end tell
        ''',
        # Method 7: Try with exact name match first, then fallback
        '''
        tell application "System Events"
            try
                set museScoreProcess to first process whose name is "mscore"
                tell museScoreProcess
                    keystroke "s" using {command down, shift down}
                end tell
                return true
            on error
                try
                    set museScoreProcess to first process whose name is "MuseScore 4"
                    tell museScoreProcess
                        keystroke "s" using {command down, shift down}
                    end tell
                    return true
                on error
                    try
                        set museScoreProcess to first process whose name contains "MuseScore"
                        tell museScoreProcess
                            keystroke "s" using {command down, shift down}
                        end tell
                        return true
                    on error
                        return false
                    end try
                end try
            end try
        end tell
        '''
    ]
    
    for script in scripts:
        success, output, error = run_applescript(script)
        if success:
            return True, output, error
    
    return False, "", "Could not send keyboard shortcut to MuseScore"


class MuseScoreExtractorApp:
    def __init__(self, root, trigger_on_start=False, disable_global_hotkey=False):
        self.root = root
        self.root.title("MuseScore Pitch Extractor")
        self.root.geometry("800x700")
        
        # Check if extraction function is available
        if EXTRACTION_FUNCTION is None:
            messagebox.showerror(
                "Error", 
                "Could not import extraction scripts.\n\n"
                "Please ensure one of these files is in the same directory:\n"
                "- extract_pitches_with_position.py\n"
                "- extract_pitches.py"
            )
            root.destroy()
            return
        
        self.trigger_on_start = trigger_on_start
        self.disable_global_hotkey = disable_global_hotkey

        # Variables
        self.watch_folder = tk.StringVar()
        self.watching = False
        self.watch_thread = None
        self.processed_files = set()
        self.output_format = tk.StringVar(value="text")  # "text" or "midi"
        self.last_extracted_file = None  # Store path of last extracted file
        self.preferences = self.load_preferences()
        self._hotkey_monitor_stop = threading.Event()
        self._last_hotkey_request = 0
        
        # Create UI
        self.create_widgets()
        self.apply_saved_preferences()
        self.setup_hotkey_request_monitor()
        
        # Register global keyboard shortcut
        self.register_global_hotkey()

        if self.trigger_on_start:
            self.root.after(500, self.trigger_save_selection)
        
        # Handle window close to cleanup resources and save state
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def load_preferences(self):
        """Load saved watch folder and watching flag."""
        if not CONFIG_FILE.exists():
            return {}
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                lines = [line.rstrip('\n') for line in f.readlines()]
            folder = lines[0] if lines else ""
            watching = False
            if len(lines) > 1:
                watching = lines[1].strip().lower() == "true"
            return {"watch_folder": folder, "watching": watching}
        except Exception:
            return {}

    def save_preferences(self, watching_override=None):
        """Write the current folder and watching state to disk."""
        folder = self.watch_folder.get()
        watching = self.watching if watching_override is None else watching_override
        self.preferences = {"watch_folder": folder, "watching": watching}
        try:
            CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                f.write(f"{folder}\n")
                f.write("true\n" if watching else "false\n")
        except Exception as exc:
            self.log(f"⚠ Could not save preferences: {exc}")

    def apply_saved_preferences(self):
        """Restore the saved watch folder and optional watching state."""
        default_folder = os.path.join(os.path.expanduser("~"), "Documents", "MuseScore4", "Scores")
        saved_folder = self.preferences.get("watch_folder")

        if saved_folder and os.path.exists(saved_folder):
            self.watch_folder.set(saved_folder)
        elif os.path.exists(default_folder):
            self.watch_folder.set(default_folder)
        else:
            self.watch_folder.set(os.path.join(os.path.expanduser("~"), "Documents"))

        if self.preferences.get("watching"):
            folder = self.watch_folder.get().strip()
            if folder and os.path.exists(folder):
                self.root.after(200, self.toggle_watch)
    
    def create_widgets(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="MuseScore Pitch & Position Extractor", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File Selection Section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="Select MuseScore File:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50)
        file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(file_frame, text="Browse...", command=self.browse_file).grid(row=0, column=2, padx=5)
        ttk.Button(file_frame, text="Extract", command=self.extract_file).grid(row=0, column=3, padx=5)
        
        # Output format selection
        format_frame = ttk.Frame(file_frame)
        format_frame.grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))
        ttk.Label(format_frame, text="Output Format:").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(format_frame, text="Text", variable=self.output_format, value="text").grid(row=0, column=1, padx=5)
        ttk.Radiobutton(format_frame, text="MIDI", variable=self.output_format, value="midi").grid(row=0, column=2, padx=5)
        
        # Watch Folder Section
        watch_frame = ttk.LabelFrame(main_frame, text="Auto-Process Saved Selections", padding="10")
        watch_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        watch_frame.columnconfigure(1, weight=1)
        
        ttk.Label(watch_frame, text="Watch Folder:").grid(row=0, column=0, sticky=tk.W, padx=5)
        watch_entry = ttk.Entry(watch_frame, textvariable=self.watch_folder, width=50)
        watch_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(watch_frame, text="Browse...", command=self.browse_watch_folder).grid(row=0, column=2, padx=5)
        
        # Automation button row
        automation_frame = ttk.Frame(watch_frame)
        automation_frame.grid(row=1, column=0, columnspan=3, pady=5, sticky=tk.W)
        
        self.save_selection_button = ttk.Button(
            automation_frame, 
            text="Trigger Save Selection in MuseScore", 
            command=self.trigger_save_selection
        )
        self.save_selection_button.grid(row=0, column=0, padx=5)
        
        # Global hotkey indicator
        if KEYBOARD_AVAILABLE:
            hotkey_label = ttk.Label(
                automation_frame,
                text="Global Hotkey: Ctrl+Cmd+S (background listener keeps it active even when the GUI is closed)",
                foreground="green",
                font=("Arial", 9, "bold")
            )
            hotkey_label.grid(row=0, column=1, padx=10)
        else:
            hotkey_label = ttk.Label(
                automation_frame,
                text="(Install 'keyboard' for global hotkey: pip install keyboard)",
                foreground="gray",
                font=("Arial", 8)
            )
            hotkey_label.grid(row=0, column=1, padx=5)
        
        if not IS_MACOS:
            self.save_selection_button.config(state='disabled')
            ttk.Label(
                automation_frame, 
                text="(macOS only)", 
                foreground="gray", 
                font=("Arial", 8)
            ).grid(row=1, column=0, columnspan=2, padx=5, sticky=tk.W)
        elif not PYAUTOGUI_AVAILABLE:
            self.save_selection_button.config(state='disabled')
            ttk.Label(
                automation_frame, 
                text="(Install: pip install pyautogui)", 
                foreground="gray", 
                font=("Arial", 8)
            ).grid(row=1, column=0, columnspan=2, padx=5, sticky=tk.W)
        
        self.watch_button = ttk.Button(watch_frame, text="Start Watching", command=self.toggle_watch)
        self.watch_button.grid(row=2, column=0, columnspan=3, pady=10)
        
        self.watch_status_label = ttk.Label(watch_frame, text="Status: Not watching", foreground="gray")
        self.watch_status_label.grid(row=3, column=0, columnspan=3)
        
        # Instructions
        shortcut_key = "Cmd+Shift+S" if IS_MACOS else "Ctrl+Shift+S"
        instructions = f"""
Instructions:
1. Manual Mode: 
   - Select a .mscx or .mscz file
   - Optionally specify measure range (e.g., measures 5-10)
   - Click 'Extract' to process
2. Auto Mode (Save Selection): 
   - Set the watch folder (where MuseScore saves selections)
   - In MuseScore: Select the measures you want to extract
   - Click 'Trigger Save Selection in MuseScore' button (or manually: File → Save Selection)
   - Complete the save dialog in MuseScore
   - The app will automatically process the new file if watching is enabled
   - Click 'Start Watching' to begin monitoring
   - Note: Saved selections already contain only selected measures
        """
        ttk.Label(watch_frame, text=instructions.strip(), justify=tk.LEFT, foreground="gray").grid(
            row=4, column=0, columnspan=3, pady=10, sticky=tk.W
        )
        
        # Output Section
        output_frame = ttk.LabelFrame(main_frame, text="Output", padding="10")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        self.output_text = scrolledtext.ScrolledText(output_frame, height=15, width=80, wrap=tk.WORD)
        self.output_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Button frame for output actions
        button_frame = ttk.Frame(output_frame)
        button_frame.grid(row=1, column=0, pady=5)
        
        # Clear button
        ttk.Button(button_frame, text="Clear Output", command=self.clear_output).grid(row=0, column=0, padx=5)
        
        # Open file location button
        self.open_location_button = ttk.Button(
            button_frame, 
            text="Open File Location", 
            command=self.open_file_location,
            state='disabled'
        )
        self.open_location_button.grid(row=0, column=1, padx=5)
    
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select MuseScore File",
            filetypes=[("MuseScore files", "*.mscx *.mscz"), ("All files", "*.*")]
        )
        if filename:
            self.file_path_var.set(filename)
    
    def browse_watch_folder(self):
        folder = filedialog.askdirectory(title="Select Folder to Watch")
        if folder:
            self.watch_folder.set(folder)
            self.save_preferences()
    
    def log(self, message):
        """Add message to output text area"""
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_output(self):
        self.output_text.delete(1.0, tk.END)
    
    def open_file_location(self):
        """Open the file location of the last extracted file in Finder"""
        if self.last_extracted_file and os.path.exists(self.last_extracted_file):
            try:
                # On macOS, use 'open' command with -R to reveal in Finder
                subprocess.run(['open', '-R', self.last_extracted_file])
            except Exception as e:
                # Fallback: just open the folder
                try:
                    folder = os.path.dirname(self.last_extracted_file)
                    subprocess.run(['open', folder])
                except Exception as e2:
                    self.log(f"Error opening file location: {str(e2)}\n")
                    messagebox.showerror("Error", f"Could not open file location:\n{str(e2)}")
        else:
            messagebox.showwarning("Warning", "No extracted file location available.")

    def _delete_previous_extracted_file(self, new_path):
        """Remove the previously extracted file so only the latest remains."""
        previous_path = self.last_extracted_file
        if not previous_path or previous_path == new_path:
            return
        try:
            if os.path.exists(previous_path):
                os.remove(previous_path)
                self.log(f"Deleted previous extraction: {previous_path}")
        except Exception as exc:
            self.log(f"Failed to delete previous extraction ({previous_path}): {exc}")

    def _handle_successful_extraction(self, extracted_path):
        """Record the new extraction and update the UI."""
        if not extracted_path:
            return
        self._delete_previous_extracted_file(extracted_path)
        self.last_extracted_file = extracted_path
        self.root.after(0, lambda: self.open_location_button.config(state='normal'))
        self.root.after(0, self.open_file_location)
    
    def extract_file(self, file_path=None):
        """Extract pitches from a MuseScore file"""
        if file_path is None:
            file_path = self.file_path_var.get().strip().strip('"').strip("'")
        
        if not file_path:
            messagebox.showwarning("Warning", "Please select a file first.")
            return
        
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"File not found: {file_path}")
            return
        
        # Run extraction in a separate thread to avoid freezing UI
        thread = threading.Thread(target=self._extract_thread, args=(file_path,), daemon=True)
        thread.start()
    
    def _extract_thread(self, file_path):
        """Extraction logic running in background thread"""
        output_format = self.output_format.get()
        
        self.log(f"\n{'='*60}")
        self.log(f"Processing: {os.path.basename(file_path)}")
        self.log(f"Output format: {output_format.upper()}")
        self.log(f"{'='*60}\n")
        
        try:
            if output_format == "midi":
                # MIDI extraction
                if MIDI_EXTRACTION_FUNCTION is None:
                    error_msg = "MIDI extraction function not available. Please ensure extract_midi.py is in the same directory."
                    self.log(f"{error_msg}\n")
                    self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
                    return
                
                # Create output directory if it doesn't exist
                os.makedirs(MIDI_OUTPUT_DIR, exist_ok=True)
                
                # Determine output file path
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                output_file = os.path.join(MIDI_OUTPUT_DIR, base_name + ".mid")
                
                # Extract MIDI
                try:
                    midi_path = MIDI_EXTRACTION_FUNCTION(file_path, output_file)
                    
                    if midi_path and os.path.exists(midi_path):
                        self.log("Successfully extracted MIDI!")
                        self.log(f"MIDI file saved to: {midi_path}\n")
                        # Store the extracted file path and enable open location button
                        self._handle_successful_extraction(midi_path)
                    else:
                        error_msg = "Failed to extract MIDI file."
                        self.log(f"{error_msg}\n")
                        self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
                except Exception as e:
                    error_msg = f"Error extracting MIDI: {str(e)}"
                    self.log(f"{error_msg}\n")
                    import traceback
                    self.log(traceback.format_exc())
                    self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            else:
                # Text extraction (existing logic)
                # Create output directory if it doesn't exist
                os.makedirs(OUTPUT_DIR, exist_ok=True)
                
                # Determine output file path
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                if EXTRACTION_SCRIPT == "extract_pitches_with_position":
                    output_file = os.path.join(OUTPUT_DIR, base_name + "_pitches_with_position.txt")
                else:
                    output_file = os.path.join(OUTPUT_DIR, base_name + "_pitches.txt")
                
                # Extract pitches with positions
                result = EXTRACTION_FUNCTION(file_path, output_file, debug=False)
                
                # Handle return value - could be (notes, path) tuple or just notes
                if isinstance(result, tuple) and len(result) == 2:
                    notes, actual_output_path = result
                    output_file = actual_output_path  # Use the actual path from extraction function
                else:
                    notes = result
                
                if notes:
                    self.log(f"Successfully extracted {len(notes)} notes!")
                    self.log(f"Output saved to: {output_file}\n")
                    # Store the extracted file path and enable open location button
                    self._handle_successful_extraction(output_file)
                    
                    # Show first 10 notes
                    self.log("First 10 notes:")
                    for i, (pitch, position, tick) in enumerate(notes[:10], 1):
                        if tick is not None:
                            self.log(f"  {i}. {pitch} | {position} | (tick: {tick})")
                        else:
                            self.log(f"  {i}. {pitch} | {position}")
                    
                    if len(notes) > 10:
                        self.log(f"  ... and {len(notes) - 10} more\n")
                else:
                    self.log("No notes extracted. Please check the file format.\n")
                    self.root.after(0, lambda: messagebox.showerror("Error", "No notes were extracted from the file."))
        except Exception as e:
            error_msg = f"Error processing file: {str(e)}"
            self.log(f"{error_msg}\n")
            import traceback
            self.log(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
    
    def toggle_watch(self):
        """Start or stop watching the folder"""
        if not self.watching:
            folder = self.watch_folder.get().strip()
            if not folder or not os.path.exists(folder):
                messagebox.showerror("Error", "Please select a valid folder to watch.")
                return
            
            self.watching = True
            self.watch_button.config(text="Stop Watching")
            self.watch_status_label.config(text=f"Status: Watching '{os.path.basename(folder)}'", foreground="green")
            self.log(f"Started watching folder: {folder}\n")
            self.save_preferences()
            
            # Start watching in background thread
            self.watch_thread = threading.Thread(target=self._watch_folder, args=(folder,), daemon=True)
            self.watch_thread.start()
        else:
            self.watching = False
            self.watch_button.config(text="Start Watching")
            self.watch_status_label.config(text="Status: Not watching", foreground="gray")
            self.log("Stopped watching folder.\n")
            self.save_preferences()
    
    def _watch_folder(self, folder):
        """Monitor folder for new MuseScore files"""
        # Get initial list of files
        initial_files = set()
        for file in os.listdir(folder):
            if file.endswith(('.mscx', '.mscz')):
                full_path = os.path.join(folder, file)
                initial_files.add(full_path)
        
        self.processed_files.update(initial_files)
        
        while self.watching:
            try:
                # Check for new files
                current_files = set()
                for file in os.listdir(folder):
                    if file.endswith(('.mscx', '.mscz')):
                        full_path = os.path.join(folder, file)
                        current_files.add(full_path)
                        
                        # If it's a new file (not processed yet), process it
                        if full_path not in self.processed_files:
                            # Wait a moment to ensure file is fully written
                            time.sleep(0.5)
                            
                            # Check if file is still new (not modified recently)
                            try:
                                mod_time = os.path.getmtime(full_path)
                                if time.time() - mod_time > 1:  # File hasn't been modified in last second
                                    self.processed_files.add(full_path)
                                    self.root.after(0, lambda f=full_path: self.extract_file(f))
                                    self.log(f"Detected new file: {os.path.basename(full_path)}")
                            except OSError:
                                pass  # File might be locked, skip it
                
                # Remove files that no longer exist from processed set
                self.processed_files.intersection_update(current_files)
                
                # Sleep before next check
                time.sleep(1)
            
            except Exception as e:
                if self.watching:  # Only log if still watching
                    self.log(f"Error watching folder: {str(e)}\n")
                time.sleep(2)
    
    def trigger_save_selection(self):
        """Trigger Save Selection in MuseScore by sending keyboard shortcut"""
        if not IS_MACOS:
            messagebox.showerror(
                "Platform Error",
                "This feature is only available on macOS."
            )
            return
        
        # Run in a separate thread to avoid freezing UI
        thread = threading.Thread(target=self._trigger_save_selection_thread, daemon=True)
        thread.start()
    
    def _trigger_save_selection_thread(self):
        """Thread function to trigger Save Selection on macOS"""
        try:
            self.log("Attempting to trigger Save Selection in MuseScore...")
            self.log("Step 1: Finding MuseScore window...")
            
            # Find MuseScore using AppleScript with better error reporting
            success, output, error = find_musescore_window_macos()
            if not success:
                self.log("✗ Could not find MuseScore window")
                self.log(f"Error: {error}")
                
                # Try to list all running processes to help debug
                self.log("\nDebug: Searching for MuseScore processes...")
                found_processes = []
                
                if PSUTIL_AVAILABLE:
                    try:
                        for proc in psutil.process_iter(['pid', 'name']):
                            try:
                                proc_name = proc.info.get('name') or ''
                                proc_lower = proc_name.lower()
                                # Look for mscore or musescore
                                if proc_lower == 'mscore' or 'musescore' in proc_lower:
                                    found_processes.append(proc_name)
                                    self.log(f"  Found process: {proc_name} (PID: {proc.info.get('pid')})")
                            except (psutil.NoSuchProcess, psutil.AccessDenied):
                                continue
                    except Exception as e:
                        self.log(f"  Could not list processes: {e}")
                
                # Also try AppleScript to list processes
                list_script = '''
                tell application "System Events"
                    set processList to name of every process
                    set resultList to {}
                    repeat with procName in processList
                        if procName contains "Muse" or procName contains "muse" or procName contains "Score" or procName contains "score" then
                            set end of resultList to procName
                        end if
                    end repeat
                    return resultList
                end tell
                '''
                list_success, list_output, list_error = run_applescript(list_script)
                if list_success and list_output:
                    self.log(f"  AppleScript found processes: {list_output}")
                
                if not found_processes and not list_output:
                    self.log("  No MuseScore-related processes found.")
                    self.log("\nTroubleshooting tips:")
                    self.log("  1. Make sure MuseScore 4 is running with a score open")
                    self.log("  2. Check System Preferences > Security & Privacy > Privacy > Accessibility")
                    self.log("     - Ensure Terminal/Python has accessibility permissions")
                    self.log("  3. Try restarting MuseScore 4")
                
                self.root.after(0, lambda: messagebox.showerror(
                    "MuseScore Not Found",
                    "Could not find a running MuseScore window.\n\n"
                    "Please ensure MuseScore 4 is open with a score loaded and try again.\n\n"
                    "Check the output log for debugging information.\n\n"
                    "Note: You may need to grant accessibility permissions to Terminal/Python\n"
                    "in System Preferences > Security & Privacy > Privacy > Accessibility."
                ))
                return
            
            self.log(f"✓ Found MuseScore: {output}")
            
            # Activate the window
            self.log("Step 2: Activating MuseScore window...")
            success, output, error = activate_musescore_window_macos()
            if success:
                self.log("✓ Activated MuseScore window")
            else:
                self.log(f"⚠ Warning: Could not activate window: {error}")
            
            # Wait a moment for window to activate
            time.sleep(0.5)
            
            # Send keyboard shortcut
            self.log("Step 3: Sending keyboard shortcut Cmd+Shift+S...")
            success, output, error = send_shortcut_macos()
            if success:
                self.log("✓ Keyboard shortcut sent successfully!")
                time.sleep(0.5)  # Wait a moment for dialog to appear
                self.log("Please complete the save dialog in MuseScore...")
                self.root.after(0, lambda: messagebox.showinfo(
                    "Save Selection Triggered",
                    "Save Selection dialog should now be open in MuseScore.\n\n"
                    "Please:\n"
                    "1. Choose the save location (preferably the watch folder)\n"
                    "2. Enter a filename\n"
                    "3. Click Save\n\n"
                    "If watching is enabled, the file will be processed automatically.\n\n"
                    "If the dialog didn't open, check the output log for details."
                ))
            else:
                self.log(f"✗ Failed to send keyboard shortcut: {error}")
                self.root.after(0, lambda: messagebox.showerror(
                    "Error",
                    "Could not send keyboard shortcut to MuseScore.\n\n"
                    "Please check the output log for details.\n\n"
                    "You can still use the manual method:\n"
                    "File → Save Selection (or Cmd+Shift+S)"
                ))
        
        except Exception as e:
            error_msg = f"Error triggering Save Selection: {str(e)}"
            self.log(f"✗ {error_msg}")
            import traceback
            self.log(f"Traceback:\n{traceback.format_exc()}")
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))

    def setup_hotkey_request_monitor(self):
        """Monitor a temp file for hotkey requests from the background listener."""
        self.hotkey_request_path = HOTKEY_REQUEST_FILE
        try:
            if not self.hotkey_request_path.exists():
                self.hotkey_request_path.write_text("0")
            self._last_hotkey_request = self.hotkey_request_path.stat().st_mtime
        except Exception:
            self._last_hotkey_request = 0

        monitor_thread = threading.Thread(target=self._monitor_hotkey_request, daemon=True)
        monitor_thread.start()

    def _monitor_hotkey_request(self):
        while not self._hotkey_monitor_stop.is_set():
            try:
                if self.hotkey_request_path.exists():
                    mtime = self.hotkey_request_path.stat().st_mtime
                    if mtime > self._last_hotkey_request:
                        self._last_hotkey_request = mtime
                        self.log("Global hotkey request detected (background listener).")
                        self.root.after(0, self.trigger_save_selection)
                time.sleep(0.6)
            except Exception:
                time.sleep(1)
    
    def register_global_hotkey(self):
        """Register a global keyboard shortcut to trigger Save Selection"""
        if self.disable_global_hotkey:
            self.log("Global hotkey registration skipped (external listener handles Ctrl+Cmd+S).")
            return
        if not KEYBOARD_AVAILABLE:
            self.log("Global hotkey not available: Install 'keyboard' library (pip install keyboard)")
            return
        
        # Default hotkey: Ctrl+Cmd+S (won't conflict with MuseScore's Cmd+Shift+S)
        hotkey = "ctrl+cmd+s"
        
        try:
            # Register the hotkey
            keyboard.add_hotkey(hotkey, self.trigger_save_selection, suppress=False)
            self.log(f"✓ Global hotkey registered: {hotkey.upper()}")
            self.log("  You can now press Ctrl+Cmd+S from anywhere to trigger Save Selection!")
        except Exception as e:
            self.log(f"✗ Failed to register global hotkey: {str(e)}")
            messagebox.showwarning(
                "Hotkey Registration Failed",
                f"Could not register global hotkey:\n{str(e)}\n\n"
                "You can still use the button in the app."
            )
    
    def on_closing(self):
        """Cleanup when window is closed"""
        # Unregister hotkey if it was registered
        if KEYBOARD_AVAILABLE:
            try:
                keyboard.unhook_all_hotkeys()
            except:
                pass
        self._hotkey_monitor_stop.set()
        watching_state = self.watching
        self.save_preferences(watching_override=watching_state)
        self.watching = False
        self.root.destroy()


def main():
    parser = argparse.ArgumentParser(description="MuseScore Pitch Extractor GUI")
    parser.add_argument(
        "--trigger-save-selection",
        action="store_true",
        help="Trigger the Save Selection automation as soon as the UI loads."
    )
    parser.add_argument(
        "--disable-global-hotkey",
        action="store_true",
        help="Skip registering the in-app Ctrl+Cmd+S hotkey (used by the background listener)."
    )

    args = parser.parse_args()

    root = tk.Tk()
    app = MuseScoreExtractorApp(
        root,
        trigger_on_start=args.trigger_save_selection,
        disable_global_hotkey=args.disable_global_hotkey
    )
    root.mainloop()


if __name__ == "__main__":
    main()

