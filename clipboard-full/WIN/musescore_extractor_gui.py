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
import tempfile

# Try to import automation libraries
try:
    import pyautogui
    # Disable pyautogui failsafe (the mouse moving to corner)
    pyautogui.FAILSAFE = False
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False

try:
    from pywinauto import Application
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False

# Try Windows API for additional fallback
try:
    import win32gui
    import win32con
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# Try psutil for process searching
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

# Try keyboard library for global hotkeys
try:
    import keyboard
    KEYBOARD_AVAILABLE = True
except ImportError:
    KEYBOARD_AVAILABLE = False

# Path to persist preferences
CONFIG_FILE = Path(os.path.expanduser("~")) / ".musescore_pitch_extractor_prefs"

# Location for hotkey trigger requests from the background listener
HOTKEY_REQUEST_FILE = Path(tempfile.gettempdir()) / "musescore_hotkey_request.txt"

# Output directories for extracted files
OUTPUT_DIR = r"C:\Users\janet\Desktop\Music-Clipboard\clipboard-full\txts"
MIDI_OUTPUT_DIR = r"C:\Users\janet\Desktop\Music-Clipboard\clipboard-full\midis"

# Import the extraction functions
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
        format_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))
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
                text="Global Hotkey: Ctrl+Alt+S (background listener keeps it active even when the GUI is closed)",
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
        
        if not PYAUTOGUI_AVAILABLE or not PYWINAUTO_AVAILABLE:
            self.save_selection_button.config(state='disabled')
            missing_libs = []
            if not PYWINAUTO_AVAILABLE:
                missing_libs.append("pywinauto")
            if not PYAUTOGUI_AVAILABLE:
                missing_libs.append("pyautogui")
            ttk.Label(
                automation_frame, 
                text=f"(Install: pip install {' '.join(missing_libs)})", 
                foreground="gray", 
                font=("Arial", 8)
            ).grid(row=1, column=0, columnspan=2, padx=5, sticky=tk.W)
        
        self.watch_button = ttk.Button(watch_frame, text="Start Watching", command=self.toggle_watch)
        self.watch_button.grid(row=2, column=0, columnspan=3, pady=10)
        
        self.watch_status_label = ttk.Label(watch_frame, text="Status: Not watching", foreground="gray")
        self.watch_status_label.grid(row=3, column=0, columnspan=3)
        
        # Instructions
        instructions = """
Instructions:
1. Manual Mode:
   - Select a .mscx or .mscz file
   - Click 'Extract' to process
2. Auto Mode (Save Selection):
   - Set the watch folder (where MuseScore saves selections)
   - Click 'Start Watching' to begin monitoring
   - In MuseScore: Select the measures you want to extract
   - Click 'Trigger Save Selection in MuseScore' button (or manually: File > Save Selection)
   - In the save dialog that opens:
     • Choose the save location (preferably the watch folder)
     • Enter a filename
     • Click Save
   - The app will automatically process the new file if watching is enabled
   - Note: Saved selections already contain only selected measures
   - If the dialog didn't open, check the output log for details
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
        """Open the file location of the last extracted file in Windows Explorer"""
        if self.last_extracted_file and os.path.exists(self.last_extracted_file):
            try:
                # Open Windows Explorer and select the file
                subprocess.run(['explorer', '/select,', os.path.normpath(self.last_extracted_file)])
            except Exception as e:
                # Fallback: just open the folder
                try:
                    folder = os.path.dirname(self.last_extracted_file)
                    os.startfile(folder)
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
        if not PYWINAUTO_AVAILABLE or not PYAUTOGUI_AVAILABLE:
            messagebox.showerror(
                "Missing Dependencies",
                "This feature requires pywinauto and pyautogui.\n\n"
                "Install them with:\n"
                "pip install pywinauto pyautogui"
            )
            return
        
        # Run in a separate thread to avoid freezing UI
        thread = threading.Thread(target=self._trigger_save_selection_thread, daemon=True)
        thread.start()
    
    def _trigger_save_selection_thread(self):
        """Thread function to trigger Save Selection"""
        app = None
        main_window = None
        
        try:
            self.log("Attempting to trigger Save Selection in MuseScore...")
            self.log("Step 1: Finding MuseScore window...")
            
            # Find MuseScore window using pywinauto - try multiple methods
            # IMPORTANT: We need to find the actual MuseScore application, not our own app window
            found = False
            methods_tried = []
            
            # Method 1: Try by process name first (most reliable - looks for MuseScore4.exe)
            try:
                app = Application(backend="uia").connect(path="MuseScore4.exe")
                methods_tried.append("UIA backend by process")
                found = True
                self.log("✓ Found MuseScore process (UIA backend by process name)")
            except Exception as e:
                methods_tried.append(f"UIA by process failed: {str(e)[:50]}")
            
            # Method 2: Try by process name with win32 backend
            if not found:
                try:
                    app = Application(backend="win32").connect(path="MuseScore4.exe")
                    methods_tried.append("Win32 backend by process")
                    found = True
                    self.log("✓ Found MuseScore process (Win32 backend by process name)")
                except Exception as e:
                    methods_tried.append(f"Win32 by process failed: {str(e)[:50]}")
            
            # Method 3: Try by process name without backend
            if not found:
                try:
                    app = Application().connect(path="MuseScore4.exe")
                    methods_tried.append("Default backend by process")
                    found = True
                    self.log("✓ Found MuseScore process (default backend by process name)")
                except Exception as e:
                    methods_tried.append(f"Default by process failed: {str(e)[:50]}")
            
            # Method 4: Try by title but exclude our own app window
            if not found:
                try:
                    # Try to get all windows with MuseScore in title and filter
                    # Use a more specific pattern that excludes "Pitch Extractor"
                    try:
                        # Try connecting to MuseScore 4 specifically
                        app = Application(backend="uia").connect(title_re=".*MuseScore 4.*")
                        found = True
                        self.log("✓ Found MuseScore window (UIA backend by 'MuseScore 4' title)")
                    except:
                        # If that fails, try to enumerate and filter
                        try:
                            # Get all processes and find MuseScore4.exe
                            if PSUTIL_AVAILABLE:
                                for proc in psutil.process_iter(['pid', 'name']):
                                    try:
                                        proc_name = proc.info['name'] or ''
                                        if 'MuseScore4.exe' in proc_name or ('MuseScore' in proc_name and 'Pitch' not in proc_name):
                                            app = Application(backend="uia").connect(process=proc.info['pid'])
                                            found = True
                                            self.log(f"✓ Found MuseScore by process enumeration: {proc_name}")
                                            break
                                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                                        continue
                        except Exception as e:
                            methods_tried.append(f"Process enumeration failed: {str(e)[:50]}")
                except Exception as e:
                    methods_tried.append(f"UIA by filtered title failed: {str(e)[:50]}")
            
            # Method 5: Try using Windows API to find window handle (exclude our app)
            if not found and WIN32_AVAILABLE:
                try:
                    def enum_handler(hwnd, ctx):
                        if win32gui.IsWindowVisible(hwnd):
                            window_text = win32gui.GetWindowText(hwnd)
                            # Exclude our own app and look for actual MuseScore
                            if 'MuseScore' in window_text and 'Pitch Extractor' not in window_text:
                                # Also check the process name if psutil is available
                                try:
                                    _, pid = win32gui.GetWindowThreadProcessId(hwnd)
                                    if PSUTIL_AVAILABLE:
                                        try:
                                            proc = psutil.Process(pid)
                                            proc_name = proc.name()
                                            if 'MuseScore' in proc_name and 'Pitch Extractor' not in proc_name:
                                                ctx.append((hwnd, window_text, proc_name))
                                        except:
                                            # If we can't check process, still add it if title looks right
                                            if 'MuseScore 4' in window_text or window_text.startswith('MuseScore'):
                                                ctx.append((hwnd, window_text, 'unknown'))
                                    else:
                                        # No psutil, but title looks right
                                        if 'MuseScore 4' in window_text or window_text.startswith('MuseScore'):
                                            ctx.append((hwnd, window_text, 'unknown'))
                                except:
                                    pass
                    
                    windows = []
                    win32gui.EnumWindows(enum_handler, windows)
                    
                    if windows:
                        hwnd, window_text, proc_name = windows[0]
                        # Try to connect using handle
                        try:
                            app = Application().connect(handle=hwnd)
                            found = True
                            self.log(f"✓ Found MuseScore window using Windows API: '{window_text}' (process: {proc_name})")
                        except:
                            pass
                except Exception as e:
                    methods_tried.append(f"Windows API failed: {str(e)[:50]}")
            
            # Method 6: Last resort - try to find by checking all processes (if psutil available)
            if not found and PSUTIL_AVAILABLE:
                try:
                    for proc in psutil.process_iter(['pid', 'name', 'exe']):
                        try:
                            proc_name = proc.info['name'] or ''
                            proc_exe = proc.info['exe'] or ''
                            # Look for MuseScore4.exe specifically
                            if 'MuseScore4.exe' in proc_exe or (proc_name and 'MuseScore' in proc_name and '4' in proc_name and 'Pitch' not in proc_name):
                                app = Application(backend="uia").connect(process=proc.info['pid'])
                                found = True
                                self.log(f"✓ Found MuseScore by process search: {proc_name}")
                                break
                        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                            continue
                except Exception as e:
                    methods_tried.append(f"Process search failed: {str(e)[:50]}")
            
            if not found:
                self.log("✗ Could not find MuseScore window")
                self.log("Methods tried:")
                for method in methods_tried:
                    self.log(f"  - {method}")
                self.root.after(0, lambda: messagebox.showerror(
                    "MuseScore Not Found",
                    "Could not find a running MuseScore window.\n\n"
                    "Please ensure MuseScore 4 is open with a score loaded and try again.\n\n"
                    "Check the output log for details."
                ))
                return
            
            # Get the main window
            self.log("Step 2: Accessing main window...")
            try:
                main_window = app.top_window()
                window_title = main_window.window_text()
                self.log(f"✓ Found window: '{window_title}'")
            except Exception as e:
                self.log(f"✗ Error accessing window: {str(e)}")
                self.root.after(0, lambda: messagebox.showerror(
                    "Error",
                    f"Could not access MuseScore window: {str(e)}"
                ))
                return
            
            # Activate the window - try multiple methods
            self.log("Step 3: Activating MuseScore window...")
            activated = False
            
            # Method 1: set_focus
            try:
                main_window.set_focus()
                time.sleep(0.5)  # Give it more time to activate
                activated = True
                self.log("✓ Activated using set_focus()")
            except Exception as e:
                self.log(f"  set_focus() failed: {str(e)[:50]}")
            
            # Method 2: set_foreground
            if not activated:
                try:
                    main_window.set_foreground()
                    time.sleep(0.5)
                    activated = True
                    self.log("✓ Activated using set_foreground()")
                except Exception as e:
                    self.log(f"  set_foreground() failed: {str(e)[:50]}")
            
            # Method 3: restore and activate
            if not activated:
                try:
                    main_window.restore()
                    main_window.set_focus()
                    time.sleep(0.5)
                    activated = True
                    self.log("✓ Activated using restore() + set_focus()")
                except Exception as e:
                    self.log(f"  restore() + set_focus() failed: {str(e)[:50]}")
            
            # Method 4: Use Windows API to bring to front
            if not activated and WIN32_AVAILABLE:
                try:
                    hwnd = main_window.handle
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.BringWindowToTop(hwnd)
                    time.sleep(0.5)
                    activated = True
                    self.log("✓ Activated using Windows API")
                except Exception as e:
                    self.log(f"  Windows API activation failed: {str(e)[:50]}")
            
            if not activated:
                self.log("⚠ Warning: Could not activate window, but continuing anyway...")
            
            # Additional delay to ensure window is ready
            time.sleep(0.5)  # Increased delay
            
            # Send keyboard shortcut - try multiple methods
            self.log("Step 4: Sending keyboard shortcut Ctrl+Shift+S...")
            shortcut_sent = False
            
            # Method 1: Use pywinauto's send_keystrokes (more reliable for specific window)
            try:
                # Ensure window is still focused
                main_window.set_focus()
                time.sleep(0.2)
                main_window.type_keys('^+s', with_spaces=False, pause=0.1)
                shortcut_sent = True
                self.log("✓ Sent shortcut using pywinauto type_keys()")
            except Exception as e:
                self.log(f"  pywinauto type_keys() failed: {str(e)[:50]}")
            
            # Method 1b: Try alternative pywinauto syntax
            if not shortcut_sent:
                try:
                    main_window.set_focus()
                    time.sleep(0.2)
                    # Try with different syntax - send_keys is a method on the window
                    main_window.send_keystrokes('^+s')
                    shortcut_sent = True
                    self.log("✓ Sent shortcut using pywinauto send_keystrokes()")
                except Exception as e:
                    self.log(f"  pywinauto send_keystrokes() failed: {str(e)[:50]}")
            
            # Method 2: Use pyautogui (global)
            if not shortcut_sent:
                try:
                    # Make sure we're still focused on the window
                    main_window.set_focus()
                    time.sleep(0.2)
                    pyautogui.hotkey('ctrl', 'shift', 's')
                    shortcut_sent = True
                    self.log("✓ Sent shortcut using pyautogui hotkey()")
                except Exception as e:
                    self.log(f"  pyautogui hotkey() failed: {str(e)[:50]}")
            
            # Method 3: Use pyautogui with individual key presses
            if not shortcut_sent:
                try:
                    main_window.set_focus()
                    time.sleep(0.2)
                    pyautogui.keyDown('ctrl')
                    pyautogui.keyDown('shift')
                    pyautogui.press('s')
                    pyautogui.keyUp('shift')
                    pyautogui.keyUp('ctrl')
                    shortcut_sent = True
                    self.log("✓ Sent shortcut using pyautogui keyDown/Up()")
                except Exception as e:
                    self.log(f"  pyautogui keyDown/Up() failed: {str(e)[:50]}")
            
            if shortcut_sent:
                self.log("✓ Keyboard shortcut sent successfully!")
                time.sleep(0.5)  # Wait a moment for dialog to appear
                self.log("Please complete the save dialog in MuseScore...")
            else:
                self.log("✗ Failed to send keyboard shortcut with all methods")
                self.root.after(0, lambda: messagebox.showerror(
                    "Error",
                    "Could not send keyboard shortcut to MuseScore.\n\n"
                    "Please check the output log for details.\n\n"
                    "You can still use the manual method:\n"
                    "File → Save Selection (or Ctrl+Shift+S)"
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
            self.log("Global hotkey registration skipped (external listener handles Ctrl+Alt+S).")
            return
        if not KEYBOARD_AVAILABLE:
            self.log("Global hotkey not available: Install 'keyboard' library (pip install keyboard)")
            return
        
        # Default hotkey: Ctrl+Alt+S (won't conflict with MuseScore's Ctrl+Shift+S)
        hotkey = "ctrl+alt+s"
        
        try:
            # Register the hotkey
            keyboard.add_hotkey(hotkey, self.trigger_save_selection, suppress=False)
            self.log(f"✓ Global hotkey registered: {hotkey.upper()}")
            self.log("  You can now press Ctrl+Alt+S from anywhere to trigger Save Selection!")
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
        help="Skip registering the in-app Ctrl+Alt+S hotkey (used by the background listener)."
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
