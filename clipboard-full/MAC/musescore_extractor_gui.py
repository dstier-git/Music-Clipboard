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

# Try to import automation libraries
try:
    import pyautogui
    # Disable pyautogui failsafe (the mouse moving to corner)
    pyautogui.FAILSAFE = False
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False

# Check if we're on macOS
IS_MACOS = platform.system() == 'Darwin'

# Output directory for extracted files - use relative path from script location
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "txts")

# Import the extraction function
EXTRACTION_FUNCTION = None
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


def run_applescript(script):
    """Run an AppleScript command and return the result"""
    try:
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True,
            text=True,
            timeout=10
        )
        return result.returncode == 0, result.stdout.strip(), result.stderr.strip()
    except subprocess.TimeoutExpired:
        return False, "", "Timeout"
    except Exception as e:
        return False, "", str(e)


def find_musescore_window_macos():
    """Find MuseScore window on macOS using AppleScript"""
    script = '''
    tell application "System Events"
        set museScoreProcess to first process whose name contains "MuseScore"
        return name of museScoreProcess
    end tell
    '''
    success, output, error = run_applescript(script)
    return success, output, error


def activate_musescore_window_macos():
    """Activate MuseScore window on macOS using AppleScript"""
    # Try multiple approaches to activate MuseScore
    scripts = [
        # Try MuseScore 4 first
        '''
        tell application "MuseScore 4"
            activate
        end tell
        ''',
        # Try MuseScore 3
        '''
        tell application "MuseScore 3"
            activate
        end tell
        ''',
        # Try generic MuseScore
        '''
        tell application "MuseScore"
            activate
        end tell
        ''',
        # Fallback: Use System Events
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
            return success, output, error
    
    return False, "", "Could not activate MuseScore"


def send_shortcut_macos():
    """Send Cmd+Shift+S shortcut on macOS"""
    # Try multiple process names
    process_names = ["MuseScore 4", "MuseScore 3", "MuseScore"]
    
    for process_name in process_names:
        script = f'''
        tell application "System Events"
            try
                tell process "{process_name}"
                    keystroke "s" using {{command down, shift down}}
                    return true
                end tell
            on error
                return false
            end try
        end tell
        '''
        success, output, error = run_applescript(script)
        if success:
            return success, output, error
    
    # Fallback: try to find process dynamically
    script = '''
    tell application "System Events"
        set museScoreProcess to first process whose name contains "MuseScore"
        tell museScoreProcess
            keystroke "s" using {command down, shift down}
        end tell
    end tell
    '''
    success, output, error = run_applescript(script)
    return success, output, error


class MuseScoreExtractorApp:
    def __init__(self, root):
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
        
        # Variables
        self.watch_folder = tk.StringVar()
        self.watching = False
        self.watch_thread = None
        self.processed_files = set()
        self.use_measure_range = tk.BooleanVar(value=False)
        self.measure_start = tk.StringVar(value="1")
        self.measure_end = tk.StringVar(value="1")
        
        # Create UI
        self.create_widgets()
        
        # Set default watch folder (MuseScore default export location or user's Documents)
        default_folder = os.path.join(os.path.expanduser("~"), "Documents", "MuseScore4", "Scores")
        if os.path.exists(default_folder):
            self.watch_folder.set(default_folder)
        else:
            self.watch_folder.set(os.path.join(os.path.expanduser("~"), "Documents"))
    
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
        
        # Measure range selection
        measure_frame = ttk.Frame(file_frame)
        measure_frame.grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))
        
        ttk.Checkbutton(measure_frame, text="Extract only specific measures:", 
                       variable=self.use_measure_range,
                       command=self.toggle_measure_range).grid(row=0, column=0, sticky=tk.W, padx=5)
        
        ttk.Label(measure_frame, text="From measure:").grid(row=0, column=1, padx=(20, 5))
        measure_start_entry = ttk.Entry(measure_frame, textvariable=self.measure_start, width=5)
        measure_start_entry.grid(row=0, column=2, padx=5)
        
        ttk.Label(measure_frame, text="To measure:").grid(row=0, column=3, padx=(10, 5))
        measure_end_entry = ttk.Entry(measure_frame, textvariable=self.measure_end, width=5)
        measure_end_entry.grid(row=0, column=4, padx=5)
        
        # Store widgets for enabling/disabling
        self.measure_range_widgets = [measure_start_entry, measure_end_entry]
        for widget in self.measure_range_widgets:
            widget.config(state='disabled')
        
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
        
        if not IS_MACOS:
            self.save_selection_button.config(state='disabled')
            ttk.Label(
                automation_frame, 
                text="(macOS only)", 
                foreground="gray", 
                font=("Arial", 8)
            ).grid(row=0, column=1, padx=5)
        elif not PYAUTOGUI_AVAILABLE:
            self.save_selection_button.config(state='disabled')
            ttk.Label(
                automation_frame, 
                text="(Install: pip install pyautogui)", 
                foreground="gray", 
                font=("Arial", 8)
            ).grid(row=0, column=1, padx=5)
        
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
        
        # Clear button
        ttk.Button(output_frame, text="Clear Output", command=self.clear_output).grid(row=1, column=0, pady=5)
    
    def toggle_measure_range(self):
        """Enable/disable measure range entry fields"""
        state = 'normal' if self.use_measure_range.get() else 'disabled'
        for widget in self.measure_range_widgets:
            widget.config(state=state)
    
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
    
    def log(self, message):
        """Add message to output text area"""
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_output(self):
        self.output_text.delete(1.0, tk.END)
    
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
        
        # Get measure range if specified
        measure_range = None
        if self.use_measure_range.get():
            try:
                start = int(self.measure_start.get())
                end = int(self.measure_end.get())
                if start < 1 or end < 1:
                    messagebox.showerror("Error", "Measure numbers must be 1 or greater.")
                    return
                if start > end:
                    messagebox.showerror("Error", "Start measure must be less than or equal to end measure.")
                    return
                measure_range = (start, end)
            except ValueError:
                messagebox.showerror("Error", "Please enter valid measure numbers.")
                return
        
        # Run extraction in a separate thread to avoid freezing UI
        thread = threading.Thread(target=self._extract_thread, args=(file_path, measure_range), daemon=True)
        thread.start()
    
    def _extract_thread(self, file_path, measure_range=None):
        """Extraction logic running in background thread"""
        self.log(f"\n{'='*60}")
        self.log(f"Processing: {os.path.basename(file_path)}")
        if measure_range:
            self.log(f"Measure range: {measure_range[0]}-{measure_range[1]}")
        self.log(f"{'='*60}\n")
        
        try:
            # Create output directory if it doesn't exist
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            
            # Determine output file path
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            if EXTRACTION_SCRIPT == "extract_pitches_with_position":
                suffix = "_pitches_with_position"
                if measure_range:
                    suffix += f"_m{measure_range[0]}-{measure_range[1]}"
                output_file = os.path.join(OUTPUT_DIR, base_name + suffix + ".txt")
            else:
                suffix = "_pitches"
                if measure_range:
                    suffix += f"_m{measure_range[0]}-{measure_range[1]}"
                output_file = os.path.join(OUTPUT_DIR, base_name + suffix + ".txt")
            
            # Extract pitches with positions
            # Check if the function accepts measure_range parameter
            import inspect
            sig = inspect.signature(EXTRACTION_FUNCTION)
            if 'measure_range' in sig.parameters:
                result = EXTRACTION_FUNCTION(file_path, output_file, debug=False, measure_range=measure_range)
            else:
                result = EXTRACTION_FUNCTION(file_path, output_file, debug=False)
            
            # Handle return value - could be (notes, path) tuple or just notes
            if isinstance(result, tuple) and len(result) == 2:
                notes, actual_output_path = result
                output_file = actual_output_path  # Use the actual path from extraction function
            else:
                notes = result
            
            if notes:
                self.log(f"✓ Successfully extracted {len(notes)} notes!")
                self.log(f"✓ Output saved to: {output_file}\n")
                
                # Show first 10 notes
                self.log("First 10 notes:")
                for i, (pitch, position, tick) in enumerate(notes[:10], 1):
                    if tick is not None:
                        self.log(f"  {i}. {pitch} | {position} | (tick: {tick})")
                    else:
                        self.log(f"  {i}. {pitch} | {position}")
                
                if len(notes) > 10:
                    self.log(f"  ... and {len(notes) - 10} more\n")
                
                # Show success message
                self.root.after(0, lambda: messagebox.showinfo(
                    "Success", 
                    f"Extracted {len(notes)} notes!\n\nSaved to:\n{output_file}"
                ))
            else:
                self.log("✗ No notes extracted. Please check the file format.\n")
                self.root.after(0, lambda: messagebox.showerror("Error", "No notes were extracted from the file."))
        
        except Exception as e:
            error_msg = f"Error processing file: {str(e)}"
            self.log(f"✗ {error_msg}\n")
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
            
            # Start watching in background thread
            self.watch_thread = threading.Thread(target=self._watch_folder, args=(folder,), daemon=True)
            self.watch_thread.start()
        else:
            self.watching = False
            self.watch_button.config(text="Start Watching")
            self.watch_status_label.config(text="Status: Not watching", foreground="gray")
            self.log("Stopped watching folder.\n")
    
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
            
            # Find MuseScore using AppleScript
            success, output, error = find_musescore_window_macos()
            if not success:
                self.log("✗ Could not find MuseScore window")
                self.log(f"Error: {error}")
                self.root.after(0, lambda: messagebox.showerror(
                    "MuseScore Not Found",
                    "Could not find a running MuseScore window.\n\n"
                    "Please ensure MuseScore 4 is open with a score loaded and try again.\n\n"
                    "Check the output log for details."
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


def main():
    root = tk.Tk()
    app = MuseScoreExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

