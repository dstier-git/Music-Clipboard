import xml.etree.ElementTree as ET
import os
import zipfile
import subprocess
import shutil

MIDI_OUTPUT_DIR = r"C:\Users\janet\Desktop\Music-Clipboard\clipboard-full\midis"


def extract_midi_from_mscx(mscx_file_path, output_file_path=None, measure_range=None):
    """Extract MIDI from .mscx or .mscz file using MuseScore CLI or mido library
    
    Args:
        mscx_file_path: Path to the MuseScore file
        output_file_path: Path for output MIDI file (optional, will be auto-generated if None)
        measure_range: Tuple (start_measure, end_measure) to extract only specific measures (1-indexed, inclusive)
                       If None, extracts all measures. Note: Only supported with library-based extraction, not MuseScore CLI.
    
    Returns:
        Path to the created MIDI file, or None if extraction failed
    """
    # Create output directory if it doesn't exist
    MIDI_OUTPUT_DIR_NORMALIZED = os.path.normpath(MIDI_OUTPUT_DIR)
    os.makedirs(MIDI_OUTPUT_DIR_NORMALIZED, exist_ok=True)
    
    # Get base filename from input file
    base_name = os.path.splitext(os.path.basename(mscx_file_path))[0]
    if output_file_path is None:
        filename = base_name + ".mid"
        output_file_path = os.path.normpath(os.path.join(MIDI_OUTPUT_DIR_NORMALIZED, filename))
    
    # Method 1: Try using MuseScore command-line tool (most reliable)
    musescore_paths = [
        r"C:\Program Files\MuseScore 4\bin\MuseScore4.exe",
        r"C:\Program Files (x86)\MuseScore 4\bin\MuseScore4.exe",
        r"C:\Program Files\MuseScore 3\bin\MuseScore3.exe",
        r"C:\Program Files (x86)\MuseScore 3\bin\MuseScore3.exe",
    ]
    
    musescore_exe = None
    for path in musescore_paths:
        if os.path.exists(path):
            musescore_exe = path
            break
    
    # Also try to find MuseScore in PATH
    if musescore_exe is None:
        musescore_exe = shutil.which("MuseScore4.exe") or shutil.which("MuseScore3.exe") or shutil.which("mscore")
    
    if musescore_exe:
        # MuseScore CLI doesn't support measure ranges, so if measure_range is specified,
        # we should use library-based method instead
        if measure_range is None:
            try:
                # Use MuseScore CLI to export to MIDI
                # MuseScore export command: mscore.exe -o output.mid input.mscx
                result = subprocess.run(
                    [musescore_exe, "-o", output_file_path, mscx_file_path],
                    capture_output=True,
                    text=True,
                    timeout=30
                )
                if result.returncode == 0 and os.path.exists(output_file_path):
                    return output_file_path
                else:
                    # If export failed, try with different format
                    # Some versions use different syntax
                    result = subprocess.run(
                        [musescore_exe, mscx_file_path, "-o", output_file_path],
                        capture_output=True,
                    )
                    if result.returncode == 0 and os.path.exists(output_file_path):
                        return output_file_path
            except (subprocess.TimeoutExpired, FileNotFoundError, Exception):
                # Fall through to library-based method
                pass
    
    # Method 2: Use mido library to create MIDI from extracted notes
    try:
        import mido
        from mido import MidiFile, MidiTrack, Message
        
        # Extract note data from the XML
        if mscx_file_path.endswith('.mscz'):
            with zipfile.ZipFile(mscx_file_path, 'r') as zip_ref:
                file_list = zip_ref.namelist()
                score_file = None
                for name in file_list:
                    if name.endswith('.mscx') or (not '.' in name and not name.endswith('/')):
                        score_file = name
                        break
                if score_file is None:
                    score_file = file_list[0]
                with zip_ref.open(score_file) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
        else:
            tree = ET.parse(mscx_file_path)
            root = tree.getroot()
        
        # Get division (ticks per quarter note)
        division = 480  # Default
        for elem in root.iter('Division'):
            if elem.text:
                division = int(elem.text)
                break
        
        # Create MIDI file
        mid = MidiFile()
        track = MidiTrack()
        mid.tracks.append(track)
        
        # Set tempo (default 120 BPM)
        tempo = mido.bpm2tempo(120)
        track.append(Message('set_tempo', tempo=tempo, time=0))
        
        # Extract notes with timing information
        notes = []  # List of (tick, pitch, duration)
        
        # Get all measures and filter if measure_range is specified
        all_measures = list(root.iter('Measure'))
        measures_to_process = []
        start_tick_offset = 0  # Offset to subtract when filtering measures
        
        if measure_range:
            start_measure, end_measure = measure_range
            # Calculate tick offset for measures before the start measure
            for idx, measure in enumerate(all_measures):
                # Try to get measure number from attribute, otherwise use position (1-indexed)
                measure_no_attr = measure.get('no')
                if measure_no_attr is not None:
                    try:
                        measure_no = int(measure_no_attr)
                    except (ValueError, TypeError):
                        measure_no = idx + 1
                else:
                    measure_no = idx + 1
                
                if measure_no < start_measure:
                    # Calculate measure length to accumulate offset
                    # Simple approximation: assume 4/4 time
                    start_tick_offset += division * 4
                elif start_measure <= measure_no <= end_measure:
                    measures_to_process.append((measure_no, measure))
        else:
            measures_to_process = [(idx + 1, m) for idx, m in enumerate(all_measures)]
        
        # Process measures
        current_tick = 0
        for measure_no, measure in measures_to_process:
            measure_tick = 0
            
            for chord in measure.iter('Chord'):
                # Get tick position
                chord_tick = measure_tick
                if 'tick' in chord.attrib:
                    tick_val = int(chord.attrib['tick'])
                    if tick_val < division * 4:  # Likely relative to measure
                        chord_tick = tick_val
                    else:
                        # Absolute tick - subtract offset if filtering
                        chord_tick = tick_val - start_tick_offset
                
                # Get duration
                duration = division  # Default quarter note
                duration_elem = chord.find('duration')
                if duration_elem is not None and duration_elem.text:
                    duration = int(duration_elem.text)
                
                # Get notes in chord
                for note in chord.findall('Note'):
                    pitch_elem = note.find('pitch')
                    if pitch_elem is not None and pitch_elem.text:
                        midi_pitch = int(pitch_elem.text)
                        notes.append((current_tick + chord_tick, midi_pitch, duration))
                
                measure_tick = max(measure_tick, chord_tick + duration)
            
            current_tick += measure_tick if measure_tick > 0 else division * 4
        
        # Sort notes by tick
        notes.sort(key=lambda x: x[0])
        
        # Add notes to MIDI track
        last_tick = 0
        for tick, pitch, duration in notes:
            # Calculate delta time
            delta = tick - last_tick
            
            # Note on
            track.append(Message('note_on', channel=0, note=pitch, velocity=64, time=delta))
            
            # Note off (after duration)
            track.append(Message('note_off', channel=0, note=pitch, velocity=64, time=duration))
            
            last_tick = tick + duration
        
        # Save MIDI file
        mid.save(output_file_path)
        return output_file_path
        
    except ImportError:
        # mido not available, try music21 as fallback
        try:
            from music21 import converter
            
            # Convert MuseScore file to music21 stream
            score = converter.parse(mscx_file_path)
            
            # Export to MIDI
            score.write('midi', output_file_path)
            return output_file_path
        except ImportError:
            raise ImportError(
                "MIDI extraction requires either:\n"
                "1. MuseScore command-line tool installed, or\n"
                "2. Python library 'mido' (pip install mido), or\n"
                "3. Python library 'music21' (pip install music21)"
            )
    except Exception as e:
        raise Exception(f"Failed to extract MIDI: {str(e)}")
