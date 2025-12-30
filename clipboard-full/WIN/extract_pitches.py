import xml.etree.ElementTree as ET
import os
import zipfile
import traceback

from extract_midi import extract_midi_from_mscx

# Output directory for extracted files
OUTPUT_DIR = r"C:\Users\janet\Desktop\Music-Clipboard\clipboard-full\txts"

def get_pitch_name(pitch_value):
    """Convert MIDI pitch note name"""
    pitch_names = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]
    note = pitch_value % 12
    octave = (pitch_value // 12) - 1
    return f"{pitch_names[note]}{octave}"

def extract_pitches_from_mscx(mscx_file_path, output_file_path=None):
    """Extract pitch names from .mscx or .mscz file"""
    # Check if file is .mscz
    if mscx_file_path.endswith('.mscz'):
        with zipfile.ZipFile(mscx_file_path, 'r') as zip_ref:
            file_list = zip_ref.namelist()
            
            # Look for the main score file
            score_file = None
            for name in file_list:
                if name.endswith('.mscx') or (not '.' in name and not name.endswith('/')):
                    score_file = name
                    break
            
            if score_file is None:
                # Try first file
                score_file = file_list[0]
            
            with zip_ref.open(score_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
    else:
        # Regular .mscx file
        tree = ET.parse(mscx_file_path)
        root = tree.getroot()
    
    pitches = []
    
    # Try multiple approaches to find notes
    
    # Approach 1: Look for Chord elements (MuseScore 4 structure)
    for chord in root.iter('Chord'):
        for note in chord.findall('Note'):
            pitch_elem = note.find('pitch')
            if pitch_elem is not None and pitch_elem.text:
                midi_pitch = int(pitch_elem.text)
                pitch_name = get_pitch_name(midi_pitch)
                pitches.append(pitch_name)
    
    # Approach 2: Direct Note elements with pitch
    if not pitches:
        for note in root.iter('Note'):
            pitch_elem = note.find('pitch')
            if pitch_elem is not None and pitch_elem.text:
                midi_pitch = int(pitch_elem.text)
                pitch_name = get_pitch_name(midi_pitch)
                pitches.append(pitch_name)
    
    # Approach 3: Try with namespaces
    if not pitches:
        ns = {'m': 'http://www.musescore.org/mscx'}
        for note in root.findall('.//m:Note', ns):
            pitch_elem = note.find('m:pitch', ns)
            if pitch_elem is not None and pitch_elem.text:
                midi_pitch = int(pitch_elem.text)
                pitch_name = get_pitch_name(midi_pitch)
                pitches.append(pitch_name)
    
    # Write to output file
    if pitches:
        # ALWAYS save to OUTPUT_DIR - ignore any provided path
        OUTPUT_DIR_NORMALIZED = os.path.normpath(OUTPUT_DIR)
        os.makedirs(OUTPUT_DIR_NORMALIZED, exist_ok=True)
        
        # Get base filename from input file (always use this, ignore output_file_path)
        base_name = os.path.splitext(os.path.basename(mscx_file_path))[0]
        filename = base_name + "_pitches.txt"
        
        # Always construct the full path to OUTPUT_DIR
        output_file_path = os.path.normpath(os.path.join(OUTPUT_DIR_NORMALIZED, filename))
        
        with open(output_file_path, 'w', encoding='utf-8') as f:
            for pitch in pitches:
                f.write(pitch + '\n')
    
    return pitches

def main():
    """Interactive main function"""
    print("-" * 60)
    print("MuseScore Pitch Extractor")
    print("-" * 60)
    print()
    
    # Get file path from user
    file_path = input("Enter the path to MuseScore file (.mscx or .mscz): ").strip()
    
    # Remove quotes
    file_path = file_path.strip('"').strip("'")
    
    if not os.path.exists(file_path):
        print(f"\nError: File not found: {file_path}")
        return
    
    # Check if it's a valid file
    if not (file_path.endswith('.mscx') or file_path.endswith('.mscz')):
        print(f"\nWarning: File doesn't have .mscx or .mscz extension...")
    
    # Ask for output format
    print("\nOutput format:")
    print("1. Text (pitch names)")
    print("2. MIDI")
    format_choice = input("Select format (1 or 2, default: 1): ").strip() or "1"
    
    print(f"\nProcessing: {file_path}\n")
    
    try:
        if format_choice == "2":
            # Extract MIDI
            midi_path = extract_midi_from_mscx(file_path)
            if midi_path:
                print(f"\n{'-'*60}")
                print(f"MIDI extraction complete!")
                print(f"MIDI file saved to: {midi_path}")
            else:
                print("\nFailed to extract MIDI.")
        else:
            # Extract pitches (text)
            pitches = extract_pitches_from_mscx(file_path)
            
            # Display results
            if pitches:
                print(f"\n{'-'*60}")
                print(f"Extraction complete!")
                print(f"Total notes extracted: {len(pitches)}")
                print(f"\nFirst 20 pitches:")
                for i, pitch in enumerate(pitches[:20], 1):
                    print(f"  {i}. {pitch}")
                if len(pitches) > 20:
                    print(f"  ... and {len(pitches) - 20} more")
                
                # Show output file location
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                output_file = os.path.join(OUTPUT_DIR, base_name + "_pitches.txt")
                print(f"\nNotes saved to: {output_file}")
            else:
                print("\nNo pitches extracted.")
    except Exception as e:
        print(f"\nError processing file: {e}")
        traceback.print_exc()
        return

if __name__ == "__main__":
    main()

