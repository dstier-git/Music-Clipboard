import xml.etree.ElementTree as ET
import os
import zipfile
from pathlib import Path

# Output directory for extracted files - use relative path from script location
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "txts")

def get_pitch_name(pitch_value):
    """Convert MIDI pitch number to note name (e.g., 60 -> C4)"""
    pitch_names = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]
    note = pitch_value % 12
    octave = (pitch_value // 12) - 1
    return f"{pitch_names[note]}{octave}"

def extract_pitches_from_mscx(mscx_file_path, output_file_path=None, debug=True):
    """Extract pitch names from a MuseScore .mscx or .mscz file"""
    try:
        # Check if file is .mscz (compressed)
        if mscx_file_path.endswith('.mscz'):
            print("Detected .mscz file, extracting...")
            with zipfile.ZipFile(mscx_file_path, 'r') as zip_ref:
                # The main score is usually in a file without extension or named with .mscx
                file_list = zip_ref.namelist()
                
                # Look for the main score file (usually the largest XML file)
                score_file = None
                for name in file_list:
                    if name.endswith('.mscx') or (not '.' in name and not name.endswith('/')):
                        score_file = name
                        break
                
                if score_file is None:
                    # Try the first file
                    score_file = file_list[0]
                
                print(f"Reading {score_file} from archive...")
                with zip_ref.open(score_file) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
        else:
            # Regular .mscx file
            tree = ET.parse(mscx_file_path)
            root = tree.getroot()
        
        if debug:
            print(f"\nRoot tag: {root.tag}")
            print(f"Root attributes: {root.attrib}")
            
            # Print first few levels of structure
            print("\nXML Structure (first 20 elements):")
            count = 0
            for elem in root.iter():
                if count < 20:
                    print(f"  {elem.tag}: {elem.text[:50] if elem.text and elem.text.strip() else ''}")
                    count += 1
        
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
                    if debug and len(pitches) <= 5:
                        print(f"Found note (Chord/Note/pitch): {pitch_name} (MIDI: {midi_pitch})")
        
        # Approach 2: Direct Note elements with pitch
        if not pitches:
            for note in root.iter('Note'):
                pitch_elem = note.find('pitch')
                if pitch_elem is not None and pitch_elem.text:
                    midi_pitch = int(pitch_elem.text)
                    pitch_name = get_pitch_name(midi_pitch)
                    pitches.append(pitch_name)
                    if debug and len(pitches) <= 5:
                        print(f"Found note (Note/pitch): {pitch_name} (MIDI: {midi_pitch})")
        
        # Approach 3: Try with namespaces
        if not pitches:
            ns = {'m': 'http://www.musescore.org/mscx'}
            for note in root.findall('.//m:Note', ns):
                pitch_elem = note.find('m:pitch', ns)
                if pitch_elem is not None and pitch_elem.text:
                    midi_pitch = int(pitch_elem.text)
                    pitch_name = get_pitch_name(midi_pitch)
                    pitches.append(pitch_name)
        
        if debug:
            print(f"\n{'='*50}")
            if pitches:
                print(f"Successfully extracted {len(pitches)} pitches!")
            else:
                print("No pitches found. Checking for Note elements...")
                note_count = len(list(root.iter('Note')))
                chord_count = len(list(root.iter('Chord')))
                print(f"Found {note_count} Note elements")
                print(f"Found {chord_count} Chord elements")
        
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
            
            print(f"Extracted {len(pitches)} pitches to: {output_file_path}")
        
        return pitches
        
    except Exception as e:
        print(f"Error processing file: {e}")
        import traceback
        traceback.print_exc()
        return None

def main():
    """Interactive main function"""
    print("=" * 60)
    print("MuseScore Pitch Extractor")
    print("=" * 60)
    print()
    
    # Get file path from user
    file_path = input("Enter the path to your MuseScore file (.mscx or .mscz): ").strip()
    
    # Remove quotes if user pasted a path with quotes
    file_path = file_path.strip('"').strip("'")
    
    # Check if file exists
    if not os.path.exists(file_path):
        print(f"\nError: File not found: {file_path}")
        return
    
    # Check if it's a valid MuseScore file
    if not (file_path.endswith('.mscx') or file_path.endswith('.mscz')):
        print(f"\nWarning: File doesn't have .mscx or .mscz extension. Continuing anyway...")
    
    print(f"\nProcessing: {file_path}")
    print("-" * 60)
    
    # Extract pitches
    pitches = extract_pitches_from_mscx(file_path, debug=True)
    
    # Display results
    if pitches:
        print(f"\n{'='*60}")
        print(f"Extraction complete!")
        print(f"Total pitches extracted: {len(pitches)}")
        print(f"\nFirst 20 pitches:")
        for i, pitch in enumerate(pitches[:20], 1):
            print(f"  {i}. {pitch}")
        if len(pitches) > 20:
            print(f"  ... and {len(pitches) - 20} more")
        
        # Show output file location
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_file = os.path.join(OUTPUT_DIR, base_name + "_pitches.txt")
        print(f"\nPitches saved to: {output_file}")
    else:
        print("\nNo pitches extracted. Please check the debug output above.")

if __name__ == "__main__":
    main()

