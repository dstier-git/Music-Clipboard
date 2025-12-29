import QtQuick 2.0
import MuseScore 3.0

MuseScore {
    version: "1.0"
    description: "Adds a textbox above selected notes displaying the note name"
    menuPath: "Plugins.Note Name Above"
    
    onRun: {
        if (!curScore) {
            return;
        }
        
        var selection = curScore.selection;
        if (!selection) {
            return;
        }
        
        var elements = selection.elements;
        if (!elements || elements.length === 0) {
            return;
        }
        
        curScore.startCmd();
        
        var notesProcessed = 0;
        var processedChords = {}; // Track which chords we've already processed
        
        for (var i = 0; i < elements.length; i++) {
            var element = elements[i];
            
            // Check if the element is a note
            if (element.type === Element.NOTE) {
                var note = element;
                
                // Get the chord parent
                var chord = note.parent;
                if (!chord) {
                    continue;
                }
                
                // Skip if we've already processed this chord
                var chordId = chord.tick + "_" + chord.track;
                if (processedChords[chordId]) {
                    continue;
                }
                processedChords[chordId] = true;
                
                // Get the note name using tpc (tonal pitch class) from the selected note
                var tpc = note.tpc;
                var noteName = getNoteName(tpc);
                
                // Create a staff text element
                var staffText = newElement(Element.STAFF_TEXT);
                if (!staffText) {
                    continue;
                }
                
                staffText.text = noteName;
                
                // Position the text above the note
                // offsetY is in sp (spatium units), negative values move up
                staffText.offsetY = -3.5;
                
                // Set the track (staff) to match the note
                staffText.track = chord.track;
                
                // Add the staff text to the note using note.add()
                // This automatically sets the parent property correctly
                note.add(staffText);
                notesProcessed++;
            }
        }
        
        curScore.endCmd();
    }
    
    // Helper function to convert tpc (tonal pitch class) to note name
    function getNoteName(tpc) {
        // Comprehensive mapping of tpc values to note names
        var tpcToName = {
            // Double flats (rare)
            10: "Fbb", 11: "Cbb", 12: "Gbb", 13: "Dbb", 14: "Abb", 15: "Ebb", 16: "Bbb",
            // Flats
            17: "Fb", 18: "Cb", 19: "Gb", 20: "Db", 21: "Ab", 22: "Eb", 23: "Bb",
            // Naturals (standard)
            24: "F", 25: "C", 26: "G", 27: "D", 28: "A", 29: "E", 30: "B",
            // Sharps
            31: "F#", 32: "C#", 33: "G#", 34: "D#", 35: "A#", 36: "E#", 37: "B#",
            // Double sharps (rare)
            38: "F##", 39: "C##", 40: "G##", 41: "D##", 42: "A##", 43: "E##", 44: "B##"
        };
        
        if (tpcToName[tpc]) {
            return tpcToName[tpc];
        }
        
        // Fallback: calculate from tpc using modulo arithmetic
        var baseTpc = 25; // C natural
        var diff = tpc - baseTpc;
        var chromaticIndex = ((diff % 12) + 12) % 12;
        var noteNames = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"];
        return noteNames[chromaticIndex];
    }
}
