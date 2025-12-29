import QtQuick 2.15
import MuseScore 4.0

MuseScore {
    menuPath: "Plugins.Export Pitch Names"
    description: "Exports the pitch names of highlighted section as txt"
    version: "1.0"
    
    onRun: {
        if (!curScore) {
            console.log("No score open");
            return;
        }
        
        // Check if there's a selection
        var selection = curScore.selection;
        if (!selection || selection.elements.length === 0) {
            console.log("No selection found. Please select elements first.");
            return;
        }
        
        console.log("Selection contains " + selection.elements.length + " elements");
        
        // Collect pitch names from selected notes
        var pitchNames = [];
        var processedChords = {}; // Track which chords we've already processed
        
        for (var i = 0; i < selection.elements.length; i++) {
            var element = selection.elements[i];
            
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
                
                // Get the note name using tpc (tonal pitch class)
                var tpc = note.tpc;
                var noteName = getNoteName(tpc);
                
                // Get measure and beat information for context
                var measure = chord.measure;
                var measureNo = measure ? measure.no + 1 : 0; // 1-indexed for display
                var tick = chord.tick;
                
                // Store pitch name with context
                pitchNames.push({
                    name: noteName,
                    measure: measureNo,
                    tick: tick,
                    track: chord.track
                });
            }
        }
        
        if (pitchNames.length === 0) {
            console.log("No notes found in selection.");
            return;
        }
        
        // Sort by measure, then by tick, then by track
        pitchNames.sort(function(a, b) {
            if (a.measure !== b.measure) {
                return a.measure - b.measure;
            }
            if (a.tick !== b.tick) {
                return a.tick - b.tick;
            }
            return a.track - b.track;
        });
        
        // Build the text content
        var textContent = "Pitch Names Export\n";
        textContent += "==================\n\n";
        textContent += "Total notes: " + pitchNames.length + "\n\n";
        textContent += "Pitch Names (in order):\n";
        textContent += "----------------------\n";
        
        var currentMeasure = -1;
        for (var j = 0; j < pitchNames.length; j++) {
            var pitch = pitchNames[j];
            
            // Add measure separator if measure changed
            if (pitch.measure !== currentMeasure) {
                if (currentMeasure !== -1) {
                    textContent += "\n";
                }
                textContent += "Measure " + pitch.measure + ":\n";
                currentMeasure = pitch.measure;
            }
            
            textContent += pitch.name + " ";
        }
        
        textContent += "\n\n";
        textContent += "Pitch Names (comma-separated):\n";
        textContent += "-----------------------------\n";
        var namesOnly = [];
        for (var k = 0; k < pitchNames.length; k++) {
            namesOnly.push(pitchNames[k].name);
        }
        textContent += namesOnly.join(", ");
        textContent += "\n";
        
        // Generate filename with timestamp
        var timestamp = new Date().toISOString().replace(/[:.]/g, "-").substring(0, 19);
        var defaultFileName = "pitch_names_export_" + timestamp + ".txt";
        
        // Write to file
        writeContentToFile(defaultFileName, textContent);
    }
    
    function writeContentToFile(filePath, content) {
        // Output to console with clear formatting for easy copying
        console.log("\n" + "=".repeat(70));
        console.log("PITCH NAMES EXPORT");
        console.log("=".repeat(70));
        console.log(content);
        console.log("=".repeat(70));
        console.log("\nTo save this content to a file:");
        console.log("1. Copy all content between the === markers above");
        console.log("2. Open a text editor (Notepad, TextEdit, etc.)");
        console.log("3. Paste the content");
        console.log("4. Save as: " + filePath);
        console.log("\nThe content has also been saved to MuseScore's log files.");
        console.log("Log location: %LOCALAPPDATA%\\MuseScore\\MuseScore4\\logs\\");
        console.log("Look for the most recent log file to find this export.");
        console.log("\n");
        
        // Try to write using FileIO if available (some MuseScore installations have it)
        try {
            var fileIO = Qt.createQmlObject('
                import QtQuick 2.15
                import FileIO 1.0
                FileIO {
                    source: "' + filePath + '"
                }
            ', this);
            if (fileIO) {
                fileIO.write(content);
                console.log("✓ File successfully written to: " + filePath);
                return;
            }
        } catch (e) {
            // FileIO not available - that's okay, console output is the fallback
        }
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

