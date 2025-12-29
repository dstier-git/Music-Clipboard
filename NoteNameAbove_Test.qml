import QtQuick 2.0
import MuseScore 3.0

MuseScore {
    version: "1.0"
    description: "TEST: Adds a textbox above selected notes displaying the note name"
    menuPath: "Plugins.Note Name Above Test"
    
    onRun: {
        // Simple test - just try to add text to see if the API works
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
        
        // Try to process the first element only
        var element = elements[0];
        
        if (element.type === Element.NOTE) {
            var note = element;
            var tpc = note.tpc;
            var noteName = getNoteName(tpc);
            
            var chord = note.parent;
            if (chord) {
                var segment = chord.parent;
                if (segment) {
                    var staffText = newElement(Element.STAFF_TEXT);
                    if (staffText) {
                        staffText.text = noteName;
                        staffText.offsetY = -3.5;
                        staffText.track = chord.track;
                        segment.add(staffText);
                    }
                }
            }
        }
        
        curScore.endCmd();
    }
    
    function getNoteName(tpc) {
        var tpcToName = {
            10: "Fbb", 11: "Cbb", 12: "Gbb", 13: "Dbb", 14: "Abb", 15: "Ebb", 16: "Bbb",
            17: "Fb", 18: "Cb", 19: "Gb", 20: "Db", 21: "Ab", 22: "Eb", 23: "Bb",
            24: "F", 25: "C", 26: "G", 27: "D", 28: "A", 29: "E", 30: "B",
            31: "F#", 32: "C#", 33: "G#", 34: "D#", 35: "A#", 36: "E#", 37: "B#",
            38: "F##", 39: "C##", 40: "G##", 41: "D##", 42: "A##", 43: "E##", 44: "B##"
        };
        
        if (tpcToName[tpc]) {
            return tpcToName[tpc];
        }
        
        var baseTpc = 25;
        var diff = tpc - baseTpc;
        var chromaticIndex = ((diff % 12) + 12) % 12;
        var noteNames = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"];
        return noteNames[chromaticIndex];
    }
}

