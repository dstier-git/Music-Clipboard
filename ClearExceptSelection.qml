import QtQuick 2.15
import MuseScore 3.0

MuseScore {
    menuPath: "Plugins.Clear Except Selection"
    description: "Clears all elements in the score except the highlighted selection"
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
        
        // Create a set of selected element IDs for quick lookup
        var selectedIds = {};
        for (var i = 0; i < selection.elements.length; i++) {
            var selElem = selection.elements[i];
            if (selElem) {
                selectedIds[selElem.id] = true;
                // Also mark parent elements as selected (so we don't remove containers of selected items)
                var parent = selElem.parent;
                while (parent) {
                    selectedIds[parent.id] = true;
                    parent = parent.parent;
                }
            }
        }
        
        // Find the range of measures containing selected elements
        var startMeasureNo = null;
        var endMeasureNo = null;
        
        for (var j = 0; j < selection.elements.length; j++) {
            var elem = selection.elements[j];
            var measure = elem.measure;
            if (measure) {
                if (startMeasureNo === null || measure.no < startMeasureNo) {
                    startMeasureNo = measure.no;
                }
                if (endMeasureNo === null || measure.no > endMeasureNo) {
                    endMeasureNo = measure.no;
                }
            }
        }
        
        // Start undo operation
        curScore.startCmd();
        
        try {
            // Remove measures outside the selection range
            if (startMeasureNo !== null && endMeasureNo !== null) {
                var cursor = curScore.newCursor();
                
                // Remove measures after the selection
                cursor.rewind(1); // Go to end
                var lastMeasure = cursor.measure;
                if (lastMeasure) {
                    while (lastMeasure && lastMeasure.no > endMeasureNo) {
                        curScore.selection.select(lastMeasure, true);
                        curScore.deleteSelection();
                        cursor.rewind(1);
                        lastMeasure = cursor.measure;
                    }
                }
                
                // Remove measures before the selection
                cursor.rewind(0); // Go to start
                var firstMeasure = cursor.measure;
                if (firstMeasure) {
                    while (firstMeasure && firstMeasure.no < startMeasureNo) {
                        curScore.selection.select(firstMeasure, true);
                        curScore.deleteSelection();
                        cursor.rewind(0);
                        firstMeasure = cursor.measure;
                    }
                }
            }
            
            // Now iterate through all remaining elements and remove unselected ones
            var cursor2 = curScore.newCursor();
            cursor2.rewind(0);
            
            // Collect all elements to remove (we need to collect first, then remove)
            var elementsToRemove = [];
            
            while (cursor2.segment) {
                var seg = cursor2.segment;
                
                // Check all elements in this segment
                var element = seg.firstElement;
                while (element) {
                    // Check if this element or any parent is selected
                    var isSelected = false;
                    var checkElem = element;
                    while (checkElem) {
                        if (selectedIds[checkElem.id]) {
                            isSelected = true;
                            break;
                        }
                        checkElem = checkElem.parent;
                    }
                    
                    // Also check children for notes in chords
                    if (!isSelected && element.type === Element.CHORD) {
                        if (element.notes) {
                            for (var n = 0; n < element.notes.length; n++) {
                                if (selectedIds[element.notes[n].id]) {
                                    isSelected = true;
                                    break;
                                }
                            }
                        }
                    }
                    
                    if (!isSelected) {
                        elementsToRemove.push(element);
                    }
                    
                    element = element.next;
                }
                
                cursor2.next();
            }
            
            // Remove all unselected elements
            for (var k = 0; k < elementsToRemove.length; k++) {
                try {
                    curScore.removeElement(elementsToRemove[k]);
                } catch (e) {
                    // Element may have already been removed, continue
                }
            }
            
            curScore.endCmd();
            console.log("Clearing complete. Only selected elements remain.");
            
        } catch (e) {
            curScore.endCmd();
            console.log("Error: " + e);
        }
    }
}

