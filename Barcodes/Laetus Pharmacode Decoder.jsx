#target illustrator

function decodeLaetusPharmacode() {
    var doc = app.activeDocument;
    try {
        var barcodeLayer = doc.layers.getByName("Layer 1");
    } catch (e) {
        alert("Error: The specified layer does not exist.");
        return;
    }
    var bars = [];
    // Collect bars and their positions
    for (var i = 0; i < barcodeLayer.pathItems.length; i++) {
        var bar = barcodeLayer.pathItems[i];
        bars.push({
            bar: bar,
            position: bar.position[0],
            width: bar.width
        });

    }
    // Sort bars by their horizontal position to process them left to right
    bars.sort(function(a, b) {
        return a.position - b.position;
    });
    // Calculate gaps between bars as "white bars"
    var gaps = [];
    for (var i = 0; i < bars.length - 1; i++) {
        var gap = bars[i + 1].position - (bars[i].position + bars[i].width);
        gaps.push(gap);
    }

    // Determine the threshold for thin vs thick gaps
    var gapThreshold = Math.round(calculateGapThreshold(gaps) * 10) / 10;

    var widths = [];
    for (var i = 0; i < bars.length; i++) {
        widths.push(bars[i].width);
    }
    var thinWidth = Math.round(Math.min.apply(null, widths) * 100) / 100;
    var thickWidth = Math.round(Math.max.apply(null, widths) * 100) / 100;
    var additionalInfo = "";
    var binarySequence = "";
    var thinWidthCount = 0,
        thickWidthCount = 0;
    var isUniformThin = true,
        isUniformThick = true;
    var lastThinWidth = null,
        lastThickWidth = null;
    var errorWidth = false;
    var sameGap = true;

    for (var i = 0; i < bars.length; i++) {
        var currentWidth = Math.round(bars[i].width * 100) / 100;
        binarySequence += (currentWidth === thinWidth) ? "0" : "1";

        if (currentWidth === thinWidth) {
            if (lastThinWidth !== null && lastThinWidth !== currentWidth) {
                isUniformThin = false;
            }
            lastThinWidth = currentWidth;
            thinWidthCount++;
        } else if (currentWidth === thickWidth) {
            if (lastThickWidth !== null && lastThickWidth !== currentWidth) {
                isUniformThick = false;
            }
            lastThickWidth = currentWidth;
            thickWidthCount++;
        } else {
            errorWidth = true;
        }
        if (i < bars.length - 1) {
            if (Math.round(gaps[i] * 10) / 10 != gapThreshold) {
                sameGap = false;
            }

        }
    }

    var decimalValue = decode(binarySequence);

    if (errorWidth === true) {
        additionalInfo += "Bars' widths are not consistent!\n";
    } else {
        if (isUniformThin && thinWidthCount > 0) {
            additionalInfo += "Thin bars width: " + convertPointsToMM(thinWidth) + " mm.\n";
        }
        if (isUniformThick && thickWidthCount > 0 && thickWidth !== thinWidth) {
            additionalInfo += "Thick bars width: " + convertPointsToMM(thickWidth) + " mm.\n";
        }
    }
    if (sameGap === true) {
        additionalInfo += "Gaps width: " + convertPointsToMM(gapThreshold) + " mm.";
    } else {
        additionalInfo += "Spaces' widths are not consistent!";
    }
    var textFrame = findOrCreateTextFrame(doc, barcodeLayer, binarySequence, decimalValue, additionalInfo);
    textFrame.contents = "Binary Value: " + binarySequence + "\nDecimal Value: " + decimalValue + "\n" + additionalInfo;
    textFrame.left = barcodeLayer.pathItems[0].position[0];
    textFrame.top = barcodeLayer.pathItems[0].position[1] - 30;
}

function decode(encodedString) {
    var result = 0;
    for (var i = 0; i < encodedString.length; i++) {
        result *= 2;
        if (encodedString.charAt(i) === '1') {
            result += 2;
        } else {
            result += 1;
        }
    }
    return result;
}

function findOrCreateTextFrame(doc, barcodeLayer, binarySequence, decimalValue, additionalInfo) {
    var frame;
    var found = false;
    for (var i = 0; i < doc.textFrames.length; i++) {
        frame = doc.textFrames[i];
        if (frame.contents.indexOf("Binary Value:") !== -1 && frame.contents.indexOf("Decimal Value:") !== -1) {
            found = true; // Indicate that the existing text frame has been found
            break; // Exit the loop as we've found our text frame
        }
    }

    // If no suitable text frame is found, create a new one
    if (!found) {
        frame = doc.textFrames.add();
    }

    // Set the font and size for the text frame
    frame.textRange.characterAttributes.textFont = app.textFonts.getByName("ArialMT");
    frame.textRange.characterAttributes.size = 4	;

    return frame;
}

function convertPointsToMM(points) {
    return (points * 0.352778).toFixed(1);
}

function calculateGapThreshold(gaps) {
    var total = 0;
    for (var i = 0; i < gaps.length; i++) {
        total += gaps[i];
    }
    return total / gaps.length; // Return the average gap size
}


function ungroupLayer1() {
    var doc = app.activeDocument; // Get the active document
    try {
        var layer = doc.layers.getByName("Layer 1"); // Try to get Layer 1
        ungroupItems(layer);
    } catch (e) {
        alert("Error: " + e.message);
    }
}

function ungroupItems(parentItem) {
    for (var i = parentItem.pageItems.length - 1; i >= 0; i--) {
        var item = parentItem.pageItems[i];
        // Check if the item is a group
        if (item.typename === "GroupItem") {
            ungroupItems(item); // Recursively ungroup nested groups
            while (item.pageItems.length > 0) { // Ensure all nested items are moved
                item.pageItems[0].move(parentItem, ElementPlacement.PLACEATEND);
            }
            item.remove(); // Remove the now-empty group container
        }
    }
}

function releaseCompounds() {
    function dismantle(group) {
        var parent = group.parent;
        var containerType = (group.typename == 'CompoundPathItem') ? "pathItems" : "pageItems";

        // Find the correct position for insertion
        var insertPosition = null; // This will be a PageItem reference
        for (var i = 0; i < parent.pageItems.length; i++) {
            if (parent.pageItems[i] == group) {
                // If the group is the last item, insertPosition remains null, indicating insertion at the end
                if (i < parent.pageItems.length - 1) {
                    insertPosition = parent.pageItems[i + 1];
                }
                break;
            }
        }

        // Correctly move each item to the new location
        for (var j = group[containerType].length - 1; j >= 0; j--) {
            var item = group[containerType][j];
            if (insertPosition) {
                item.move(insertPosition, ElementPlacement.PLACEBEFORE);
            } else {
                item.move(parent, ElementPlacement.INSIDE);
            }
        }

        group.remove(); // Remove the empty group
    }

    function breakUpLayer(layer) {
        while (layer.groupItems.length > 0 || layer.compoundPathItems.length > 0) {
            for (var i = layer.groupItems.length - 1; i >= 0; i--) {
                dismantle(layer.groupItems[i]);
            }

            for (var j = layer.compoundPathItems.length - 1; j >= 0; j--) {
                dismantle(layer.compoundPathItems[j]);
            }

            redraw(); // Necessary to refresh the UI and document structure
        }
    }

    if (app.name == "Adobe Illustrator" && app.documents.length > 0) {
        var doc = app.activeDocument;
        var firstLayer = doc.layers[0];
        breakUpLayer(firstLayer);
    }
}

releaseCompounds();
ungroupLayer1();
decodeLaetusPharmacode();