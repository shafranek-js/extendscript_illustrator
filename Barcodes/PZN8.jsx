#target illustrator

function generatePZN8Barcode() {
    // Create a new document, converted from mm to pt
    var doc = app.documents.add(null, convertMmToPt(200), convertMmToPt(50));

    // Prompt user for PZN8 number
    var userInput = prompt("Enter the 7-digit PZN number:", "");

    // Validate the input
    if (!isValidPZN(userInput)) {
        alert("Please enter a valid 7-digit number.");
        return;
    }

    // Prompt user for size selection
    var size = prompt("Select barcode size:\n1 - Small\n2 - Normal\n3 - Large", "2");
    var barcodeParams = getBarcodeParams(size);

    // Calculate the check digit and validate it
    var checkDigit = calculateCheckDigit(userInput);
    if (checkDigit === 10) {
        alert("Invalid PZN number. Check digit calculation resulted in 10.");
        return;
    }

    var pzn = userInput + checkDigit;

    // Starting position
    var xPos = 0;
    var yPos = doc.height - barcodeParams.height - convertMmToPt(10);

    // Create barcode group
    var barcodeGroup = doc.groupItems.add();

    // Draw the barcode
    var barcodeEndXPos = drawBarcode(barcodeGroup, xPos, yPos, pzn, barcodeParams.singleBarWidth, barcodeParams.doubleBarWidth, barcodeParams.height);

    // Calculate the width of the barcode
    var barcodeWidth = barcodeEndXPos - xPos;

    // Add human-readable text 5 points below the barcode
    var textFrame = addHumanReadableText(barcodeGroup, "PZN- " + pzn, xPos, barcodeWidth, yPos - barcodeParams.height - 5);

    // Convert text to outlines
    textFrame.createOutline();

    // Fit artboard to the barcode group and enlarge by 2.5 mm on each side
    fitArtboardToGroup(doc, barcodeGroup, convertMmToPt(2.5));

    // Fit artboard to window
    fitArtboardToWindow();

    alert("PZN8 Barcode generated successfully!");
}

function getBarcodeParams(size) {
    switch(size) {
        case '1':
            return {
                height: convertMmToPt(7),
                singleBarWidth: convertMmToPt(0.187),
                doubleBarWidth: convertMmToPt(0.187 * 2.5)
            };
        case '2':
            return {
                height: convertMmToPt(10),
                singleBarWidth: convertMmToPt(0.25),
                doubleBarWidth: convertMmToPt(0.25 * 2.5)
            };
        case '3':
            return {
                height: convertMmToPt(20),
                singleBarWidth: convertMmToPt(0.337),
                doubleBarWidth: convertMmToPt(0.337 * 2.5)
            };
        default:
            alert("Invalid size selection. Defaulting to Normal.");
            return {
                height: convertMmToPt(10),
                singleBarWidth: convertMmToPt(0.25),
                doubleBarWidth: convertMmToPt(0.25 * 2.5)
            };
    }
}

function convertMmToPt(mm) {
    return mm / 0.35278;
}

function isValidPZN(pzn) {
    return pzn.length === 7 && !isNaN(pzn);
}

function calculateCheckDigit(pzn) {
    var sum = 0;
    for (var i = 0; i < pzn.length; i++) {
        sum += parseInt(pzn[i]) * (i + 1);
    }
    return sum % 11;
}

function drawBarcode(group, xPos, yPos, pzn, singleBarWidth, doubleBarWidth, barcodeHeight) {
    var code39Map = getCode39Map();
    var drawBars = getDrawBarsFunction(group, yPos, singleBarWidth, doubleBarWidth, barcodeHeight);

    // Add start character "*"
    xPos = drawBars(code39Map["*"], xPos);
    xPos += singleBarWidth; // Space after start character

    // Add start character "-"
    xPos = drawBars(code39Map["-"], xPos);
    xPos += singleBarWidth; // Space after "-" character

    // Loop through PZN digits and generate barcode
    for (var i = 0; i < pzn.length; i++) {
        xPos = drawBars(code39Map[pzn[i]], xPos);
        xPos += singleBarWidth; // Space between characters
    }

    // Add stop character "*"
    xPos = drawBars(code39Map["*"], xPos);

    // Return end position of the barcode
    return xPos;
}

function getCode39Map() {
    return {
        "0": "101001101101",
        "1": "110100101011",
        "2": "101100101011",
        "3": "110110010101",
        "4": "101001101011",
        "5": "110100110101",
        "6": "101100110101",
        "7": "101001011011",
        "8": "110100101101",
        "9": "101100101101",
        "-": "100101011011",
        "*": "100101101101"
    };
}

function getDrawBarsFunction(group, yPos, singleBarWidth, doubleBarWidth, barcodeHeight) {
    return function(pattern, xPos) {
        for (var i = 0; i < pattern.length; i++) {
            var width = singleBarWidth;
            if (i < pattern.length - 1 && pattern[i] === pattern[i + 1]) {
                width = doubleBarWidth;
                i++; // Skip the next character since it's part of this double width
            }
            if (pattern[i] === "1") {
                var rect = group.pathItems.rectangle(yPos, xPos, width, barcodeHeight);
                rect.filled = true;
                rect.stroked = false;
                rect.fillColor = new CMYKColor();
                rect.fillColor.black = 100;
            }
            xPos += width;
        }
        return xPos;
    };
}

function addHumanReadableText(group, textContent, startX, barcodeWidth, top) {
    var text = group.textFrames.add();
    text.contents = textContent;

    // Calculate centered position
    text.textRange.characterAttributes.size = 8;
    text.textRange.characterAttributes.textFont = app.textFonts.getByName("MicrosoftSansSerif");

    // Calculate the width of the text frame
    text.textRange.justification = Justification.CENTER;
    var textFrameWidth = text.width;

    // Center the text frame below the barcode
    var left = startX + (barcodeWidth - textFrameWidth) / 2;

    text.left = left;
    text.top = top;

    return text;
}

function fitArtboardToGroup(doc, group, margin) {
    var bounds = group.visibleBounds;
    var left = bounds[0] - margin;
    var top = bounds[1] + margin;
    var right = bounds[2] + margin;
    var bottom = bounds[3] - margin;
    doc.artboards[0].artboardRect = [left, top, right, bottom];
}

function fitArtboardToWindow() {
    app.activeDocument.artboards.setActiveArtboardIndex(0);
    app.executeMenuCommand('fitin');
}

generatePZN8Barcode();
