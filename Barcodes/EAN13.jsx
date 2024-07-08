#target illustrator 
var PTS_MM = 2.834645;
//var width = 10 * PTS_MM;
//var height = 10 * PTS_MM;

//nexusFont(); /// just delete this line if you are not from Zentiva


var docPreset = new DocumentPreset;
docPreset.width = 60*PTS_MM;
docPreset.height = 10*PTS_MM;
docPreset.units = RulerUnits.Millimeters;
var doc = app.documents.addDocument(DocumentColorSpace.CMYK, docPreset);
docPreset.title  = "Your Title Is Here";

// DIALOG
// ======
var dialog = new Window("dialog"); 
    dialog.text = "Dialog"; 
    dialog.orientation = "row"; 
    dialog.alignChildren = ["center","top"]; 
    dialog.spacing = 10; 
    dialog.margins = 16; 

// GROUP1
// ======
var group1 = dialog.add("group", undefined, {name: "group1"}); 
    group1.orientation = "row"; 
    group1.alignChildren = ["left","center"]; 
    group1.spacing = 10; 
    group1.margins = 0; 

var enterCode = group1.add("statictext", undefined, undefined, {name: "enterCode"}); 
    enterCode.text = "Enter EAN13 code"; 
    enterCode.alignment = ["left","center"]; 

var inputCode = group1.add('edittext {properties: {name: "inputCode"}}'); 
    inputCode.text = "5450557010956"; 
    inputCode.preferredSize.width = 104; 

// GROUP2
// ======
var group2 = dialog.add("group", undefined, {name: "group2"}); 
    group2.orientation = "row"; 
    group2.alignChildren = ["left","center"]; 
    group2.spacing = 10; 
    group2.margins = 0; 

var ok = group2.add("button", undefined, undefined, {name: "ok"}); 
    ok.text = "OK"; 
    
    ok.onClick = main;
    
       
    

var cancel = group2.add("button", undefined, undefined, {name: "cancel"}); 
    cancel.text = "Cancel"; 

dialog.show();

function main() {
    //nexusFont();
    var code = inputCode.text;
    var codeArr = code.split('');
    dialog.close(); 
    

    if (codeArr.length == 13) {
         codeArr.pop();
         //alert("the code was reduced to " + codeArr.length + " symbols");
         var finalCode = codeArr.join("");
         var resultCode = finalCode.concat(eanCheckDigit(finalCode));
         //alert(resultCode);
         //placeText();
         var myTextFrame = app.activeDocument.textFrames.add();
         myTextFrame.contents = (resultCode.toString());
         var textRange = myTextFrame.textRange;
         var charAttributes = textRange.characterAttributes;
         charAttributes.size = 12;

         }
    else {
         //alert("the code has initially " + codeArr.length + " symbols");
         var finalCode = codeArr.join("");
         var resultCode = finalCode.concat(eanCheckDigit(finalCode));
         //alert(resultCode);
         //placeText();
         var myTextFrame = app.activeDocument.textFrames.add();
         myTextFrame.contents = (resultCode.toString());
         var textRange = myTextFrame.textRange;
         var charAttributes = textRange.characterAttributes;
         charAttributes.size = 12;
         }
     createEans();
     app.executeMenuCommand('selectall');
     app.executeMenuCommand("outline");
     app.executeMenuCommand ('Fit Artboard to selected Art');
     //app.activeDocument.activeView.zoom = 5;
     app.executeMenuCommand("deselectall");
     };
    
 
 // Function to return the reverse of a number
function reverse(finalCode) {
    var rev = 0;
    while (finalCode != 0) {
        rev = (rev * 10) + (finalCode % 10);
        finalCode = Math.floor(finalCode / 10);
        }
        return rev;
    }

//Sum all the digits in odd positions. Sum all the digits in even positions and multiply the result by 3. 
//Add the results, and take just the final digit (the ‘units’ digit) of the answer. 
//This is equivalent to taking the answer modulo-10.
function eanCheckDigit(finalCode) {
    finalCode = reverse(finalCode);
    var sumOdd = 0, sumEven = 0, c = 1;
    while (finalCode != 0) {
    // If c is even number then it means
    // digit extracted is at even place
    if (c % 2 == 0)
        sumEven += finalCode % 10;
        else
            sumOdd += finalCode % 10;
        finalCode = Math.floor(finalCode / 10);
        c++;
    }
    var controlDitit = (sumOdd + sumEven*3);
    cdMod = (controlDitit % 10);
    if (cdMod > 0)
        resultDitig = 10 - cdMod;
    else
        resultDitig = cdMod;
    return resultDitig;
    };


//Text
function placeText() {
        var myTextFrame = app.activeDocument.textFrames.add();
        //myTextFrame.position = [0, -myTextFrame.height+3];
        alert(resultCode);
        myTextFrame.contents = (resultCode.toString());
        
        var textRange = myTextFrame.textRange;
        var charAttributes = textRange.characterAttributes;
        
        charAttributes.size = 12;
        
        }
    
    
//Let's create EAN
function createEans() {
    var doc = app.activeDocument;
    app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
    var fontName = "OCRB";
    //var fontName = "OCRBStd";
    var barcodeTextArr = [];
    if (fontAvailable(fontName)) {
        //Select 13 digits and replace it with a barcode
        for (var i = 0; i < doc.textFrames.length; i++) {
            if (doc.textFrames[i].contents.match(/^\d{13}$/gi)) {
                var textParent = doc.textFrames[i].parent;
                //If layer with text is locked or invisible the script will ignore them
                if (textParent.locked == true || textParent.visible == false) {
                    continue;
                };
                if (doc.textFrames[i].kind == TextType.AREATEXT) {
                    areaTextJust(doc.textFrames[i]);
                    doc.textFrames[i].convertAreaObjectToPointObject();
                }
                var myCode = doc.textFrames[i].contents;
                if (CheckDigit(myCode) != myCode[12]) {
                    alert("The barcode " + myCode + " does not have the right checksum. The checksum digit must be " + CheckDigit(myCode));
                }
                else barcodeTextArr.push(doc.textFrames[i]);
            }
        };
        for (var i = 0; i < barcodeTextArr.length; i++) {
            var barcodeText = barcodeTextArr[i];
            var textParent = barcodeText.parent;
            var EANGroup = textParent.groupItems.add();
            justText(barcodeText);
            var matrix = barcodeText.matrix;
            var Angle = -180 / Math.PI * Math.atan2(matrix.mValueC, matrix.mValueD)
            var Width = barcodeText.textRange.size * 18;
            var Height = Width * 0.2;
            var newX = barcodeText.anchor[0] - Width * 0.01;
            var newY = barcodeText.anchor[1] //+ Height * 0.344;
            var roto_X = barcodeText.anchor[0];
            var roto_Y = barcodeText.anchor[1];
            var delta = Width / 20;
            var EANtext = barcodeText.contents;
            var barColor = make_cmyk(0, 0, 0, 100);
            var barcode = CreateBarcode(Width, Height, EANtext, EANGroup, barColor, barcodeText);
            var newY = barcodeText.anchor[1] + EANGroup.height;
            barcode.position = [newX, newY];
            rotate_around_point(barcode, roto_Y, roto_X, Angle);
        }
    }
    else {
        alert("Looks, like the " + fontName + " font is not available in you computer. Please install " + fontName + " font and try again.");
    };
    for (var i = 0; i < barcodeTextArr.length; i++) {
        barcodeTextArr[i].remove();
    };

    ////////////////          FUNCTIONS        /////////////////////////////////////////////////////////////////
    function bcRenderChar(x, y, w, h, col, gr, gapD) {
        this.x = x;
        this.y = y;
        this.w = w;
        this.h = h;
        this.col = col;
        this.gr = gr;
        this.gapD = gapD;
        this.L = {
            "0": [3, 2, 6, 1],
            "1": [2, 2, 6, 1],
            "2": [2, 1, 5, 2],
            "3": [1, 4, 6, 1],
            "4": [1, 1, 5, 2],
            "5": [1, 2, 6, 1],
            "6": [1, 1, 3, 4],
            "7": [1, 3, 5, 2],
            "8": [1, 2, 4, 3],
            "9": [3, 1, 5, 2]
        }
        this.G = {
            "0": [1, 1, 4, 3],
            "1": [1, 2, 5, 2],
            "2": [2, 2, 5, 2],
            "3": [1, 1, 6, 1],
            "4": [2, 3, 6, 1],
            "5": [1, 3, 6, 1],
            "6": [4, 1, 6, 1],
            "7": [2, 1, 6, 1],
            "8": [3, 1, 6, 1],
            "9": [2, 1, 4, 3]
        }
        this.dictL = {
            "0": "LLLLLL",
            "1": "LLGLGG",
            "2": "LLGGLG",
            "3": "LLGGGL",
            "4": "LGLLGG",
            "5": "LGGLLG",
            "6": "LGGGLL",
            "7": "LGLGLG",
            "8": "LGLGGL",
            "9": "LGGLGL"
        }
        this.dictR = {
            "sep": [1, 1, 3, 1],
            "0": [0, 3, 5, 1],
            "1": [0, 2, 4, 2],
            "2": [0, 2, 3, 2],
            "3": [0, 1, 5, 1],
            "4": [0, 1, 2, 3],
            "5": [0, 1, 3, 3],
            "6": [0, 1, 2, 1],
            "7": [0, 1, 4, 1],
            "8": [0, 1, 3, 1],
            "9": [0, 3, 4, 1]
        }
        this.drawLeft = function (content) {
            var mySeq = this.dictL[content[0]];
            for (var i = 1; i < content.length; i++) {
                var myLG = mySeq[i - 1];
                if (myLG == "L") {
                    var parameters = this.L[content[i]];
                }
                else if (myLG == "G") {
                    var parameters = this.G[content[i]];
                }
                rect = this.gr.pathItems.rectangle(this.y, this.x + this.w * parameters[0], this.w * parameters[1], this.h);
                rect.stroked = false;
                rect.filled = true;
                rect.fillColor = this.col;
                rect = this.gr.pathItems.rectangle(this.y, this.x + this.w * parameters[2], this.w * parameters[3], this.h);
                rect.stroked = false;
                rect.filled = true;
                rect.fillColor = this.col;
                this.x = this.x += gapD;
            }
        }
        this.draw = function (textChar) {
            var parameters = this.dictR[textChar];
            rect = this.gr.pathItems.rectangle(this.y, this.x + this.w * parameters[0], this.w * parameters[1], this.h);
            rect.stroked = false;
            rect.filled = true;
            rect.fillColor = this.col;
            rect = this.gr.pathItems.rectangle(this.y, this.x + this.w * parameters[2], this.w * parameters[3], this.h);
            rect.stroked = false;
            rect.filled = true;
            rect.fillColor = this.col;
        }
    };

    function fontAvailable(myName) {
        //Function checks if font exists
        var myFont = true;
        try {
            var myFont = textFonts.getByName(myName);
        }
        catch (e) {
            var myFont = false;
        }
        return myFont;
    };

    function make_cmyk(c, m, y, k) {
        var colorRef = new CMYKColor();
        colorRef.cyan = c;
        colorRef.magenta = m;
        colorRef.yellow = y;
        colorRef.black = k;
        return colorRef;
    };

    function CreateBarcode(Width, Height, BarcodeNr, Group, barColor, textArtRange) {
        // Function creates barcode
        //adjust width for correct fit
        var block = Width * 0.00346;
        var blockHeightExtra = Height * 1.07; //heigtht of barcode
        var fontSize = block * 10;//font size
        var zX = 0;
        var zY = 0;
        var BarcodeNr = textArtRange.contents;
        var gapD = block * 7;
        var whiteRect = Group.pathItems.rectangle(zY, zX - block * 10, block * 118, blockHeightExtra * 1.16);
        whiteRect.stroked = false;
        whiteRect.filled = true;
        whiteRect.fillColor = make_cmyk(0, 0, 0, 0);
        var bcRenderObject = new bcRenderChar(zX, zY, block, blockHeightExtra, barColor, Group, gapD);
        bcRenderObject.draw("sep");
        bcRenderObject.x += block * 4;
        bcRenderObject.h = Height;
        bcRenderObject.drawLeft(BarcodeNr.substring(0, 7));
        bcRenderObject.h = blockHeightExtra;
        bcRenderObject.draw("sep");
        bcRenderObject.x += block * 5;
        bcRenderObject.h = Height;
        for (var j = 7; j < 13; j++) {
            bcRenderObject.draw(BarcodeNr[j]);
            bcRenderObject.x += gapD;
        }
        bcRenderObject.h = blockHeightExtra;
        bcRenderObject.draw("sep");
        var topPos = - Height * 1.03;
        pointText(Group, fontSize, BarcodeNr.charAt(0), topPos, block - block * 9, fontName, 1);
        pointText(Group, fontSize, BarcodeNr.substring(1, 7), topPos, block + block * 3, fontName, 1);
        pointText(Group, fontSize, BarcodeNr.substring(7, 13), topPos, block + block * 49, fontName, 1);
        return Group;
    };

    function pointText(Group, fontSize, charNr, topPos, leftPos, fontName, fontScale) {
        //Creates barcode text
        var pointTextRef = Group.textFrames.add();
        pointTextRef.textRange.size = fontSize;
        pointTextRef.contents = charNr;
        pointTextRef.position = [leftPos, topPos];
        pointTextRef.textRange.characterAttributes.textFont = textFonts.getByName(fontName);
        pointTextRef.textRange.characterAttributes.size = pointTextRef.textRange.characterAttributes.size * fontScale;
        return pointTextRef;
    };

    function CheckDigit(myCode) {
        //Calculate checksum of a 13 digit number and compare to last digit
        //Number must be 13 digit long, or the calculation will be wrong
        var mySum = 0;
        for (var j = 0; j < myCode.length - 1; j = j + 1) {
            //Determine weight to multiply to current digit
            if (j % 2 == 0) {
                var weight = 1
            }
            else {
                var weight = 3
            }
            var myNumber = myCode[j] * weight;
            mySum = mySum + myNumber;
        }
        checkDigit = Math.ceil(mySum / 10) * 10 - mySum
        return checkDigit;
    };

    function rotate_around_point(obj, roto_X, roto_Y, angle) {
        // Utility that rotates an object around an arbitrary point
        //Translate from point to document origin 
        obj.translate(-roto_Y, -roto_X);
        //Rotate around document origin
        obj.rotate(angle, true, true, true, true, Transformation.DOCUMENTORIGIN);
        //Translate back
        obj.translate(roto_Y, roto_X);
    };

    function justText(mySelection) {
        //Function justificates text without moving it around
        var locArr = new Array();
        var myJust = Justification.FULLJUSTIFYLASTLINELEFT;
        locArr = addtoList(mySelection, locArr);
        for (all in locArr) {
            locArr[all][0].story.textRange.justification = myJust;
            locArr[all][0].top = locArr[all][1];
            locArr[all][0].left = locArr[all][2];
        };
        function addtoList(obj, myArray) {
            var temp = new Array();
            temp[0] = obj;
            temp[1] = obj.top;
            temp[2] = obj.left;
            myArray.push(temp);
            return myArray;
        }
    };

    function areaTextJust(myText) {
        //If FULLJUSTIFYLASTLINE changes to Justification
        //NB:Justification.LEFT doesn´t work
        if (myText.kind == TextType.AREATEXT) {
            if (myText.paragraphs[0].paragraphAttributes.justification == Justification.FULLJUSTIFYLASTLINECENTER) {
                myText.paragraphs[0].paragraphAttributes.justification = Justification.CENTER;
            }
            else if (myText.paragraphs[0].paragraphAttributes.justification == Justification.FULLJUSTIFYLASTLINERIGHT) {
                myText.paragraphs[0].paragraphAttributes.justification = Justification.RIGHT;
            }
            else if (myText.paragraphs[0].paragraphAttributes.justification == Justification.FULLJUSTIFYLASTLINELEFT) {
                myText.paragraphs[0].paragraphAttributes.justification = Justification.LEFT;
            }
        }
    }
};



////// Только для Zentiva
function nexusFont() {
    userNames = $.getenv("USERNAME");

    if (userNames != "pavel") {
        var pathToFile = "~/Popovic, Pavlo CZ - AW SCRIPTS/Illustrator/nexusfont/Run NexusFont.bat";
        var fileObj = new File(pathToFile + "/Run NexusFont");
        }
    else{
        var pathToFile = "~/AW SCRIPTS/Illustrator/nexusfont/Run NexusFontPavel.bat";
        var fileObj = new File(pathToFile + "/Run NexusFontPavel.bat");
        }
    
    var parent = fileObj.parent.fsName; // >> /path/to
    alert(parent);
    var batFile = new File(parent);
    batFile.execute();
};