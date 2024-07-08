#target illustrator 
var docPreset = new DocumentPreset;
docPreset.width = 10;
docPreset.height = 10;
docPreset.units = RulerUnits.Millimeters;
var doc = app.documents.addDocument(DocumentColorSpace.CMYK, docPreset);

var PTS_MM = 2.834645;
var width = 10 * PTS_MM;
var height = 10 * PTS_MM;

var artLayer = doc.layers[0];

//считываем данные из ini файла, если существуем
var readIni = File("~/Desktop/Laetus Pharmacode/PharmaCode.ini");    
    if (readIni.exists){ 
    readIni.open("r");
    var fileContentsString = readIni.read();
    readIni.close();
    var arr = fileContentsString.split("\n");
    }    
else {
    arr = [100,8,0.5,1.5,1];  
    }

//

// DIALOG
// ======
var dialog = new Window("dialog"); 
  dialog.text = "Laetus Pharma Code"; 
  dialog.orientation = "row"; 
  dialog.alignChildren = ["left","top"]; 
  dialog.spacing = 10; 
  dialog.margins = 16; 

// GROUP1
// ======
var group1 = dialog.add("group", undefined, {name: "group1"}); 
  group1.orientation = "column"; 
  group1.alignChildren = ["fill","top"]; 
  group1.spacing = 10; 
  group1.margins = 0; 

// PANEL1
// ======
var panel1 = group1.add("panel", undefined, undefined, {name: "panel1"}); 
  panel1.text = "Parameters"; 
  panel1.preferredSize.height = 75; 
  panel1.orientation = "column"; 
  panel1.alignChildren = ["left","top"]; 
  panel1.spacing = 10; 
  panel1.margins = 10; 

// GROUP2
// ======
var group2 = panel1.add("group", undefined, {name: "group2"}); 
  group2.orientation = "row"; 
  group2.alignChildren = ["left","center"]; 
  group2.spacing = 10; 
  group2.margins = 0; 

var statictext1 = group2.add("statictext", undefined, undefined, {name: "statictext1"}); 
  statictext1.text = "Code:"; 

var edittext1 = group2.add('edittext {properties: {name: "edittext1"}}'); 
  edittext1.text = arr[0]; ///////////////////////////////////////////////////////////////////////////////////
  edittext1.preferredSize.width = 243; 

// PANEL1
// ======
var checkbox1 = panel1.add("checkbox", undefined, undefined, {name: "checkbox1"}); 
  checkbox1.text = "Decimal input"; 
  checkbox1.value = true;   
var prevent = function (e){e.preventDefault();}
  checkbox1.addEventListener("click",prevent);
  checkbox1.onClick = function(event){
      checkbox1.value ? edittext1.text = decode(edittext1.text).data : edittext1.text = encode(edittext1.text).data;
      }

// PANEL2
// ======
var panel2 = group1.add("panel", undefined, undefined, {name: "panel2"}); 
  panel2.text = "Additional Parameters"; 
  panel2.preferredSize.height = 160; 
  panel2.orientation = "column"; 
  panel2.alignChildren = ["right","top"]; 
  panel2.spacing = 10; 
  panel2.margins = 10; 

// GROUP3
// ======
var group3 = panel2.add("group", undefined, {name: "group3"}); 
  group3.orientation = "row"; 
  group3.alignChildren = ["left","center"]; 
  group3.spacing = 10; 
  group3.margins = 0; 

var statictext2 = group3.add("statictext", undefined, undefined, {name: "statictext2"}); 
  statictext2.text = "Bar Height (mm): "; 
  statictext2.alignment = ["left","center"]; 

var edittext2 = group3.add('edittext {justify: "right", properties: {name: "edittext2"}}'); 
  edittext2.text = arr[1]; 
  edittext2.preferredSize.width = 50; 
  edittext2.alignment = ["left","fill"]; 

// GROUP4
// ======
var group4 = panel2.add("group", undefined, {name: "group4"}); 
  group4.orientation = "row"; 
  group4.alignChildren = ["left","center"]; 
  group4.spacing = 10; 
  group4.margins = 0; 

var statictext3 = group4.add("statictext", undefined, undefined, {name: "statictext3"}); 
  statictext3.text = "Narrow Bar (mm):"; 
  statictext3.alignment = ["left","center"]; 

var edittext3 = group4.add('edittext {justify: "right", properties: {name: "edittext3"}}'); 
  edittext3.text = arr[2]; 
  edittext3.preferredSize.width = 50; 
  edittext3.alignment = ["left","fill"]; 

// GROUP5
// ======
var group5 = panel2.add("group", undefined, {name: "group5"}); 
  group5.orientation = "row"; 
  group5.alignChildren = ["left","center"]; 
  group5.spacing = 10; 
  group5.margins = 0; 

var statictext4 = group5.add("statictext", undefined, undefined, {name: "statictext4"}); 
  statictext4.text = "Wide Bar (mm):"; 

var edittext4 = group5.add('edittext {justify: "right", properties: {name: "edittext4"}}'); 
  edittext4.text = arr[3]; 
  edittext4.preferredSize.width = 50; 
  edittext4.alignment = ["left","fill"]; 

// GROUP6
// ======
var group6 = panel2.add("group", undefined, {name: "group6"}); 
  group6.orientation = "row"; 
  group6.alignChildren = ["left","center"]; 
  group6.spacing = 10; 
  group6.margins = 0; 

var statictext5 = group6.add("statictext", undefined, undefined, {name: "statictext5"}); 
  statictext5.text = "Gap (mm): "; 

var edittext5 = group6.add('edittext {justify: "right", properties: {name: "edittext5"}}'); 
  edittext5.text = arr[4];
  edittext5.preferredSize.width = 50; 
  edittext5.alignment = ["left","fill"]; 

// GROUP7
// ======
var group7 = dialog.add("group", undefined, {name: "group7"}); 
  group7.orientation = "column"; 
  group7.alignChildren = ["fill","top"]; 
  group7.spacing = 10; 
  group7.margins = 0; 

var ok = group7.add("button", undefined, undefined, {name: "ok"}); 
  ok.text = "OK"; 
  ok.onClick = mainFunc;

var cancel = group7.add("button", undefined, undefined, {name: "cancel"}); 
  cancel.text = "Cancel"; 

dialog.show();

//Основная функция
function mainFunc() {
    
    if(checkbox1.value == false){
        checkbox1.value = true;
        edittext1.text = decode(edittext1.text).data;
        }
    
    dialog.close();    
    drawCode();
    
    app.executeMenuCommand('selectall');
    distribute();
    app.executeMenuCommand("group");
    

    

    placeText();
    //app.copy();
    
    app.executeMenuCommand ('Fit Artboard to selected Art');
    changeArtboard();
    app.executeMenuCommand ('fitall');
    
    
    app.executeMenuCommand("deselectall");
    alignTextToBottom();
    app.executeMenuCommand("deselectall");
    
    //app.preferences.setStringPreference("myScriptPref", JSON_strngs);
    
    
    //app.activeDocument.activeView.zoom = 10;
    //app.activeDocument.views[0].zoom = 10;
    saveToPharmaFolder();
    var original_file = app.activeDocument.fullName; 
    saveIni();
    app.activeDocument.close(); 
    
    app.open (File (original_file));
}


///Сохраняем в папку Laetus Pharmacode
function saveToPharmaFolder() {
    var f = new Folder('~/Desktop/Laetus Pharmacode/');
    if (!f.exists) {f.create();}    
    doc.saveAs(File('~/Desktop/Laetus Pharmacode/'+ edittext1.text + '.ai'));  
    //var original_file = app.activeDocument.fullName; 
    }
     



//Кодируем число
function encode(z) {
    var result = "";
    while (!isNaN(z) && z != 0) {
        if (z % 2 === 0) {result = "1" + result;
            z = (z - 2) / 2;} 
        else {result = "0" + result; 
            z = (z - 1) / 2;}}

    return {data: result};
    //edittext1.text = encode(edittext1.text).data;
    //alert (encode(edittext1.text).data);
    }


//Рисуем узкий прямоугольник
function drawNarrowSquare() {
    var rect = artLayer.pathItems.rectangle(0, 0, parseFloat(edittext3.text)* PTS_MM, parseFloat(edittext2.text)* PTS_MM);
    rect.fillColor = makeColorCMYK(0,0,0,100);
    rect.stroked = false;
    }


//Рисуем широкий прямоугольник
function drawWideSquare() {
    var rect = artLayer.pathItems.rectangle(0, 1.5, parseFloat(edittext4.text)* PTS_MM, parseFloat(edittext2.text)* PTS_MM);
    rect.fillColor = makeColorCMYK(0,0,0,100);
    rect.stroked = false;
    }
// Определяем цвет заливки прямоугольников
function makeColorCMYK(c,m,y,k){
    var ink = new CMYKColor();
    ink.cyan   = c;
    ink.magenta = m;
    ink.yellow  = y;
    ink.black  = k;
    return ink;
}

//Рисуем линию
function drawLine() {
   var line = doc.pathItems.add();
   line.strokeWidth = 0.5
   line.setEntirePath(Array(Array(0, 0), Array(0, 30)));
   rect.fillColor = false;
   rect.stroked = makeColorCMYK(0,0,0,100);
   }

//Декодируем фармакод
function decode() {
    var listZero = [ 68719476736, 34359738368, 17179869184,  8589934592, 4294967296, 2147483648, 1073741824,  536870912, 268435456, 134217728,  67108864, 33554432, 16777216,  8388608, 4194304, 2097152, 1048576,  524288, 262144, 131072,  65536, 32768, 16384,  8192, 4096, 2048, 1024,  512, 256, 128,  64, 32, 16,  8, 4, 2, 1];
    var listOne  = [137438953472, 68719476736, 34359738368, 17179869184, 8589934592, 4294967296, 2147483648, 1073741824, 536870912, 268435456, 134217728, 67108864, 33554432, 16777216, 8388608, 4194304, 2097152, 1048576, 524288, 262144, 131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2];
    var result = 0;
    var userPharmacode = edittext1.text;
    var countb = userPharmacode.length;
    var d = userPharmacode.split('');
    //alert ("code length is " + countb);
         
    while (countb > 0) {
        e = parseInt(d.pop());
        if ( e > 0) 
            {
            var lastEntryZero = listZero.pop();
            var lastEntryOne = listOne.pop();
            result = result + parseInt(lastEntryOne);
            countb = countb-1;
            }
        else 
            {
            var lastEntryOne = listOne.pop();
            var lastEntryZero = listZero.pop();
            result= result + parseInt(lastEntryZero);
            countb = countb-1;
                }
            }
    return {data: result};
}

//Рисуем полоски
function drawCode() {
    userInput = encode(edittext1.text).data;
    var userInputCount = userInput.length;
    var userInputArray = userInput.split('');
    //var userInputArrayRev = userInputArray.reverse();
    for (i = 0; i < userInput.length; i++) {
         checkvalue = (userInputArray.pop());
         if (checkvalue > 0) {drawWideSquare();}
         else {drawNarrowSquare();}
        }
    }


//Дистрибутим полоски
function distribute() {
    var doc = activeDocument;
    var selx = doc.selection;
    if(selx.length ==0){alert("You must select objects to distribute.");}
    else{makeGrid(selx);}

    function makeGrid(sel)
    {
        var objectsCentered = true;
        if(objectsCentered){
             var newGroup = app.activeDocument.groupItems.add();
        }
        var maxW = maxH = currentX = currentY  = maxRowH = 0;
        var padding = "not valid";
        while (isNaN(padding)){
            padding = parseFloat(edittext5.text)* PTS_MM; /////////////////////////////////////////////////////////
        }
        var layout ='h';
        layout = 'H';
        var gridCols =  (layout.toLowerCase() == "h") ? sel.length : 1;
               gridCols = (layout.toLowerCase() == "g" )  ?  Math.round(Math.sqrt(sel.length))  : gridCols; 
        for(var e=0, slen=sel.length;e<slen;e++)
        {
            if(objectsCentered){
                    // ::Add to group
                    sel[e].moveToBeginning( newGroup );
            }
            //   :::SET POSITIONS:::
            sel[e].top = currentY;
            sel[e].left = currentX;
            //  :::DEFINE X POSITION:::
            currentX += (sel[e].width + padding);
            var itembottom = (sel[e].top-sel[e].height);
            maxRowH = itembottom <  maxRowH ? itembottom : maxRowH;
            if((e % gridCols) == (gridCols - 1))
            { 
                currentX = 0;    
                maxH =  (maxRowH);
                //  :::DEFINE Y POSITION:::
                currentY  = maxH-padding; 
                maxRowH=0;
            }
        }
        if(objectsCentered){ 
                newGroup.top = -( doc.height/2) + newGroup.height/2;
                newGroup.left = (doc.width/2)-newGroup.width/2;
                //   :::UNGROUP:::
                var sLen=sel.length;
                while(sLen--)
                {
                    sel[sLen].moveToBeginning( doc.activeLayer );
                }
        }
    }
}

//Text
function placeText() {
        var myTextFrame = doc.textFrames.add();
        myTextFrame.position = [0, -myTextFrame.height+3];
        myTextFrame.contents = edittext1.text;
        
        var textRange = myTextFrame.textRange;
        var charAttributes = textRange.characterAttributes;
        
        charAttributes.size = 7;
        
        }
    
    
//Меняем размер артборда
function changeArtboard() {
    var userWidth = 40; 
    var userDepth = 20;// + parseInt(statictext2.text);
    doc.artboards.setActiveArtboardIndex(0);
    var middlePanelWidth = userWidth * PTS_MM;
    var middlePanelHeight = userDepth * PTS_MM;
    var width  = middlePanelWidth;
    var height = middlePanelHeight;
    for (i=0; i<doc.artboards.length; i++) {
            var abBounds = doc.artboards[0].artboardRect; // left, top, right, bottom
            var ableft   = abBounds[0];
            var abtop    = abBounds[1];
            var abwidth  = abBounds[2] - ableft;
            var abheight = abtop - abBounds[3];
            var abctrx   = abwidth / 2 + ableft;
            var abctry   = abtop - abheight / 2;
            var ableft   = (abctrx - width  / 2);
            var abtop    = abctry + height / 2;
            var abright  = abctrx + width  / 2;
            var abbottom = abctry - height / 2;
            doc.artboards[0].artboardRect = [ableft, abtop, abright, abbottom];
            }
        }
    
//сохраняем ini файл
function saveIni() {
    var file = new File('~/Desktop/Laetus Pharmacode/' + "PharmaCode.ini");   
    file.open("w");
    file.writeln(edittext1.text);
    file.writeln(edittext2.text);
    file.writeln(edittext3.text);
    file.writeln(edittext4.text);
    file.writeln(edittext5.text);
    file.close();
    }

// текст к нижнему краю страницы
function alignTextToBottom() {
    doc.textFrames[0].selected = true;  
    
    app.coordinateSystem = CoordinateSystem.ARTBOARDCOORDINATESYSTEM;
    var abIdx = doc.artboards.getActiveArtboardIndex();
    var actAbBds = doc.artboards[abIdx].artboardRect;
    var obj2move = doc.selection[0];
    obj2move.position = new Array ((actAbBds[2]-actAbBds[0])/2 - obj2move.width/2, (actAbBds[3]-actAbBds[1])  + obj2move.height);
    }




////// Только для Zentiva
function nexusFont() {
    userNames = $.getenv("USERNAME");

    if (userNames != "pavel") {
        var pathToFile = "~/ZENTIVA/Popovic, Pavlo CZ - AW SCRIPTS/Illustrator/nexusfont/Run NexusFont.bat"
        }
    else{
        var pathToFile = "~/AW SCRIPTS/Illustrator/nexusfont/Run NexusFontPavel.bat"
        }
    var fileObj = new File(pathToFile + "/Run NexusFont.bat");
    var parent = fileObj.parent.fsName; // >> /path/to
    //alert(parent);
    var batFile = new File(parent);
    batFile.execute();
};