////////////////////////////////////////////////////////////////
//                         v0.2-beta                          //
// https://github.com/jcodesmn/google-apps-script-cheat-sheet //
////////////////////////////////////////////////////////////////

//////////////////////////////
// global, onOpen(), config //
//////////////////////////////

var ui    = SpreadsheetApp.getUi();
var uProp = PropertiesService.getUserProperties();

/**
 * @requires global ui / uProp vlues 
 * Creates a menu with that allows the user to set the configuration options and run the script.
 */

function onOpen() {
  ui.createMenu("Easy CSV")
  .addItem("Run Script", "runScript")
  .addSeparator()
  .addSubMenu(ui.createMenu("Configuration")
    .addItem("Set Configuration", "setConfiguration")
    .addItem("Show Configuration", "showConfiguration")
    .addItem("Clear Configuration", "clearConfiguration"))
  .addToUi();
}

/**
 * @requires global ui / uProp values
 * Parses JSON at the value given and sets config equal to the stringified result.
 */

function setConfiguration() {
  var uiPrompt = ui.prompt(
      "Please enter a URL or a path to a file in Drive:",
      ui.ButtonSet.OK_CANCEL);
  var button = uiPrompt.getSelectedButton();
  if (button == ui.Button.OK) {
    var text   = uiPrompt.getResponseText();
    var config = JSON.stringify(objFromUrlOrFile(text));
    uProp.setProperty("config", config);
  } 
}

/**
 * @requires global ui / uProp values
 * Displays an alert of the config value or a displays a message if config is unset.
 */

function showConfiguration() {
  var config = uProp.getProperty("config");
  if (config) {
    ui.alert(config);
  } else {
    ui.alert("No configuration set.");
  }
}

/**
 * @requires global ui / uProp values
 * Clears out all user properties.
 */

function clearConfiguration() {
  uProp.deleteAllProperties();
  ui.alert("All settings cleared.");
}

//////////
// JSON //
//////////

/**
 * Returns an object from a URL.
 *
 * @param {string} url
 * @returns {Object}
 */

function objFromUrl(url) {
  var rsp  = UrlFetchApp.fetch(url);
  var data = rsp.getContentText();
  return JSON.parse(data);
} 

/**
 * Returns an object from a file in Drive.
 *
 * @param {File} file
 * @returns {Object}
 */

function objFromFile(file) {
  var data = file.getBlob().getDataAsString();
  return JSON.parse(data);
} 

/**
 * Returns an object from a URL or from a file in Drive.
 *
 * @param {string || File} input
 * @returns {Object}
 */

function objFromUrlOrFile(input) {
  var regExp = new RegExp("^(http|https)://");
  var test   = regExp.test(input);
  if (test) {
    return objFromUrl(input);
  } else {
    var file = findFileAtPath(input); 
    return objFromFile(file);
  }
}

///////////////////////
// Files and Folders //
///////////////////////

function filesIn(fldr) {
  var fi  = fldr.getFiles();
  var arr = [];
  while (fi.hasNext()) {
    var file = fi.next();
    arr.push(file);
  } 
  return arr;
}

function foldersIn(fldr) {
  var fi  = fldr.getFolders();
  var arr = [];
  while (fi.hasNext()) {
    var _fldr = fi.next();
    arr.push(_fldr);
  } 
  return arr;
}

function createVerifyPath(path) {
  var arr = path.split('/');
  var fldr;
  for (i = 0; i < arr.length; i++) {
    var fi = DriveApp.getRootFolder().getFoldersByName(arr[i]);
    if (i === 0) {
      if(!(fi.hasNext())) {
        DriveApp.createFolder(arr[i]);
        fi = DriveApp.getFoldersByName(arr[i]);
      } 
      fldr = fi.next();
    } else if (i >= 1) {
      fi = fldr.getFoldersByName(arr[i]);
      if(!(fi.hasNext())) {
        fldr.createFolder(arr[i]);
        fi = DriveApp.getFoldersByName(arr[i]);
      } 
      fldr = fi.next();
    }
  } 
  return fldr;
}

function findFileAtPath(path) {
  var arr  = path.split('/');
  var file = arr[arr.length -1];
  var fldr, fi;
  for (i = 0; i < arr.length - 1; i++) {
    if (i === 0) {
      fi = DriveApp.getRootFolder().getFoldersByName(arr[i]);
      if (fi.hasNext()) {
        fldr = fi.next();
      } else { 
        return null;
      }
    } else if (i >= 1) {
        fi = fldr.getFoldersByName(arr[i]);
        if (fi.hasNext()) {
          fldr = fi.next();
        } else { 
          return null;
        }
    }
  } 
  return findFileIn(fldr, file);
} 

// FLAG -> add to main.js

function zipFilesIn(fldrIn, name, fldrOut) {
  var validName, validFldrOut;
  if (typeof name === "undefined") {
    validName = "Archive.zip";
  } else {
    validName = name;
  }
  var blobs = [];
  var files = filesIn(fldrIn);
  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    var blob = file.getBlob();
    blobs.push(blob);
  } 
  var zips = Utilities.zip(blobs, validName);
  if (typeof fldrOut === "undefined") {
    validFldrOut = fldrIn;
  } else {
    validFldrOut = fldrOut;
  }
  var zip = validFldrOut.createFile(zips);
  return zip;
}

// arrays 

function checkValIn(arr, val) { 
  return arr.indexOf(val) > -1; 
}

function arrSheetNames(ssObj) {
  var sheets = ssObj.getSheets();
  var arr    = [];
  for (var i = 0; i < sheets.length; i++) {
    arr.push(sheets[i].getName());
  } 
  return arr;
} 

// columns & ranges

function numToCol(number) {
  var num = number - 1, chr;
  if (num <= 25) {
    chr = String.fromCharCode(97 + num).toUpperCase();
    return chr;
  } else if (num >= 26 && num <= 51) {
    num -= 26;
    chr = String.fromCharCode(97 + num).toUpperCase();
    return "A" + chr;
  } else if (num >= 52 && num <= 77) {
    num -= 52;
    chr = String.fromCharCode(97 + num).toUpperCase();
    return "B" + chr;
  } else if (num >= 78 && num <= 103) {
    num -= 78;
    chr = String.fromCharCode(97 + num).toUpperCase();
    return "C" + chr;
  }
}

function colToNum(column) {
  var col = column.toUpperCase();
  var chr0, chr1;
  if (col.length === 1)  {
    chr0 = col.charCodeAt(0) - 64;
    return chr0;
  } else if (col.length === 2) {
    chr0 = (col.charCodeAt(0) - 64) * 26;
    chr1 = col.charCodeAt(1) - 64;
    return chr0 + chr1;
  }
}

function arrheaderVal(rangeObj){
  var vals = rangeObj.getValues();
  var arr  = [];
  for (var i = 0; i < vals[0].length; i++) {
    var val = vals[0][i];
    arr.push(val);
  } 
  return arr;
}

function arrForColNo(sheetObj, hRow, colIndex){
  var lColNum  = sheetObj.getLastColumn();
  var lColABC  = numToCol(lColNum);
  var lRow     = sheetObj.getLastRow();
  var hRange   = sheetObj.getRange("A" + hRow + ":" + lColABC + hRow);
  var tColABC  = numToCol(colIndex);
  var rangeObj = sheetObj.getRange(tColABC + (hRow +1) + ":" + tColABC + lRow);
  var h        = rangeObj.getHeight();
  var vals     = rangeObj.getValues();
  var arr      = [];
  for (var i = 0; i < h; i++) {
      var val  = vals[i][0];
      arr.push(String(val));
  }  
  return arr;
}

// datestamp

function fmat12DT() {
  var n = new Date();
  var d = [ n.getMonth() + 1, n.getDate(), n.getYear() ];
    var t = [ n.getHours(), n.getMinutes(), n.getSeconds() ];
    var s = ( t[0] < 12 ) ? "AM" : "PM";
  t[0]  = ( t[0] <= 12 ) ? t[0] : t[0] - 12;
  for ( var i = 1; i < 3; i++ ) {
    if ( t[i] < 10 ) {
      t[i] = "0" + t[i];
    }
  }
  return d.join("-") + " " + t.join(":") + " " + s;
}

function Scope(sheet, a1Notation) {
  var a1, startCol, startRow, endCol, endRow;
  var dataRange = sheet.getDataRange().getA1Notation();

  if (typeof a1Notation === "undefined") {
    a1 = dataRange;
  } else {
    a1 = a1Notation;
  }

  var lColNum  = sheet.getLastColumn();
  var lRow     = sheet.getLastRow();
  var split    = a1.split(":");

  if (split.length == 2) {
    startCol = colToNum(split[0].match(/\D/g,'').toString());
    startRow = parseInt(split[0].match(/\d+/g));
    endCol = colToNum(split[1].match(/\D/g,'').toString());
    endRow = parseInt(split[1].match(/\d+/g));
  }

  if (isNaN(startRow)) {
    startRow = 1;
  }

  if (isNaN(endCol) || (endCol > lColNum)) {
    endCol = lColNum;
  }

  if (isNaN(endRow) || (endRow > lRow)) {
    endRow = lRow;
  }

  this.sheet    = sheet;
  this.startCol = startCol;
  this.startRow = startRow;
  this.endCol   = endCol;
  this.endRow   = endRow;

  this.chopHeaders = function() {
    this.startRow += 1;
  };

  this.getA1Notation = function() {
    if (this.endCol >= 1 && this.endRow >= 1) {
      return numToCol(this.startCol) + this.startRow + ":" + numToCol(this.endCol) + this.endRow;
    }
  };
} 

function exportScopeToCSV(scope, folder) {
  var a1 = scope.getA1Notation();
  if (typeof a1 !== "undefined") {
    var csv = "";
    var values = scope.sheet.getRange(scope.getA1Notation()).getValues();
    for (var i = 0; i < (scope.endRow); i++) {
      for (var j = 0; j < (scope.endCol); j++) {
        var value = values[i][j];
        if (j === (scope.endCol - 1) && i === (scope.endRow -1)) {
          csv += value;
        } else if (j === (scope.endCol - 1)) {
          csv += value + "\n";
        } else {
          csv += value + ",";
        }
      } 
    } 
    var fileName = scope.sheet.getName() + ".csv";
    folder.createFile(fileName, csv);
  }
}

// FLAG -> config isn't global

function expandScopeAndExportToCSV(scope, folder) {
  var firstColumn, secondColumn, header, csv;
  var a1 = scope.getA1Notation();
  if (typeof a1 !== "undefined") {
    var range   = scope.sheet.getRange(scope.getA1Notation());
    var headers = arrheaderVal(range);
    var numCols = range.getNumColumns();
    var numRows = range.getLastRow() - 1;
    if (config.removeHeaders) {
      csv = "";
      firstColumn = arrForColNo(scope.sheet, 1, 1);
      if (numCols >= 2) {
        for (var i = 2; i < numCols + 1; i++) {
          header = headers[i - 1];
          secondColumn = arrForColNo(scope.sheet, 1, i);
          for (var j = 0; j < firstColumn.length; j++) {
            csv += firstColumn[j] + ", " + secondColumn[j] + "\n";
          } 
          var fileName = header + ".csv";
          folder.createFile(fileName, csv);
        }
      } 
    } else {
      // need to write alternate
    }
  }
} 

function runScript() {

  var config     = JSON.parse(uProp.getProperty("config"));
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var sheets     = ss.getSheets();
  var sheetNames = arrSheetNames(ss);
  var folder     = createVerifyPath(config.projectPath + " " + fmat12DT());

  var sheet, validSheet, scope, zip;

  switch(config.process) {

    case "exportSpreadsheet":
      for (var i = 0; i < sheets.length; i++) {
        scope = new Scope(sheets[i]);
        exportScopeToCSV(scope, folder);
      } 
      if (config.zipCSVs === true) {
        zip = zipFilesIn(folder, config.zipName);
      }
    break;

    case "exportSheets":
      for (var j = 0; j < config.targets.length; j++) {
        sheet = config.targets[j].sheet;
        if (checkValIn(sheetNames, sheet)) { 
          validSheet = ss.getSheetByName(sheet); 
          scope = new Scope(validSheet, config.targets[j].range);
          if (config.chopHeaders === true) {
            scope.chopHeaders();
          }
          exportScopeToCSV(scope, folder);
        } 
      } 
      if (config.zipCSVs === true) {
        zip = zipFilesIn(folder, config.zipName);
      }
    break;

    case "expandSheet":
      sheet = config.target.sheet;
      if (checkValIn(sheetNames, sheet)) { 
        validSheet = ss.getSheetByName(config.target.sheet );
        scope = new Scope(validSheet, config.target.range);
        expandScopeAndExportToCSV(scope, folder);
      }
      if (config.zipCSVs === true) {
        zip = zipFilesIn(folder, config.zipName);
      }
    break;

    default:
      Logger.log("please check your configuration and try again");
  }
}
