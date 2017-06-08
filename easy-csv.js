// config 

var config = importConfiguration("https://raw.githubusercontent.com/jcodesmn/easy-csv/master/exportSpreadsheet.json");
// var config = importConfiguration("https://raw.githubusercontent.com/jcodesmn/easy-csv/master/apple-school-manager.json");
// var config = importConfiguration("https://raw.githubusercontent.com/jcodesmn/easy-csv/master/jss-mutt.json");

// global

var ui         = SpreadsheetApp.getUi();
var ss         = SpreadsheetApp.getActiveSpreadsheet();
var sheets     = ss.getSheets();
var sheetNames = arrSheetNames(ss);

// menu

function onOpen(e) {
  ui.createMenu('Easy CSV')
    .addItem('Run Recipe', 'runRecipe')
    .addToUi();
}

// files and folders

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

function zipFilesIn(fldr, name){
  var validName;
  if (typeof name === "undefined") {
    validName = "Archive.zip";
  } else {
    validName = name;
  }
  var blobs = [];
  var files = filesIn(fldr);
  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    var blob = file.getBlob();
    blobs.push(blob);
  } 
  var zips = Utilities.zip(blobs, validName);
  var zip = fldr.createFile(zips);
  return zip;
}

// json

function jsonFromUrl(url) {
  var rsp  = UrlFetchApp.fetch(url);
  var data = rsp.getContentText();
  var json = JSON.parse(data);
  return json;
} 

function jsonFromFile(file) {
  var data = file.getBlob().getDataAsString();
  var json = JSON.parse(data);
  return json;
} 

function importConfiguration(scriptConfig) {
  var regExp = new RegExp("^(http|https)://");
  var test   = regExp.test(scriptConfig);
  var json;
  if (test) {
    json = jsonFromUrl(scriptConfig); 
    return json;
  } else {
    var file = findFileAtPath(scriptConfig); 
    json = jsonFromFile(file); 
    return json;
  }
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

function numCol(number) {
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

function colNum(column) {
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

function headerVal(rangeObj){
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
  var lColABC  = numCol(lColNum);
  var lRow     = sheetObj.getLastRow();
  var hRange   = sheetObj.getRange("A" + hRow + ":" + lColABC + hRow);
  var tColABC  = numCol(colIndex);
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

function Scope(sheetObj, a1Notation) {
  var a1, startCol, startRow, endCol, endRow;
  var datarange = sheetObj.getDataRange().getA1Notation();

  if (typeof a1Notation === "undefined") {
    a1 = datarange;
  } else {
    a1 = a1Notation;
  }

  var lColNum  = sheetObj.getLastColumn();
  var lRow     = sheetObj.getLastRow();
  var split    = a1.split(":");

  if (split.length == 2) {
    startCol = colNum(split[0].match(/\D/g,'').toString());
    startRow = parseInt(split[0].match(/\d+/g));
    endCol = colNum(split[1].match(/\D/g,'').toString());
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

  this.sheet    = sheetObj;
  this.startCol = startCol;
  this.startRow = startRow;
  this.endCol   = endCol;
  this.endRow   = endRow;

  this.chopHeaders = function() {
    this.startRow += 1;
  };

  this.getA1Notation = function() {
    if (this.endCol >= 1 && this.endRow >= 1) {
      var _a1Notation = numCol(this.startCol) + this.startRow + ":" + numCol(this.endCol) + this.endRow;
      return  _a1Notation;
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

function expandScopeAndExportToCSV() {
  
} 


// function convertScopeToCSV(target){
// }

function runRecipe() {
  var folder = createVerifyPath(config.projectPath + " " + fmat12DT());
  var sheet, scope, zip;

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
          var validSheet = ss.getSheetByName(sheet); 
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
      // Scope
      // getColumn(), getNumColumns(), getLastColumn()
      Logger.log("expandSheet");
    break;

    default:
      Logger.log("please check your configuration and try again");
  }
}
