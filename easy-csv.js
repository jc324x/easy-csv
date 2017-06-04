// config 

var config = importConfiguration("https://raw.githubusercontent.com/jcodesmn/easy-csv/master/apple-school-manager.json");
// var config = importConfiguration("https://raw.githubusercontent.com/jcodesmn/easy-csv/master/broken-apple-school-manager.json");
// var config = importConfiguration("https://raw.githubusercontent.com/jcodesmn/easy-csv/master/jss-mutt.json");

// https://github.com/jcodesmn/google-apps-script-cheat-sheet 

// files and folders

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

// array of objects

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


function colNum(col) {
  var col = col.toUpperCase();
  if (col.length === 1)  {
    var chr0 = col.charCodeAt(0) - 64;
    return chr0;
  } else if (col.length === 2) {
    var chr0 = (col.charCodeAt(0) - 64) * 26;
    var chr1 = col.charCodeAt(1) - 64;
    return chr0 + chr1;
  }
}


// range

function headerVal(rangeObj){
  var vals = rangeObj.getValues();
  var arr  = [];
  for (var i = 0; i < vals[0].length; i++) {
    var val = vals[0][i];
    arr.push(val);
  } 
  return arr;
}

function numCol(num) {
  var num = num - 1, chr;
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

// folders

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

// timestamp

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

// config variables and globals

var ss          = SpreadsheetApp.getActiveSpreadsheet();
var sheets      = arrSheetNames(ss);
var projectPath = config.projectPath;
var process     = config.process;
var keepHeaders = config.keepHeaders;
var targets     = config.targets;
var zipOutput   = config.zipOutput;

var target, targetSheet, targetRange;

// columns

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

// get A1 Notation as numbers?
// target (small) = A:B   = 1,(1),2,(5)
// data range     = A1:C5 = 1,1,3,5
// result         = A1:B5 = 1,1,2,5

// target (big)   = A:J100 = 1,(1),10,100
// data range     = A1:C5  = 1,1,3,5
// result         = A1:C5  = 1,1,3,5

// need to give a number to a column
// then convert into two arrays and do math (?)
// then convert back to A1 notation

function Target(sheetObj, a1Notation) {
  var a1;
  var dr = sheetObj.getDataRange().getA1Notation();

  if (typeof a1Notation === "undefined") {
    a1 = dr;
  } else {
    a1 = a1Notation;
  }

  var lColNum  = sheetObj.getLastColumn();
  var lRow     = sheetObj.getLastRow();
  var split    = a1.split(":");
  var startCol = colNum(split[0].match(/\D/g,'').toString());
  var startRow = parseInt(split[0].match(/\d+/g));
  var endCol   = colNum(split[1].match(/\D/g,'').toString());
  var endRow   = parseInt(split[1].match(/\d+/g));

  if (isNaN(startRow)) {
    startRow = 1;
  }

  if (isNaN(endCol) || (endCol > lColNum)) {
    endCol = lColNum;
  }

  if (isNaN(endRow) || (endRow > lRow)) {
    endRow = lRow;
  }

  this.startCol = startCol;
  this.startRow = startRow;
  this.endCol   = endCol;
  this.endRow   = endRow;

  this.chopHeaders = function() {
    this.startRow += 1;
  };

  this.getA1Notation = function() {
    return numCol(this.startCol) + this.startRow + ":" + numCol(this.endCol) + this.endRow;
  };
} 

function rangeAsCSV(a1Notation) {
  // -> empty sheet -> A1:`0`
}

function runScript() {

  switch(process) {
    case "exportSheets":
      var path = projectPath + " " + fmat12DT(); 
      var folder = createVerifyPath(path);

      for (var i = 0; i < targets.length; i++) {
        var sheet = targets[i].sheet;

        if (checkValIn(sheets, sheet)) {
          sheet = ss.getSheetByName(sheet);
        }

        var name = sheet.getName();
        var target = new Target(sheet, targets[i].range);
        Logger.log(name);
        Logger.log(target.getA1Notation());
        // var csv = rangeAsCSV(target);
        // test that csv is not null...then create file
        // folder.createFile(name, csv);
      } 
      break;
    default:
      Logger.log("please check your configuration and try again");
  }

}

