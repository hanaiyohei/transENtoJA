spreadsheet = null;
sheetMap = [];

function translate(){

  var sheetFrom = _getSheetByName('AEB220327');
  var sheetTo = _getSheetByName('AEB220327');
  
  for(var n=2;n<=900;n++){
    var rangeFrom = sheetFrom.getRange(n,7);
    var rangeTo = sheetTo.getRange(n,12);
  
    var sheetTable = _getSheetByName("タイトル");
    // 変換表を全て取得
    var tableValues = sheetTable.getRange(1,5,sheetTable.getLastRow(),6).getValues();

    var fromStr = rangeFrom.getValue();

    for (var i in tableValues) {
      fromStr = fromStr.split(tableValues[i][0]).join(tableValues[i][1]);
    }
      fromStr = fromStr.replace(/[a-z]/g, "")
      fromStr = fromStr.replace(/[A-Z]/g, "")
      fromStr = fromStr.replace('xxxx','xxxx')
      fromStr = fromStr.replace("     ", " ")
      fromStr = fromStr.replace("    ", " ")
      fromStr = fromStr.replace("   ", " ")
      fromStr = fromStr.replace("  ", " ")
      fromStr = fromStr.replace("\"", "")
      fromStr = fromStr.replace("\'", "")
      fromStr = fromStr.replace("\:", "-")
      if (fromStr.slice(0,1) === " "){
        fromStr = fromStr.slice(1);
      }
    else{
      fromStr = fromStr
    }
    var a = fromStr.split(' ');
    var b = a.filter(function (x, i, self) {
            return self.indexOf(x) === i;
        });
    fromStr = b.join(' ');
    rangeTo.setValue(fromStr);
  }

function _getSpreadSheet(){
  if ( ! spreadsheet) {
    spreadsheet = SpreadsheetApp.getActive();
  }

  return spreadsheet;
}

function _getSheetByName(name){
  if ( ! sheetMap[name]) {
    sheetMap[name] = _getSpreadSheet().getSheetByName(name);
  }

  return sheetMap[name];
  
}}

function spec_translate(){

  var sheetFrom = _getSheetByName('name');
  var sheetTo = _getSheetByName('name');
  
  for(var g=2;g<=3;g++){
    var rangeFrom = sheetFrom.getRange(g,1,350,1);
    var rangeTo = sheetTo.getRange(g,3, 350, 1);
    var sheetTable = _getSheetByName("変換表");
    // 変換表を全て取得
    var tableValues = sheetTable.getRange(2,1,sheetTable.getLastRow(),2).getValues();

    var fromStr = rangeFrom.getValues();
    fromStr = fromStr.join(',')
    Logger.log(fromStr);
    for (var h in tableValues) {
      fromStr = fromStr.split(tableValues[h][0]).join(tableValues[h][1]);
    }
    Logger.log(fromStr);
    //  if (fromStr.slice(0,1) === " "){
    //    fromStr = fromStr.slice(1);
    //  }
    //else{
    //  fromStr = fromStr
    //}
    
    //var a = fromStr.split(' ');
    //var b = a.filter(function (x, i, self) {
    //        return self.indexOf(x) === i;
    //    });
    //fromStr = b.join(' ');
    
    //fromStr = fromStr.split(',');
    rangeTo.setValue(fromStr);
  }

function _getSpreadSheet(){
  if ( ! spreadsheet) {
    spreadsheet = SpreadsheetApp.getActive();
  }

  return spreadsheet;
}

function _getSheetByName(name){
  if ( ! sheetMap[name]) {
    sheetMap[name] = _getSpreadSheet().getSheetByName(name);
  }

  return sheetMap[name];
  
}
}


function ver_con(){

  var sheetFrom = _getSheetByName('IS');
  var sheetTo = _getSheetByName('IS');
  
  for(var n=2;n<=358;n++){
    var rangeFrom = sheetFrom.getRange(n,2);
    var rangeTo = sheetTo.getRange(n,3);
  
    var sheetTable = _getSheetByName("変換表2");
    // 変換表を全て取得
    var tableValues = sheetTable.getRange('B1:C253').getValues();

    var fromStr = rangeFrom.getValue();

    for (var i in tableValues) {
      fromStr = fromStr.split(tableValues[i][0]).join(tableValues[i][1]);
    }

    rangeTo.setValue(fromStr);
  }

function _getSpreadSheet(){
  if ( ! spreadsheet) {
    spreadsheet = SpreadsheetApp.getActive();
  }

  return spreadsheet;
}

function _getSheetByName(name){
  if ( ! sheetMap[name]) {
    sheetMap[name] = _getSpreadSheet().getSheetByName(name);
  }

  return sheetMap[name];
  
}}
