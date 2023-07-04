
function main(){
  var ss = SpreadsheetApp.openById( "1CICwSSaCMzsJCFRmyOYy3Qhjc9xJlwOe8oxvJjSR7aA" );
	var sheet = ss.getSheetByName("Tabellenblatt1");
  var lastRow =  sheet.getLastRow();
  
 sheet.getRange("A1:Z"+lastRow).setNumberFormat("@");
 Logger.log("Delete columns")
  deleteColumns(sheet);
    Logger.log("Adding T to Date");
 addTtoDate(sheet,lastRow);
  Logger.log("Adding TimeZone to Time");
  addZtoTime(sheet,lastRow);
  Logger.log("Concat Times");
  concateDateTime(sheet, lastRow);
  Logger.log("Delete Zeit");
 deleteColumnsZeit(sheet);
 substractRanatt(sheet, lastRow);
  Logger.log("Add Gender");
  createGender(sheet,lastRow);
  Logger.log("Deleting Amazon Costumers");
  
 Logger.log("Iso 3661");
  dict(sheet,lastRow);
    Logger.log("Add Currency");
 createCurrency(sheet,lastRow);
Logger.log("Delete Amazon");
  deleteAmazon(sheet,lastRow);
 Logger.log("Delete 0.00 Umsatz");
  //deleteZeroValue(sheet);
}

function dict(sheet,lastRow){
    let data = sheet.getRange("A1:Z"+lastRow).getValues();

  	 var colname = "Kunde/Land";
    let column = data[0].indexOf(colname);
    for (var i = 1; i < lastRow; i++) {	
if(data[i][column] == "Deutschland" || data[i][column] == "Deutschland".toLocaleUpperCase() || data[i][column] == "Deutschland".toLocaleLowerCase()){
  data[i][column] = "DE";
}else if(data[i][column] == "Österreich" || data[i][column] == "Österreich".toLocaleUpperCase() || data[i][column] == "Österreich".toLocaleLowerCase()){
  data[i][column] = "AT";
  }else if(data[i][column] == "Irland" || data[i][column] == "Irland".toLocaleUpperCase() || data[i][column] == "Irland".toLocaleLowerCase()){
  data[i][column] = "IR";
}else if(data[i][column] == "Italien" || data[i][column] == "Italien".toLocaleUpperCase() || data[i][column] == "Italien".toLocaleLowerCase()){
  data[i][column] = "IT";
}else if(data[i][column] == "Belgien" || data[i][column] == "Belgien".toLocaleUpperCase() || data[i][column] == "Belgien".toLocaleLowerCase()){
  data[i][column] = "BE";
}else if(data[i][column] == "Frankreich" || data[i][column] == "Frankreich".toLocaleUpperCase() || data[i][column] == "Frankreich".toLocaleLowerCase()){
  data[i][column] = "FR";
}else if(data[i][column] == "Dänemark" || data[i][column] == "Dänemark".toLocaleUpperCase() || data[i][column] == "Dänemark".toLocaleLowerCase()){
  data[i][column] = "DK";
}else if(data[i][column] == "Ungarn" || data[i][column] == "Ungarn".toLocaleUpperCase() || data[i][column] == "Ungarn".toLocaleLowerCase()){
  data[i][column] = "HU";
}else if(data[i][column] == "Spanien" || data[i][column] == "Spanien".toLocaleUpperCase() || data[i][column] == "Spanien".toLocaleLowerCase()){
  data[i][column] = "ES";
  }else if(data[i][column] == "Luxemburg" || data[i][column] == "Luxemburg".toLocaleUpperCase() || data[i][column] == "Luxemburg".toLocaleLowerCase()){
  data[i][column] = "LU";
 }else if(data[i][column] == "Tschechien" || data[i][column] == "Tschechien".toLocaleUpperCase() || data[i][column] == "Tschechien".toLocaleLowerCase()){
  data[i][column] = "CH";
}
}
    sheet.getRange("A1:Z"+lastRow).setValues(data);
    
}


function addTtoDate(sheet,lastRow){
  let data = sheet.getRange("A1:Z"+lastRow).getValues();
    var colname = "Bestelldatum/Datum";
  let column = data[0].indexOf(colname);
    for (var i = 1; i < lastRow; i++) {
      var value = data[i][column];
      var array = value.split(".");
      var newdate = array[2]+"-"+array[1]+"-"+array[0];
      var tValue = newdate.toString() + "T";
      data[i][column] = tValue;
  //let column2 = data[1][column];
  //Logger.log(data[i][column]);
  //Logger.log(i);
    }

    sheet.getRange("A1:Z"+lastRow).setValues(data);
}

function addZtoTime(sheet,lastRow){
  let data = sheet.getRange("A1:Z"+lastRow).getValues();
  var colname = "Bestelldatum/Zeit";
  let column = data[0].indexOf(colname);
  for (var i = 1; i < lastRow; i++) {
    var value = data[i][column];
    var tValue = value.toString() + "+01:00";
    data[i][column] = tValue;
//let column2 = data[1][column];
//Logger.log(data[i][column]);
//Logger.log(i);
  }

    sheet.getRange("A1:Z"+lastRow).setValues(data);
}

function concateDateTime(sheet, lastRow){
  let data = sheet.getRange("A1:Z"+lastRow).getValues();
  var datum = "Bestelldatum/Datum";
  var zeit = "Bestelldatum/Zeit";
  let columnDate = data[0].indexOf(datum);
  let columnTime = data[0].indexOf(zeit);
  for (var i = 1; i < lastRow; i++) {
    var concatValue = data[i][columnDate].toString() + data[i][columnTime].toString();
    data[i][columnDate] = concatValue;
    }

    sheet.getRange("A1:Z"+lastRow).setValues(data);
}

function substractRanatt(sheet, lastRow){
  let data = sheet.getRange("A1:Z"+lastRow).getValues();
  var rabatt = "Artikelliste/GesamtRabatt";
  var brutto = "Artikelliste/GesamtBrutto";
  let columnrabatt = data[0].indexOf(rabatt);
  let columnbrutto = data[0].indexOf(brutto);
  for (var i = 1; i < lastRow; i++) {
    var concatValue = data[i][columnbrutto] - data[i][columnrabatt];
    data[i][columnbrutto] = concatValue;
    }

    sheet.getRange("A1:Z"+lastRow).setValues(data);
}

function createGender(sheet,lastRow){
  let data = sheet.getRange("A1:Z"+lastRow).getValues();
  var colname = "Kunde/Anrede";
  let column = data[0].indexOf(colname);
  for (var i = 1; i < lastRow; i++) {
    var value = data[i][column];
    if(value == "Herr" || value == "Mister" ){
      var tValue = "M";
    } else if(value == "Frau" || value == "Miss" || value == "Misses" ){
      var tValue = "F";
    }else if(value == ""){
      var tValue = "";
    }
    data[i][column] = tValue;
//let column2 = data[1][column];
//Logger.log(data[i][column]);
//Logger.log(i);
  }

    sheet.getRange("A1:Z"+lastRow).setValues(data);
}
function createCurrency(sheet,lastRow){
  var lastColumn = sheet.getLastColumn();
  let data = sheet.getRange("A1:Z"+lastRow).getValues();
  data[0][lastColumn+1] = "currency";
  for (var i = 1; i < lastRow; i++) {
    data[i][lastColumn+1] = "EUR";
  }
  sheet.getRange("A1:Z"+lastRow).setValues(data);
}


function deleteAmazon(sheet,lastRow){
   for (var i = lastRow; i > 0; i--) {
    var range = sheet.getRange(i,9); 
    var data = range.getValue();
    if (data.includes("marketplace.amazon")) {
      sheet.deleteRow(i);
    }
  }

}

function deleteZeroValue(sheet){
    var lastRow =  sheet.getLastRow();

   for (var i = lastRow; i > 0; i--) {
    var range = sheet.getRange(i,11); 
    var data = range.getValue();
    if (data.includes("0.00") || data.includes("0")) {
      sheet.deleteRow(i);
    }
  }

}
/*
function addZtoTime(sheet,lastRow){
    var sommerzeit = Date.parse("28.03.2021");
    var winterzeit = Date.parse("31.10.2021");
    var zeit = "Bestelldatum/Zeit";
    var colname = "Bestelldatum/Datum";
    let data = sheet.getRange("A1:Z"+lastRow).getValues();
    let column = data[0].indexOf(colname);
    for (var i = 1; i < lastRow; i++) {
      var date = new Date(data[i][column]);
    if(date < sommerzeit && date > winterzeit){
       var tValue = value.toString() + "+02:00";
       data[i][column] = tValue;
       }
    else if(date > sommerzeit && date < winterzeit){
       var tValue = value.toString() + "+01:00";
       data[i][column] = tValue;
    }

    }
      sheet.getRange("A1:Z"+lastRow).setValues(data);
}
*/
function deleteColumns(sheet) {
  var required = ["Bestellnummer", "Bestelldatum/Datum", "Bestelldatum/Zeit", "Kunde/Anrede", "Kunde/Vorname", "Kunde/Name", "Kunde/PLZ", "Kunde/Ort", "Kunde/Land", "Kunde/Email", "Artikelliste/GesamtRabatt", "Artikelliste/GesamtBrutto"]; 
  //",Artikelliste/Artikel/0/Artikelnummer", "Artikelliste/Artikel/1/Artikelnummer", "Artikelliste/Artikel/2/Artikelnummer",  "Artikelliste/Artikel/3/Artikelnummer", "Artikelliste/Artikel/4/Artikelnummer", "Artikelliste/Artikel/5/Artikelnummer", "Artikelliste/Artikel/6/Artikelnummer", "Artikelliste/Artikel/7/Artikelnummer", "Artikelliste/Artikel/8/Artikelnummer", "Artikelliste/Artikel/9/Artikelnummer", "Artikelliste/Artikel/10/Artikelnummer", "Artikelliste/Artikel/11/Artikelnummer" 

 
  var width = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  for (var i = headers.length - 1; i >= 0; i--) {
    if (required.indexOf(headers[i]) == -1) {
      sheet.deleteColumn(i+1);
    }
  }
}
function deleteColumnsZeit(sheet) {
   var required = ["Bestellnummer", "Bestelldatum/Datum", "Kunde/Anrede", "Kunde/Vorname", "Kunde/Name", "Kunde/PLZ", "Kunde/Ort", "Kunde/Land", "Kunde/Email", "Artikelliste/GesamtRabatt", "Artikelliste/GesamtBrutto"]; 

 
  var width = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  for (var i = headers.length - 1; i >= 0; i--) {
    if (required.indexOf(headers[i]) == -1) {
      sheet.deleteColumn(i+1);
    }
  }
}
function event_Time(sheet){
  var sommerzeit = new Date("31.10.2021");


}

function appendString(sheet) {
  const range = sheet.getActiveRange();
  const values = range.getValues();
  const modified = values.map(row => row.map(currentValue => currentValue + " string"));
  range.setValues(modified);
}

function getColumnRangeByName(sheet, columnName) {
  let data = sheet.getRange("A1:1").getValues();
  let column = data[0].indexOf(columnName);
  if (column != -1) {
    return sheet.getRange(2, column + 1, sheet.getMaxRows());
  }
}

function getByName( sheet) {
  var dataRange = sheet.getDataRange().getValues();
  var colData = [];

  for (var i = 1; i < dataRange.length; i++) {
    colData.push(dataRange[i][0]);
  }

  for (var i = 0; i < colData.length; i++) {

    // Take every cell except the first row on col Q (11), as that is the header
    var comments_cell = sheet.getDataRange().getCell(i + 2, 11).getValue();

    // Check for string "delete" inside cell
    if (comments_cell.toString().indexOf("delete") !== -1 || comments_cell.toString().indexOf("Delete") !== -1) {

      // Check for string "removed" not already inside cell
      if (!(comments_cell.toString().indexOf("removed") !== -1)) {

        // Append ", removed"
          sheet.getDataRange().getCell(i + 2, 11).setValue(comments_cell + ", removed");
      }
    }
  }
}