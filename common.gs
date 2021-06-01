function getActiveSheetData (GID) { 
  var ss = SpreadsheetApp.getActive();   
  var allsheets = SpreadsheetApp.getActive().getSheets();
  for(var sheetnumber in allsheets){  
    Logger.log(allsheets[sheetnumber].getSheetId());
    if(allsheets[sheetnumber].getSheetId()==GID){break}
  } 
  var sheet = ss.getSheets()[sheetnumber];
  return sheet;
}



function onEdit() { 
  var ss = SpreadsheetApp.getActiveSheet();
  var cRow = ss.getActiveCell().getRow();
  var cCol = ss.getActiveCell().getColumn();
  var cellValue = ss.getActiveCell().getValue();
  if(ss.getSheetId() == 558550610 ){
    
  }
}



function insertRow(sheet, rowData, optIndex) { 
  var index = optIndex || 1;
  sheet.insertRowBefore(index).getRange(index, 1, 1, rowData.length).setValues([rowData]); 
} 


function send_new_lead_email(){
  MailApp.sendEmail({
      	to:"mail id goes here",
      	name:"name goes here",
      	subject:"subject line goes here",
      	replyTo:"nonreply@testmail.com",
      	bcc:"bcc goes here",
      	cc:"cc goes here",
 		htmlBody:"<p>html content goes here</p>"   
    });
  
  Utilities.sleep(2000);
  SpreadsheetApp.flush();

}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grade = [];
  for(var i=1 ; i<=12 ; i++){
    grade.push({name:'Grade '+i,functionName:"grade_"+i});
    grade.push(null);  
  }
  ss.addMenu("Grade ",grade);
}


const saveGmailtoGoogleDrive = () => {
  const folderId = '1qjxPzFcqiLSXh0JuUeOdf6mH5dLXhehJ';
  const searchQuery = 'has:Attachments';
  const threads = GmailApp.search(searchQuery);
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      const attachments = message.getAttachments({
          includeInlineImages: false,
          includeAttachments: true
      });
      attachments.forEach(attachment => {
        Drive.Files.insert(
          {
            title: attachment.getName(),
            mimeType: attachment.getContentType(),
            parents: [{ id: folderId }]
          },
          attachment.copyBlob()
        );
      });
    });
  });
};



function getFormResponceData(e){
  try{
    //Logger.log("%s", JSON.stringify(e)); Logger.log(e.values);
    var totalAmount = parseFloat(Math.round((((e.values[2].split("||")[2]).split(":")[1]).trim())*e.values[3])).toFixed(2);
    var contentArray = [
          ["Details","Email","Items","Quantity","Total","Remarks"],
          ["Data",e.values[1],e.values[2],e.values[3],totalAmount,e.values[4]] 
        ];
    var ColWidthArray = [200,350];
    var Emailbody = MakeHTMLTable(contentArray,ColWidthArray);
    MailApp.sendEmail({
		    to: "kri#####@gmail.com",
		    subject: "Response: Stock Order Form",
        name:"KIHFHMAhkhiFGDGHuouHYGihE",
		    htmlBody: Emailbody
		});
    //jinling@kimage.com.sg

  }catch(e){
    Logger.log(e.message);
  }
}

function fetch_sheet_tab_data(GID) { 
  var ss = SpreadsheetApp.getActive();   
  var allsheets = SpreadsheetApp.getActive().getSheets();
  for(var sheetnumber in allsheets){  
    if(allsheets[sheetnumber].getSheetId()==GID){
    	break;
    }
  } 
  var sheet = ss.getSheets()[sheetnumber];
  return sheet;
}

function isDate(myDate) {
  return myDate.constructor.toString().indexOf("Date") > -1;
} 

function validateEmail(sEmail){
  if(typeof sEmail != "undefined"){
    var reEmail = /^(?:[\w\!\#\$\%\&\'\*\+\-\/\=\?\^\`\{\|\}\~]+\.)*[\w\!\#\$\%\&\'\*\+\-\/\=\?\^\`\{\|\}\~]+@(?:(?:(?:[a-zA-Z0-9](?:[a-zA-Z0-9\-](?!\.)){0,61}[a-zA-Z0-9]?\.)+[a-zA-Z0-9](?:[a-zA-Z0-9\-](?!$)){0,61}[a-zA-Z0-9]?)|(?:\[(?:(?:[01]?\d{1,2}|2[0-4]\d|25[0-5])\.){3}(?:[01]?\d{1,2}|2[0-4]\d|25[0-5])\]))$/;
    if(!sEmail.match(reEmail)) {
      return false;
    }
    return true;
  }else{
    return false;
  }
}//---End of validateEmail()---//


function MakeHTMLTable(TwoDArray, ColWidthArray) {

  var endofCell= "</td>";  
  //Logger.log(ColWidthArray.length);
  var FirstColWidth = ColWidthArray[0];
  var SecondColWidth = ColWidthArray[1];
  
  var ColWidthOpener = "<col width=\"";  
  var ColWidthCloser = "\">";
  
  var TableHeader = "<!DOCTYPE html><html><head><meta charset=\"UTF-8\"><title>Table created using Google Spreadsheet</title></head><body><!DOCTYPE html><html><head><meta charset=\"UTF-8\"><title>HTML Table for Email on the fly</title></head><body><table cellspacing=\"0\" cellpadding=\"0\" dir=\"ltr\" border=\"1\" style=\"table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;width:0px;border-collapse:collapse;border:none\"><colgroup>";
  //Logger.log(TableHeader);
  
  var DefiningColumns = "";
  for(var k = 0; k < ColWidthArray.length; k++) {
    //Logger.log(ColWidthArray[k]);
    DefiningColumns += ColWidthOpener + ColWidthArray[k] + ColWidthCloser; }
  //Logger.log(DefiningColumns);
  
  TableHeader += DefiningColumns;
  //Logger.log(TableHeader);
  
  //TableHeader += FirstColWidth + "\"><col width=\"" + SecondColWidth + "\"></colgroup><tbody><tr style=\"height:21px\"><td style=\"overflow:hidden;padding:2px 3px;vertical-align:bottom;background-color:rgb(139,195,74);font-weight:bold\">";
  TableHeader += "</colgroup><tbody><tr style=\"height:21px\">";
  //Logger.log(TableHeader);
  
  var HeaderRowStyle = "<td style=\"overflow:hidden;padding:2px 3px;vertical-align:bottom;background-color:rgb(139,195,74);font-weight:bold\">";
  var HeaderCols = "";
  for(var m = 0; m < TwoDArray.length; m++) {
    //   Logger.log(TwoDArray[m][0]);
    HeaderCols += HeaderRowStyle + TwoDArray[m][0] + endofCell; }
  //Logger.log(HeaderCols);
  TableHeader += HeaderCols;
  //Logger.log(TableHeader);
  
  var endofRow= "</tr>";
  var startofRow= "<tr style=\"height:21px\">";
  var whiteRowCol= "<td style=\"overflow:hidden;padding:2px 3px;vertical-align:bottom\">";
  var greenRowCol= "<td style=\"overflow:hidden;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227)\">";
  var TableEnder = "</tbody></table></body></html></body></html'>";
  
  var FullHTMLTable = TableHeader;
  
  for(var i = 1; i < TwoDArray[0].length; i++) {  
    // Logger.log("i is: "+i);
    FullHTMLTable += startofRow;
    for (var j = 0; j < TwoDArray.length; j++) {
      //Logger.log("j is: "+j);        
      if (i % 2) {FullHTMLTable += whiteRowCol + TwoDArray[j][i] + endofCell;}
      else {FullHTMLTable += greenRowCol + TwoDArray[j][i] + endofCell;}
    }
    FullHTMLTable += endofRow;
  }
  FullHTMLTable += TableEnder;
  
  return FullHTMLTable; 
}

function isString (value) {
  return typeof value === 'string' || value instanceof String;
}

function isNumber (value) {
  return typeof value === 'number' && isFinite(value);
}

function isArray (value) {
  return value && typeof value === 'object' && value.constructor === Array;
}

function isFunction (value) {
  return typeof value === 'function';
}

function isObject (value) {
  return value && typeof value === 'object' && value.constructor === Object;
}

function isNull (value) {
  return value === null;
}

function isUndefined (value) {
  return typeof value === 'undefined';
}

function isBoolean (value) {
  return typeof value === 'boolean';
}

function isSymbol (value) {
  return typeof value === 'symbol';
}

function isDate (value) {
  return value instanceof Date;
}

function isError (value) {
  return value instanceof Error && typeof value.message !== 'undefined';
}

function isRegExp (value) {
  return value && typeof value === 'object' && value.constructor === RegExp;
}

function transpose(a) {
  	return a[0].map(function (_, c) { 
  	return a.map(function (r) { 
  		return r[c]; 
  		});
  	});
}


function testingData(){
  var txt = "Category : GK || Item : Colour Sealing (650ml) - Yellow Bottle || Price : 310";
  var totalAmount = parseFloat(Math.round((txt.split("||")[2]).split(":")[1]*5)).toFixed(2);
  Logger.log(totalAmount);
}
function nDimHTMLTableColor(nDArray, ColWidthArray,colors) {
  var endofCell= "</td>";  
  //Logger.log(ColWidthArray.length);
  var FirstColWidth = ColWidthArray[0];
  var SecondColWidth = ColWidthArray[1];
  if (ColWidthArray.length == 2) {
    for (var z = 1; z < nDArray.length; z++) {
      ColWidthArray.push(SecondColWidth); 
    }
  }
  var ColWidthOpener = "<col width=\"";  
  var ColWidthCloser = "\">";
  var TableHeader = "<!DOCTYPE html><html><head><meta charset=\"UTF-8\"><title>Table created using Google Spreadsheet</title></head><body><!DOCTYPE html><html><head><meta charset=\"UTF-8\"><title>HTML Table for Email on the fly</title></head><body><table cellspacing=\"0\" cellpadding=\"0\" dir=\"ltr\" border=\"1\" style=\"table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;width:0px;border-collapse:collapse;border:none\"><colgroup>";
  //Logger.log(TableHeader);
  var DefiningColumns = "";
  for(var k = 0; k < ColWidthArray.length; k++) {
    //Logger.log(ColWidthArray[k]);
    DefiningColumns += ColWidthOpener + ColWidthArray[k] + ColWidthCloser; }
  //Logger.log(DefiningColumns);
  TableHeader += DefiningColumns;
  //Logger.log(TableHeader);
  //TableHeader += FirstColWidth + "\"><col width=\"" + SecondColWidth + "\"></colgroup><tbody><tr style=\"height:21px\"><td style=\"overflow:hidden;padding:2px 3px;vertical-align:bottom;background-color:rgb(139,195,74);font-weight:bold\">";
  TableHeader += "</colgroup><tbody><tr style=\"height:21px\">";
  //Logger.log(TableHeader);
  var HeaderRowStyle = "<td style=\"overflow:hidden;padding:2px 3px;text-align:left;vertical-align:bottom;background-color:" + colors[0] + ";font-weight:bold\">";
  var HeaderCols = "";
  for(var m = 0; m < nDArray.length; m++) {
    //   Logger.log(nDArray[m][0]);
    HeaderCols += HeaderRowStyle + nDArray[m][0] + endofCell; }
  //Logger.log(HeaderCols);
  TableHeader += HeaderCols;
  //Logger.log(TableHeader);
  var endofRow= "</tr>";
  var startofRow= "<tr style=\"height:21px\">";
  var whiteRowCol= "<td style=\"overflow:hidden;padding:2px 3px;vertical-align:bottom\">";
  var greenRowCol= "<td style=\"overflow:hidden;padding:2px 3px;vertical-align:bottom;background-color:" + colors[1] + "\">";
  var TableEnder = "</tbody></table></body></html></body></html'>";
  var FullHTMLTable = TableHeader;
  for(var i = 1; i < nDArray[0].length; i++) {  
    // Logger.log("i is: "+i);
    FullHTMLTable += startofRow;
    for (var j = 0; j < nDArray.length; j++) {
      //Logger.log("j is: "+j);        
      if (i % 2) {FullHTMLTable += whiteRowCol + nDArray[j][i] + endofCell;}
      else {FullHTMLTable += greenRowCol + nDArray[j][i] + endofCell;}
    }
    FullHTMLTable += endofRow;
  }
  FullHTMLTable += TableEnder;
  //Logger.log("FullHTMLTable")
  return FullHTMLTable; 
}
