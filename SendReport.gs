/*
@author: Rajesh Swarnkar <rjs.swarnkar@gmail.com>
*/

/*
This functions sends the actual email with full report in the body. 
*/

 

function format(x, n){
    x = parseFloat(x);
    n = n || 2;
    return parseFloat(x.toFixed(n))
}

function SendReportEmail(recepient, subject, mainMessageBody){
  
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet()
  
 
  var sheet1  = SpreadsheetApp.getActiveSpreadsheet();
  var lastRow = parseInt(sheet1.getSheetByName("StockList").getLastRow(),10)
  
  var data = sheet1.getSheetByName("StockList").getRange("B4:K"+lastRow).getValues(); // MAGIC NUMBERS

  
  var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
  var htmltable = '<br /> <table ' + TABLEFORMAT +' ">';
  
 
  for (row = 0; row<data.length; row++){ // outer for
    
    htmltable += '<tr>';
    
    for (col = 0 ;col<data[row].length; col++){
      if (data[row][col] === "" || 0) {
        htmltable += '<td>' + 'None' + '</td>';
      } 
      else if (row === 0)  {
          htmltable += '<th>' + data[row][col] + '</th>';
        }
      else {

        if(!isNaN(parseFloat(data[row][col]))){
          // data is numeric
          htmltable += '<td>' + format(data[row][col]) + '</td>';
          //Logger.log("Formatted: "+format(data[row][col],2))

        }
        else {  // data is text, so just append
          htmltable += '<td>' + data[row][col] + '</td>';
        }

        
      }
    } // end inner for

     htmltable += '</tr>';
  } // end outer for
  
  htmltable += '</table>';
  htmltable = mainMessageBody + htmltable;
  //Logger.log(data);
  //Logger.log(htmltable);

  MailApp.sendEmail(recepient, subject,'' ,{htmlBody: htmltable})
  Logger.log("Mail Sent.");
  
}
