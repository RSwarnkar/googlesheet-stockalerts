
/*
@author: Rajesh Swarnkar <rjs.swarnkar@gmail.com>
*/

var good_news_count = 0
var bad_news_count = 0

var my_recepient = "youremail@gmail.com"

var my_subject = "STOCK ALERT: "

var my_email_body = ""


// https://stackoverflow.com/questions/55135702/google-apps-script-to-stop-the-trigger-on-specific-days
function TradingTimeCheck(){
  
  var now = new Date();
  Logger.log(now)
  var k = now.getDay() * 100 + now.getHours() + now.getMinutes()/100;
   
  if(  (k>=109 && k<=116) || 
       (k>=209 && k<=216) ||
       (k>=309 && k<=316) ||
       (k>=409 && k<=416) ||
       (k>=509 && k<=516)  ){
    Logger.log("This is Trading Hour: "+k)
    return true
  }
  else{
    Logger.log("This is NOT Trading Hour: "+k)
    return false
  }
}


 

function masterFunction(){
  
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet()
  
  var last_row_num = parseInt(sheet1.getSheetByName("StockList").getLastRow(),10) // Count number of rows 
 
  var data = sheet1.getSheetByName("StockList").getRange("B5:K"+last_row_num).getValues(); // MAGIC NUMBERS
  
  up_stocks = []
  down_stocks = []
  
  //Logger.log(data)
  
  for (row = 0; row<data.length; row++){ // outer for
   // Logger.log("Upper Threshold at row("+(row)+") is "+(data[row][8])) // MAGIC NUMBERS
   // Logger.log("Lower Threshold at row("+(row)+") is "+(data[row][9])) // MAGIC NUMBERS
    
    current_price   = data[row][2] // MAGIC NUMBERS
    upper_threshold = data[row][8] // MAGIC NUMBERS
    lower_threshold = data[row][9] // MAGIC NUMBERS
    stock_symbol = data[row][1] // MAGIC NUMBERS
    
    if(current_price>upper_threshold){
      // Logger.log("1 Good news with stock "+ data[row][1] );
      good_news_count = good_news_count + 1;
      up_stocks.push(stock_symbol)
    }
    
    if(current_price<lower_threshold){
      // Logger.log("1 Bad news with stock "+ data[row][1] );
      bad_news_count = bad_news_count + 1;
      down_stocks.push(stock_symbol)
    }
 
 }  // outer for

  my_subject = "STOCK ALERT: UP(" + good_news_count + "), DOWN(" + bad_news_count + ")";
  my_email_body = "<h3>Stock Alert: There are <bold>" + good_news_count + " UP stocks</bold> and <bold>" + bad_news_count +" DOWN stocks in my portfolio.</bold>. </h3>"
  
  my_email_body = my_email_body + "<br /> <h4 style='color:green'>UP Stocks: </h4>" + up_stocks 
  my_email_body = my_email_body + "<br /> <h4 style='color:maroon'>DOWN Stocks: </h4>" + down_stocks + "<br /> <br />"
  
  if(TradingTimeCheck() && (good_news_count!=0 || bad_news_count!=0) ){
    
    SendReportEmail(my_recepient, my_subject, my_email_body)
    
  }
  
}


