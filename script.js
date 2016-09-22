var booking_no_reg           = /Booking\s*number\:\s*<\/th>\s*<td[^>]*?>\s*([^<]*?)\s*</i;
var date_reg                 = /Date\:\s*<\/th>\s*<td[^>]*?>\s*([^<]*?)\s*</i;
var date_started_reg         = /Date\s*Started\:\s*<\/th>\s*<td[^>]*?>\s*([^<]*?)\s*</i;
var skype_reg                = /Skype\s*Username\:\s*<\/th>\s*<td[^>]*?>\s*([^<]*?)\s*</i;
var email_reg                = /Email\:\s*<\/th>\s*<td[^>]*?>\s*<a[^>]*?>([^<]*?)\s*</i;
var mobile_reg               = /Mobile\:\s*<\/th>\s*<td[^>]*?>\s*<a[^>]*?>([^<]*?)\s*</i;
var name_reg                 = /Customer\:\s*<\/th>\s*<td[^>]*?>\s*([^<]*?)\s*</i;
var country_reg              = /Customer:[\w\W]*?\s*<br>([^<]*?)\s*<\/td>/i;



var siteLeaNameEmail_reg     = /From\s*\:\s*([^<]*?)\s*(?:<|\&lt\;)\s*<\/span>\s*<a\s*href=\"mailto\:([^<]*?)\"/i;
var siteLeadPhone_reg        = /Phone\s*Number\s*\:\s*([^<]*?)</i;
var siteLeadMessageBody_reg  = /Message\s*Body:(?:<[^>]*?>)+([^<]*?)</i

var siteLead_check_keyword_1 = /sent\s*via\s*contact\s*form\s*on\s*Fluent\s*Focu/i;

var GumtreeLeads_check_reg = /The\s*Gumtree\s*Team/;
var trail_booking_check_keyword_1 = /Trial\s*English\s*Lesson/i;



var GumtreeLeads_name_reg = /From:<\/span>([^<]*?)<[^>]*?>\s*(?:<[^>]*?>\s*)+([^<]*?)</i;

function Search() {
 
  var sheet1   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bookeo Trial Lessons');
  var sheet2   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Site Leads');
   var sheet3   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gumtree Leads');
  var row1     = sheet1.getLastRow()+1;
  var row2     = sheet2.getLastRow()+1;
   var row3     = sheet3.getLastRow()+1;
  // Clear existing search results
  //sheet.getRange(row, 1, sheet.getMaxRows() - 1, 5).clearContent();
  
  // Which Gmail Label should be searched?
  var label   = 'label:inbox';
   
  // Retrieve all threads of the specified label
  var threads = GmailApp.search(label);
   
  
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  var lastDate = logSheet.getRange(1, 1);
 
  var count=0;
  var updateDate = '';
  Logger.log(threads.length);
  Logger.log(lastDate.getValue());
 
  for (var i = 0; i < threads.length ; i++) {  
    
    
    var messages = threads[i].getMessages();
    Utilities.sleep(1000);
    if(i==0)
    {      
      updateDate =  messages[messages.length-1].getDate();
    }
    Logger.log('Mail:'+messages[messages.length-1].getSubject());
    Logger.log('Comparing '+lastDate.getValue() + ' and '+messages[messages.length-1].getDate());
    if(lastDate.getValue() < messages[messages.length-1].getDate())
    {
      count++;
      SpreadsheetApp.getActiveSpreadsheet().toast('Reading mail:'+messages[0].getSubject() );   
      Logger.log('---------------NEW MAIL-----------------');
      var messageBody = messages[0].getBody();     
      var result ='';
      if(result = trail_booking_check_keyword_1.exec(messageBody))
      {      
        trailMailParse(sheet1,row1,messages,messageBody);
        row1++;   
      }
      if(result = siteLead_check_keyword_1.exec(messageBody))
      {
        siteLeadMailParse(sheet2,row2,messages,messageBody);
        row2++;   
      }
      if(result = GumtreeLeads_check_reg.exec(messageBody))
      {
        
        GumtreeLeadsParse(sheet3,row3,messages,messageBody);
        row3++;   
      }
    }
    else
    {
      Logger.log('---------------OLD MAIL-----------------');
      break;
    }
    
  }
  if(count==0)
  {
    SpreadsheetApp.getActiveSpreadsheet().toast('new mail(s) not found');
  }
  if(updateDate)
  {
    lastDate.setValue(updateDate);
  }
 

}


function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Mail Parse",
    functionName : "Search"
  }];
  sheet.addMenu("Script Menu", entries);
};


function siteLeadMailParse(sheet,row,messages,messageBody)
{  
  var date = new Date();
  sheet.getRange(row,1).setValue(date);  
  sheet.getRange(row,2).setValue(messages[0].getDate()); 
  sheet.getRange(row,3).setValue(messages[0].getFrom());
  sheet.getRange(row,4).setValue(messages[0].getSubject()); 
  var result='';
  if(result = siteLeaNameEmail_reg.exec(messageBody))
  {
    sheet.getRange(row,6).setValue(result[1]);  
    sheet.getRange(row,7).setValue(result[2]); 
  }  
  
  sheet.getRange(row,9).setValue(date);    
  if(result = siteLeadPhone_reg.exec(messageBody))
  {
    sheet.getRange(row,10).setValue(result[1]);    
  }  
  
  if(result = siteLeadMessageBody_reg.exec(messageBody))
  {
    sheet.getRange(row,11).setValue(result[1]);    
  }  
  
  
  
}  


function GumtreeLeadsParse(sheet,row,messages,messageBody)
{  
  var date = new Date();
  sheet.getRange(row,1).setValue(date);  
  sheet.getRange(row,2).setValue(messages[0].getDate()); 
  sheet.getRange(row,3).setValue(messages[0].getFrom());
  sheet.getRange(row,4).setValue(messages[0].getSubject()); 
  var result='';
  if(result = GumtreeLeads_name_reg.exec(messageBody))
  {
    sheet.getRange(row,6).setValue(result[1]);  
    sheet.getRange(row,9).setValue(result[2]);
  }  
   sheet.getRange(row,7).setValue(date);    
 
  
  
  
}  



function trailMailParse(sheet,row,messages,messageBody)
{  
  var date = new Date();
  sheet.getRange(row,1).setValue(date);  
  sheet.getRange(row,2).setValue(messages[0].getDate()); 
  sheet.getRange(row,3).setValue(messages[0].getFrom());
  sheet.getRange(row,4).setValue(messages[0].getSubject());  
  if(result = booking_no_reg.exec(messageBody))
  {
    sheet.getRange(row,5).setValue(result[1]);  
  }  
  if(result = name_reg.exec(messageBody))
  {
    sheet.getRange(row,6).setValue(result[1]);  
  }
  if(result = country_reg.exec(messageBody))
  {
    sheet.getRange(row,7).setValue(result[1]);  
  }  
  if(result = date_reg.exec(messageBody))
  {
    sheet.getRange(row,8).setValue(result[1]);  
  }
  if(result = date_started_reg.exec(messageBody))
  {
    sheet.getRange(row,9).setValue(result[1]);  
  }
  if(result = skype_reg.exec(messageBody))
  {
    sheet.getRange(row,10).setValue(result[1]);  
  }
  if(result = mobile_reg.exec(messageBody))
  {
    sheet.getRange(row,11).setValue(result[1]);  
  }
  if(result = email_reg.exec(messageBody))
  {
    sheet.getRange(row,12).setValue(result[1]);  
  } 
}
