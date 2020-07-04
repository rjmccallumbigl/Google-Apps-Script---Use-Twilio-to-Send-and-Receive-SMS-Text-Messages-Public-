/**
*
* Integrate Google Sheets with Twilio.
*
* https://www.reddit.com/r/googlesheets/comments/hf05cp/hey_i_volunteer_at_a_food_coop_and_want_people_to/
*
*/

// ***********************************************************************************************************************

/* 
* Create a menu option for script functions
*
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
  .addItem('Receive Text Messages', 'receiveSMS')
  .addItem('Send Text Messages', 'sendSMS')
  .addToUi();
}


/*
* Authenticate your Twilio account. You can get a free trial SID and token by creating an account at www.twilio.com
* @return the authenticated Twilio object
*
*/

function authTwilio(){
  
  //  Declare variables
  var ryanMcSloMo = {
    accountSID : "xxx",
    authToken : "xxx"        
  };
  
  ryanMcSloMo.url = "https://api.twilio.com/2010-04-01/Accounts/" + ryanMcSloMo.accountSID + "/Messages.json";
  ryanMcSloMo.options = {
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(ryanMcSloMo.accountSID + ":" + ryanMcSloMo.authToken)
    }
  }
  ryanMcSloMo.fromNumber = "+xxx";
  
  //  Return auth object
  return ryanMcSloMo;
}

/*
*
* Send an SMS
* @param toNumber {String} The number being texted, optional if code below is updated
* @param smsText {String} Your SMS text message, optional if code below is updated
*
* References
* https://www.labnol.org/code/20200-send-sms-google-script
*
*/

function sendSMS(toNumber, smsText) {
  
  var toNumber = toNumber || "+xxx";  
  var smsText = smsText || "Testing!";
  var Ryan = authTwilio();
  
  if (smsText.length > 160) {
    Logger.log("The text should be limited to 160 characters");
    return;
  }
  
  // Update options for API
  Ryan.options.payload = {
    "From" : Ryan.fromNumber,
    "To"   : toNumber,
    "Body" : smsText
  };
  Ryan.options.method = "POST";
  Ryan.options.muteHttpExceptions = true;
  
  //  Send your text using the API
  try {
    var response = JSON.parse(UrlFetchApp.fetch(Ryan.url, Ryan.options).getContentText());
  } catch (e) {
    console.log(e);
  }
  
  console.log(response);
  
  //  Sleep for a sec to not exceed API limits
  Utilities.sleep(1000);  
}

/*
*
* Log all SMS messages to sheet
*
* References
* https://www.twilio.com/docs/chat/rest/message-resource
* https://www.twilio.com/docs/sms/api
*/

function receiveSMS() {  
  
  //  Declare variables
  var Ryan = authTwilio();
  var outerArray = [];
  var innerArray = [];
  var thisValue = "";
  var textSheetName = "Text Messages";
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var sheet = spreadsheet.getSheetByName(textSheetName).clear;
  } catch (e) {
    var sheet = spreadsheet.insertSheet(textSheetName).activate();
  }  
  
  // Update options for API
  Ryan.options.method = "GET";  
  Ryan.options.muteHttpExceptions = true;
  
  //  Receive your texts using the API
  try {
    var response = JSON.parse(UrlFetchApp.fetch(Ryan.url, Ryan.options).getContentText());
  } catch (e) {
    Logger.log(e);
  }
  
  //  Create 2D array with all texts
  innerArray.push("date_updated");
  innerArray.push("from");
  innerArray.push("to");
  innerArray.push("direction");
  innerArray.push("status");
  innerArray.push("body");    
  outerArray.push(innerArray);
  
  for (var key in response.messages) {
    console.log(key);
    innerArray = [];   
    innerArray.push(response.messages[key].date_updated);
    innerArray.push(response.messages[key].from);
    innerArray.push(response.messages[key].to);
    innerArray.push(response.messages[key].direction);
    innerArray.push(response.messages[key].status);
    innerArray.push(response.messages[key].body);    
    outerArray.push(innerArray);
  };
  
  //  Print to sheet
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, 1, outerArray.length, outerArray[0].length).setValues(outerArray);
  
  //  Sleep for a sec to not exceed API limits
  Utilities.sleep(1000);  
}
