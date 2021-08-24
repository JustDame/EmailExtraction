// To get messages from emails
function getRelevantMessages()
{
  var threads = GmailApp.search("subject: Nophin",0,1000);
  var messages=[];
  threads.forEach(function(thread)
                  { 
                    let countMessage = thread.getMessageCount();
                    let threadMessages = thread.getMessages();
                    for(let i=0;i<countMessage;i++){
                      messages.push(threadMessages[i]);
                          
                    }
                  });
                  
  
  return messages;
}

// To parse data from message
function parseMessageData(messages)
{
  var records=[];
  for(var m=0;m<messages.length;m++)
  {
    var subject = messages[m].getSubject();
    Logger.log("This is the subject " + subject);
    var text = messages[m].getPlainBody();
    var messageId = messages[m].getThread().getId();
    var fromAddress =  messages[m].getFrom();
    Logger.log("To get the email address  "+ fromAddress);
    Logger.log("Get message ID "+ messageId)
    var myQuantity = text.match(/Quantity:[(0-9)]+/g);
    var myProduct = text.match(/Product:[A-z0-9 ]+/g);
    var myDestination = text.match(/Destination:[A-z0-9 ]+/g);
    
    Logger.log("Info from emails " + myQuantity +" "+ myProduct +" "+ myDestination);
  

    if(!myQuantity ||!myProduct || !myDestination)
    {
      console.log("no matches") 
      //No matches; couldn't parse continue with the next message
      continue;
    }
    
    
    for(let i=0;i<myQuantity.length;i++){
      console.log("Checking for Quantity " + myQuantity.length) 
      let rec = {}    
      rec.Quantity = myQuantity[i].substring(9);
      rec.Product = myProduct[i].substring(8);
      rec.Destination = myDestination[i].substring(12);
      rec.OrderId = messageId + i;
      rec.CustomerEmail = fromAddress;
      //cleanup data 
      records.push(rec);
    }
  }
  
  Logger.log("How many records found "+ records.length);
  Logger.log(records);
  return records;

}

/* 1)Sets up display for the parsed data
 * 2)Creates a template from the parsed html file 
 * 3)Data from relevant message is saved to templ.records
 * 4)An evaluated templ is returned 
*/
function getParsedDataDisplay()
{
  
  var templ = HtmlService.createTemplateFromFile('parsed');
  templ.records = parseMessageData(getRelevantMessages());
  saveDataToSheet(templ.records);
  removeDuplicates();
  
  //NowRemoveDuplicates();
  //templ.records();
  return templ.evaluate();
}

// calls getParsedDataDisplay and starts web app
function doGet()
{ 
  
  return getParsedDataDisplay();
  
}

/* 
 * 1)Opens the spreadsheet with an Url
 * 3)Looks for the current sheet by name
 * 4)For each:Adding each record to row
 *  
*/
function saveDataToSheet(records)
{
   
  //START - command to clear the data in the Spreadsheet
  clearCells();
  //END - command to clear the data in the Spreadsheet 
  
  let spreadsheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1WG-gZMZuRni403_fskzHYnqW8qm0AxqeUwfnRyOw3Cc/edit#gid=0');
  let sheet = spreadsheet.getSheetByName("Sheet1");
  for(let r=1;r<records.length;r++){
    sheet.appendRow([records[r].OrderId,records[r].Quantity,records[r].Product, records[r].Destination,records[r].CustomerEmail ]);
  }
  
  
}

/* 1) Places the spreadsheet in a variable 
 * 2)Finds the last row on the sheet
 * 3)Finds the last column on the sheet 
 * 4)Builds a string with a range to be cleared
 * 5)Clears content of the sheet within the range
 */     
function clearCells(){
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1WG-gZMZuRni403_fskzHYnqW8qm0AxqeUwfnRyOw3Cc/edit#gid=0');
  var lastRow = sheet.getLastRow();
  Logger.log("what is inside lastRow "+ lastRow);
  var lastColumn = sheet.getLastColumn();
  Logger.log("What is inside lastColumn "+ lastColumn);
  var rangeForClear = "A2:E" + lastRow;
  //Logger.log("Today is Friday " + rangeForClear);
  //If statement is checking to prevent clearing of header
  if(lastRow >= 2){
    sheet.getRange(rangeForClear).clearContent();   
  }
}

/* Remove duplicates rows from spreadsheet
 * 1)Gets the active sheet and places it into a variable
 * 2)Gets all data from all the rows 
 * 3)Creates an empty array
 * 4)Looping through the rows 
 * 5)For each row: Adds to new array if it is not already there.
 * 6)Clears everything from A2:E to the last row.
 * 7)Puts new data without duplicates starting at A2.
*/ 

function NowRemoveDuplicates(){
  let sheet   = SpreadsheetApp.getActiveSheet();
  let data    = sheet.getDataRange().getValues();
  let newData = [];
  
  
  for (let i in data)
  {
    
    Logger.log("This is the value for i " + i); 
    
    let row  = data[i];
    Logger.log("This is the value for row " + row);
    
    let duplicate=false;
    Logger.log("This is the value for duplicate " + duplicate);
    
    
    for (let j in newData){
      Logger.log("This is the variable for j " + j);  
      Logger.log("Value of row[i] " + row[i]);
      
      if(row[0] == newData[j][0]){
        duplicate = true;
        }
      
    }
    
    if (!duplicate){
      newData.push(row);
    }
    
  }
  //ClearCells();
  //sheet.getRange(1,1,newData.length, newData[1].length).setValues(newData);
  
  
}

  

// Processes the emails and looks for transactions
function processTransactionEmails()
{
  let messages = getRelevantMessages();
  let records = parseMessageData(messages);
  saveDataToSheet(records);
}

// write multiple rows to the spreadsheet
function writeMultipleRows() {
 let data = getOrdersData();
 let lastRow = SpreadsheetApp.getActiveSheet().getLastRow();
 SpreadsheetApp.getActiveSheet().getRange(lastRow + 1,1,data.length, data[0].length).setValues(data);
}

// if random rows are needed
function getMultipleRandomRows() {
 let OrderId = 1800;
 let Quantity = 70;
 let Product = "Aspirin";
 let Destination = "67 New York Ave";

 var data = [];
 for(var i =0; i < 1000; i++) {
   data.push([rec.OrderId, rec.Quantity, Product, Destination]);
 }
 console.log(data.length);
 return data;
}
/*
 *1) Gets the active spreadsheet
 *2) Takes all the data and puts into the variable
 *3) Assigns and empty array to the variable newData
 *4) Loops through the data. For each row,


*/
function removeDuplicates() {
  Logger.log("Test on Friday");
  let sheet   = SpreadsheetApp.getActiveSheet();
  let data    = sheet.getDataRange().getValues();
  Logger.log("What is inside data " + data);
  Logger.log("How much email is inside data " + data.length);
  var newData = [];
  
  for (let i in data){
    var row       = data[i];
    var duplicate = false;
    for (let j in newData){
      Logger.log("What is inside row.join "+ row[0]);
      Logger.log("What is inside newData[j] "+ newData[j][0]);
       if(row[0] == newData[j][0]){
         duplicate = true;
       }
    }
    if (!duplicate){
      //Logger.log(sheet);
      newData.push(row);
      Logger.log(" What is inside newData " + newData)
    }
  }
  //Logger.log(sheet);
  sheet.clearContents();
  sheet.getRange(1,1, newData.length, newData[0].length).setValues(newData);
  
}


