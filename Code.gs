// To get messages from emails
function getRelevantMessages()
{
  var threads = GmailApp.search("subject: Nophin",0,500);
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
    var text = messages[m].getPlainBody();
    var messageId = messages[m].getThread().getId();
    var fromAddress =  messages[m].getFrom();
    var myQuantity = text.match(/Quantity:[(0-9)]+/g);
    var myProduct = text.match(/Product:[A-z0-9 ]+/g);
    var myDestination = text.match(/Destination:[A-z0-9 ]+/g);
    
    if(!myQuantity ||!myProduct || !myDestination)
    {
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
  for(let r=0;r<records.length;r++){
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
  var lastColumn = sheet.getLastColumn();
  var rangeForClear = "A2:E" + lastRow;
  //If statement is checking to prevent clearing of header
  if(lastRow >= 2){
    sheet.getRange(rangeForClear).clearContent();   
  }
}

function removeDuplicates() {
  let sheet   = SpreadsheetApp.getActiveSheet();
  let data    = sheet.getDataRange().getValues();
  var newData = [];
  
  for (let i in data){
    var row       = data[i];
    var duplicate = false;
    for (let j in newData){
      if(row[0] == newData[j][0]){
         duplicate = true;
      }
    }
    if (!duplicate){
      newData.push(row);
    }
  }
  
  sheet.clearContents();
  sheet.getRange(1,1, newData.length, newData[0].length).setValues(newData);
  
}


