//=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--
//<START> Create email function
//=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--

function myFunction1() {
  //Creates document
  const doc   =  DocumentApp.create('Test 20'); // <-- test 20 is a unique email name --- this can be changed
  
  //Opens document - address body
  const body   = doc.getBody();
  
  //Content of body
  body.appendParagraph('Dear NoPhin:');
  body.appendParagraph(''); //< - - Blank line
  body.appendParagraph('This email was generated to notify the company that the following email address sent us a notfication/order that was not in the accepted format.');
  body.appendParagraph(''); //< - - Blank line
  body.appendParagraph('The email address is: ');
  body.appendParagraph('<email address>');
  //End of body content
  
}
//=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--
//<End> of email creation function
//=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--

//=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--
//<START> Sends email function
//=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--
function myFunction2(){
  const id     =  '1LdOi0HuRnJx2ipaTNw7LIEJlN_kJsu_g0gq5HXvRqzA';
  const email  =  Session.getActiveUser().getEmail();
  const doc    =  DocumentApp.openById(id);
  const body   =  doc.getBody();
  
  let emailContent = body.editAsText().getText();
  
  Logger.log(email);
  //doc.setName('Notification to NoPhin -- non-conforming email format!');
  
  const subject =  doc.getName();
  
  Logger.log(doc.getUrl());
  
  const url = doc.getUrl();
  
  emailContent += ' ' + url;
  
  GmailApp.sendEmail(email,subject,emailContent);
  MailApp.sendEmail(email,subject,emailContent); // <-- does this address multiple platforms??

}
//=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--
//End of Sending email function
//=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--
