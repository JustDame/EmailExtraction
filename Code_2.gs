function myFunction1() {
  const doc = DocumentApp.create('Test 20');
  const body = doc.getBody();
  body.appendParagraph('Dear NoPhin:');
  body.appendParagraph('');
  body.appendParagraph('This email was generated to notify the company that the following email address sent us a notfication/order that was not in the accepted format.');
  body.appendParagraph('');
  body.appendParagraph('The email address is: ');
  body.appendParagraph('<email address>');
  
}

function myFunction2(){
  const id = '1LdOi0HuRnJx2ipaTNw7LIEJlN_kJsu_g0gq5HXvRqzA';
  const email = Session.getActiveUser().getEmail();
  const doc = DocumentApp.openById(id);
  const body = doc.getBody();
  let emailContent = body.editAsText().getText();
  Logger.log(email);
  //doc.setName('Notification to NoPhin -- non-conforming email format!');
  const subject = doc.getName();
  Logger.log(doc.getUrl());
  const url = doc.getUrl();
  emailContent += ' ' + url;
  GmailApp.sendEmail(email,subject,emailContent);


}
