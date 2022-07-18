/**
 * Get the Directory tab of the Klesis Workspace and go through the birth day column to get the birthdays coming up within two weeks
 */
function birthdayNotification() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Directory");
  if (sheet == null) {
    return;
  }
  
  const firstColumn = "A"; // Name column
  const lastColumn = "I"; // Days Until birthday column
  const columnRange = firstColumn + ":" + lastColumn;
  const directoryRange = sheet.getRange(columnRange);
  const allCells = directoryRange.getValues();
  
  const numRows = sheet.getLastRow();
  let index = 0;
  let birthdays = {}; // This will be in name -> birthday date (date object)

  while( index < numRows) {
    const daysUntil = allCells[index][8];
    if (daysUntil != "" && (daysUntil <= 14 || daysUntil == 365)) { // if the date is coming up withint 2 weeks
      const name = allCells[index][0];
      birthdays[name] = allCells[index][7];
    }
    index++;
  }
  
  if (Object.keys(birthdays).length != 0) {
    sendEmail(birthdays);
  }
  
}


/**
 * Send an email to leads using the MailApp API.
 * @param birthdays object of name -> birth date (in date object)
 */
function sendEmail(birthdays) {
  const emailAddress = "ucb_klesis_leads@gpmail.org";
  const cc = "james.han@gpmail.org"; // this is because I am sending the email to the alias so I don't get the email
  const subject = "Upcoming Birthdays";
  
  let message = "Here is a generated message for the upcoming two weeks' birthdays<br><br>";
  
  message += generateBirthdayText(birthdays);
  
  MailApp.sendEmail(emailAddress, subject, message, { htmlBody: message, cc: cc });
}


/**
 * Generate the birthday text in name: birth date format. Birth date will be MMMM dd format (i.e. July 16)
 */
function generateBirthdayText(birthdays) {
  let message = "";
  Object.keys(birthdays).forEach(
    name => {    
      const birthdate = Utilities.formatDate(birthdays[name], 'America/Los_Angeles', 'MMMM dd');
      message += `<b>${name}<\/b>: ${birthdate}<br>`;
    });
  return message;
}