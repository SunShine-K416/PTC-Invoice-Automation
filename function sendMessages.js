function sendMessages() {
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Dynamically get the active sheet
  const lastRow = sheet.getLastRow();
  
  // Array to store the messages to be sent
  let messages = [];
  
  // Iterate through each row to check for selected checkboxes in column T
  for (let row = 2; row <= lastRow; row++) {
    const checkboxCell = sheet.getRange(row, 20); // T column (20th column)
    const isChecked = checkboxCell.getValue();
    
    if (isChecked === true) {
      // Get data from the row
      const studentName = sheet.getRange(row, 3).getValue(); // Column C (Student Name)
      const monthFeesBalance = sheet.getRange(row, 8).getValue(); // Column H (September)
      const balance = sheet.getRange(row, 13).getValue(); // Column M (Balance)
      const contactNumber = sheet.getRange(row, 6).getValue(); // Column F (Contact)
      const total = monthFeesBalance + balance;
      const dueDate = new Date();
      dueDate.setDate(dueDate.getDate() + 5); // Due date is 5 days from today
      const dueDateString = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'MMMM d, yyyy');
      
      // Ensure contact number is treated as a string and remove non-digit characters
      const contactNumberStr = contactNumber.toString().replace(/\D/g, ''); // Remove any non-digit characters
      const formattedNumber = `+91${contactNumberStr}`; // Add country code
      
      // Construct the message
      const message = `Dear Parent, 👋

Reminder: ${studentName}’s fee is due.

🔹 Monthly Fees: ₹${monthFeesBalance}
🔹 Outstanding: ₹${balance}

*🔹 Total: ₹${total}*
*📅 Due Date: ${dueDateString}*

Pay via:
💰 GPay/PhonePe/Paytm: 9659055137
💰 Cash

Questions? Call 9659055137. 📞

Please pay by the due date. Thank you!

PTC BILL DESK
Perfect Tuition Center
_Automated message._`;
      
      // Encode message for URL
      const encodedMessage = encodeURIComponent(message);
      // Create WhatsApp URL
      const url = `https://api.whatsapp.com/send?phone=${formattedNumber}&text=${encodedMessage}`;
      
      messages.push({url: url, row: row, studentName: studentName});
    }
  }
  
  // Open all URLs in new tabs and mark success in the sheet
  messages.forEach(msg => {
    // Open the URL in a new tab
    const html = HtmlService.createHtmlOutput(`<html><script>window.open('${msg.url}');google.script.host.close();</script></html>`);
    SpreadsheetApp.getUi().showModalDialog(html, 'Sending Message');
    
    // Update success message in column U
    sheet.getRange(msg.row, 21).setValue(`Reminder sent to ${msg.studentName}`);
  });
}
