function generatePDFsForSelectedRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Dynamically get the active sheet
  const lastRow = sheet.getLastRow();
  let createdInvoices = []; // To track created invoice numbers

  for (let row = 2; row <= lastRow; row++) {
    const checkboxCell = sheet.getRange(row, 17); // Q column (Checkbox)
    const isChecked = checkboxCell.getValue();

    if (isChecked === true) {
      const invoiceNumber = generateInvForRow(row);
      createdInvoices.push(invoiceNumber); // Store invoice number
    }
  }

  // After generating all PDFs, show an alert with the created invoice numbers
  const ui = SpreadsheetApp.getUi();
  if (createdInvoices.length > 0) {
    ui.alert('PDFs Created', createdInvoices.join(', ') + ' have been created successfully!', ui.ButtonSet.OK);
  } else {
    ui.alert('No PDFs Created', 'No rows were selected for PDF generation.', ui.ButtonSet.OK);
  }
}

function generateInvForRow(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Dynamically get the active sheet
  
  // Get data from columns B to S for the specified row
  const data = sheet.getRange(row, 2, 1, 18).getValues()[0];
  
  const admissionNumber = data[0]; // Column B
  const studentName = data[1]; // Column C
  const address = data[2]; // Column D
  const paidOn = data[14]; // Column P
  const mop = data[12]; // Column N
  const month = sheet.getRange('H1').getValue(); // Assuming 'Month' is in cell H1
  const subjects = data[5]; // Column G
  const advance = data[7]; // Column I
  const amntPaid = data[8]; // Column J
  const misc = data[9]; // Column K
  const total = data[10]; // Column L
  const balance = data[11]; // Column M
  const utrId = data[13]; // Column O
  const wrd = data[17]; // Column S
  
  // Generate a unique invoice number
  const invoiceNumber = generateUniqueInvoiceNumber(sheet, row);

  // Open the Google Slides presentation by ID
  const presentation = SlidesApp.openById('19m33SgH8_eeNdJ32s_R57BVzH3eG6WfF_Q4cvKRymb8');

  // Make a copy of the presentation to work with
  const copy = DriveApp.getFileById(presentation.getId()).makeCopy(studentName + '_' + month +'_Invoice');

  const copiedPresentation = SlidesApp.openById(copy.getId());
  const slides = copiedPresentation.getSlides();

  // Replace placeholders with actual data
  slides.forEach(function (slide) {
    const shapes = slide.getShapes();
    shapes.forEach(function (shape) {
      if (shape.getText()) {
        shape.getText().replaceAllText('{{Admission Number}}', admissionNumber);
        shape.getText().replaceAllText('{{Student Name}}', studentName);
        shape.getText().replaceAllText('{{ADDRESS}}', address);
        shape.getText().replaceAllText('{{InvoiceNumber}}', invoiceNumber);
        shape.getText().replaceAllText('{{Paid on}}', paidOn);
        shape.getText().replaceAllText('{{M.O.P}}', mop);
        shape.getText().replaceAllText('{{Month}}', month);
        shape.getText().replaceAllText('{{Subjects}}', subjects);
        shape.getText().replaceAllText('{{Balance}}', balance);
        shape.getText().replaceAllText('{{Amnt Paid}}', amntPaid);
        shape.getText().replaceAllText('{{Advance}}', advance);
        shape.getText().replaceAllText('{{Misc}}', misc);
        shape.getText().replaceAllText('{{Total}}', total);
        shape.getText().replaceAllText('{{words}}', wrd);
        shape.getText().replaceAllText('{{UTR id}}', utrId);
      }
    });
  });

  // Save and close the copied presentation
  copiedPresentation.saveAndClose();

  // Export the copy as a PDF
  const pdfBlob = copy.getAs('application/pdf');
  const pdfFileName = studentName + '_' + month +'_Invoice';
  const folder = DriveApp.getFolderById('1_sg8oA3K8st0jaWE03cIMOpeKpeojrpu'); // Update with your Google Drive folder ID
  folder.createFile(pdfBlob).setName(pdfFileName);

  // Write the generated invoice number into column R of the selected row (Status)
  sheet.getRange(row, 18).setValue(invoiceNumber); // Column R is the 18th column

  // Log success
  Logger.log('PDF created successfully for: ' + studentName);

  // Delete the copied presentation after exporting
  copy.setTrashed(true);

  // Return the invoice number to track it
  return invoiceNumber;
}

function generateUniqueInvoiceNumber(sheet, row) {
  const year = new Date().getFullYear();
  const month = sheet.getRange('H1').getValue(); // Assuming 'Month' is in cell H1
  const monthFormatted = month.toString().padStart(2, '0'); // Ensure two digits

  // Generate a new invoice number for the selected row
  let invoiceNumber = 1;
  const existingNumbers = sheet.getRange('R2:R').getValues().filter(row => row[0]); // Filter out empty cells in column R

  // Check for the last used number in column R (Status)
  existingNumbers.forEach(function(num) {
    const numStr = num[0].toString();
    if (numStr.startsWith(`INV-${monthFormatted}-${year}-`)) { // Use backticks for template literals
      const numSuffix = parseInt(numStr.split('-').pop(), 10);
      if (numSuffix >= invoiceNumber) {
        invoiceNumber = numSuffix + 1;
      }
    }
  });

  return `INV-${monthFormatted}-${year}-${invoiceNumber.toString().padStart(3, '0')}`;
}
