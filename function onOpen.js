function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Add Checkboxes', 'addCheckboxes')
    .addToUi();
}

function addCheckboxes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  // Iterate through each row to add checkboxes
  for (let row = 2; row <= lastRow; row++) { // Starting from row 2 assuming row 1 is header
    const dataInColumnA = sheet.getRange(row, 1).getValue(); // Column A
    const checkboxCellQ = sheet.getRange(row, 17); // Q column (17th column)
    const checkboxCellT = sheet.getRange(row, 20); // T column (20th column)
    
    // Add a checkbox in column Q and T if data is present in column A
    if (dataInColumnA) {
      if (checkboxCellQ.getValue() === '') {
        checkboxCellQ.insertCheckboxes();
      }
      if (checkboxCellT.getValue() === '') {
        checkboxCellT.insertCheckboxes();
      }
    } else {
      // Clear checkbox if no data in column A
      checkboxCellQ.setValue('');
      checkboxCellT.setValue('');
    }
  }
}
