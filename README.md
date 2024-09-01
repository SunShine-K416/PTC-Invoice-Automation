# PTC-Invoice-Automation

This project automates invoice generation and fee reminders using Google Apps Script. It generates PDF invoices from Google Sheets, sends WhatsApp reminders for pending fees, and updates statuses based on checkbox selections. Ideal for managing fee processing at tuition centers.

## Features
- **Generate PDF Invoices**: Automatically create and save invoices as PDFs using Google Slides templates.
- **Send WhatsApp Reminders**: Send fee reminders via WhatsApp using customized messages for each student.
- **Checkbox Status**: Automates selection with checkboxes for invoice generation and message sending.
- **Status Updates**: Keeps track of invoices and reminders directly in the Google Sheet.

## Google Apps Script Functions
### 1. `generatePDFsForSelectedRows()`
This function generates PDF invoices for rows where the checkbox in column Q is checked.

### 2. `generateInvForRow(row)`
Creates a copy of the Google Slides template, replaces placeholders with student data, and exports it as a PDF.

### 3. `generateUniqueInvoiceNumber(sheet, row)`
Generates a unique invoice number based on the current year and month.

### 4. `addCheckboxes()`
Automatically adds checkboxes in column Q (for invoice generation) and column T (for sending WhatsApp messages) for rows with data in column A.

### 5. `sendMessages()`
Sends a WhatsApp message reminder for fee payment to students whose checkbox in column T is selected.

### 6. `onOpen()`
Adds a custom menu option in the Google Sheet for easier access to add checkboxes.

## Prerequisites
- Google Sheets with student data.
- Google Slides template for invoice generation.
- Google Apps Script attached to your Google Sheet for automation.
- WhatsApp installed to receive reminders.

## How to Use
1. **Set Up Google Sheets**:
    - Ensure columns B to S contain student data (Admission Number, Name, Fees, etc.).
    - Add checkboxes in columns Q (for generating invoices) and T (for sending reminders).

2. **Run Functions**:
    - Use the `generatePDFsForSelectedRows()` function to create PDF invoices for selected students.
    - Use the `sendMessages()` function to send WhatsApp fee reminders.

3. **Automate with Triggers**:
    - Set triggers in Google Apps Script to run functions automatically on form submissions or at specified intervals.

## Files in the Project
- **Google Sheet**: Stores student data, payment information, and checkbox statuses.
- **Google Slides Template**: Used for generating the invoices with placeholders.
- **Google Apps Script**: Contains the logic for automating the invoice generation and WhatsApp reminders.

