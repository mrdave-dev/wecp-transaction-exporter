/**
 * Adds the WECP custom menu when the spreadsheet is opened.
 */
function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('WECP')
    .addItem('Launch Expense Tools', 'showWelcomeSidebar')
    .addItem('Import Transactions', 'importTransactionsStub')
    .addToUi();
}

/**
 * Placeholder sidebar to verify deployment wiring.
 */
function showWelcomeSidebar(): void {
  const html = HtmlService.createHtmlOutput(
    `<div style="padding:16px;font-family:Roboto,Arial,sans-serif;">
       <h2 style="margin-top:0;">WECP Expenses</h2>
       <p>This is a stub sidebar. Implementation coming soon.</p>
     </div>`,
  );
  html.setTitle('WECP Expenses');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Placeholder import handler until CSV ingestion is implemented.
 */
function importTransactionsStub(): void {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Import workflow not yet implemented.',
    'WECP Expenses',
    5,
  );
}
