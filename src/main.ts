/**
 * Adds the WECP custom menu when the spreadsheet is opened.
 */
const SHEET_NAMES = {
  transactions: 'Transactions',
  chartCodes: 'ChartCodes',
  formTemplate: 'FormTemplate',
} as const;

const TRANSACTION_HEADERS: string[] = [
  'Account Number',
  'Post Date',
  'Check',
  'Description',
  'Debit',
  'Credit',
  'Status',
  'Balance',
  'Category Code',
  'Category Description',
  'Prepared By',
  'Notes',
  'Submitted',
  'Submitted At',
] ;

const TRANSACTION_SAMPLE_ROW: (string | number | boolean | Date)[] = [
  'XXXX1234',
  new Date('2023-09-15'),
  '5648',
  'Read-a-thon supplies',
  0,
  100.1,
  'Posted',
  5825.45,
  '66400',
  'Toys/Equipment Expense',
  'Miss Teacher',
  'Split with classroom supplies',
  false,
  '',
] ;

const CHART_CODE_TABLE: (string | boolean)[][] = [
  ['Code', 'Description', 'Category Type', 'Is Deposit', 'Notes'],
  ['40300', 'Fee Refunds', 'Expense', false, 'Matches form row'],
  ['62200', 'Professional Development', 'Expense', false, ''],
  ['63000', 'Fundraising: Read a thon', 'Expense', false, 'Fundraising events'],
  ['64000', 'School Dist Teacher Consult', 'Expense', false, ''],
  ['64100', 'In-Out Expense', 'Expense', false, 'Specify in Notes'],
  ['65000', 'Professional Consult Fees', 'Expense', false, ''],
  ['65100', 'ICC Fees', 'Expense', false, ''],
  ['65200', 'CC Tuition', 'Expense', false, ''],
  ['65500', 'Legal/License Fees', 'Expense', false, ''],
  ['66000', 'Curriculum Expenses', 'Expense', false, ''],
  ['66200', 'Other Supplies', 'Expense', false, ''],
  ['66300', 'Class Enrichment Prof.', 'Expense', false, ''],
  ['66400', 'Toys/Equipment Expense', 'Expense', false, ''],
  ['66500', 'Fieldtrips', 'Expense', false, ''],
  ['67000', 'Sunshine Fund', 'Expense', false, ''],
  ['67100', 'Membership Events', 'Expense', false, ''],
  ['67200', 'Meetings and Speakers', 'Expense', false, ''],
  ['68000', 'Rent', 'Expense', false, ''],
  ['68100', 'Utilities/Storage', 'Expense', false, ''],
  ['68200', 'Telephone', 'Expense', false, ''],
  ['68300', 'Maintenance Expenses', 'Expense', false, ''],
  ['69000', 'Office Supplies', 'Expense', false, ''],
  ['69100', 'Promotion & Marketing', 'Expense', false, ''],
  ['69200', 'Website/Publ. Expense', 'Expense', false, ''],
  ['69300', 'Insurance', 'Expense', false, ''],
  ['90000', 'Incoming Deposit', 'Income', true, 'Blocked from expense allocation'],
];

const FORM_TEMPLATE_HEADER_ROWS: (string | number)[][] = [
  ['Co-op Name', '', '{{CoopName}}'],
  ['Expense Allocation', '', ''],
  ['Name', '', '{{PreparedBy}}', '', 'Total', '{{TotalAmount}}'],
  ['Date', '', '{{ReportDate}}'],
  ['Check #', '', '{{CheckNumber}}', '', 'Amount', '{{CheckAmount}}'],
];

const FORM_TEMPLATE_TABLE_ROWS: (string | number)[][] = [
  ['Expense Code', 'Expense Description', 'Amount'],
  ['40300', 'Fee Refunds', ''],
  ['62200', 'Professional Development', ''],
  ['63000', 'Fundraising: List Event- Read a thon', ''],
  ['64000', 'School Dist Teacher Consult', ''],
  ['64100', 'In-Out Expense: Specify-', ''],
  ['65000', 'Professional Consult Fees', ''],
  ['65100', 'ICC Fees', ''],
  ['65200', 'CC Tuition', ''],
  ['65500', 'Legal/License Fees', ''],
  ['66000', 'Curriculum Expenses', ''],
  ['66200', 'Other Supplies', ''],
  ['66300', 'Class Enrichment Prof.', ''],
  ['66400', 'Toys/Equipment Expense', ''],
  ['66500', 'Fieldtrips', ''],
  ['67000', 'Sunshine Fund', ''],
  ['67100', 'Membership Events', ''],
  ['67200', 'Meetings and Speakers', ''],
  ['68000', 'Rent', ''],
  ['68100', 'Utilities/Storage', ''],
  ['68200', 'Telephone', ''],
  ['68300', 'Maintenance Expenses', ''],
  ['69000', 'Office Supplies', ''],
  ['69100', 'Promotion & Marketing', ''],
  ['69200', 'Website/Publ. Expense', ''],
  ['69300', 'Insurance', ''],
  ['Reserve', 'Specify-', ''],
  ['Other', 'Specify-', ''],
];

function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('WECP')
    .addItem('Launch Expense Tools', 'showWelcomeSidebar')
    .addItem('Import Transactions', 'importTransactionsStub')
    .addItem('Setup Workbook Tabs', 'setupWorkbookScaffold')
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

/**
 * Ensures the Transactions, ChartCodes, and FormTemplate sheets exist with base structure.
 */
function setupWorkbookScaffold(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ensureTransactionsSheet(spreadsheet);
  ensureChartCodesSheet(spreadsheet);
  ensureFormTemplateSheet(spreadsheet);
  spreadsheet.toast('Transactions, ChartCodes, and FormTemplate tabs are ready.', 'WECP Expenses', 5);
}

function ensureTransactionsSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
  const sheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.transactions);
  sheet.clearFormats();
  sheet.getRange(1, 1, 1, TRANSACTION_HEADERS.length).setValues([TRANSACTION_HEADERS]);
  sheet.getRange(1, 1, 1, TRANSACTION_HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, TRANSACTION_HEADERS.length);

  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, 1, TRANSACTION_HEADERS.length).setValues([TRANSACTION_SAMPLE_ROW]);
    sheet.getRange(2, 2).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(2, 5, 1, 4).setNumberFormat('#,##0.00');
    sheet.getRange(2, 14).setNumberFormat('yyyy-mm-dd hh:mm');
    sheet.getRange('M2').insertCheckboxes();
  } else {
    ensureCheckboxColumn(sheet, 'Submitted', TRANSACTION_HEADERS.indexOf('Submitted') + 1);
  }
}

function ensureChartCodesSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
  const sheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.chartCodes);
  sheet.clearFormats();
  sheet.getRange(1, 1, CHART_CODE_TABLE.length, CHART_CODE_TABLE[0].length).setValues(CHART_CODE_TABLE);
  sheet.getRange(1, 1, 1, CHART_CODE_TABLE[0].length).setFontWeight('bold');
  sheet.autoResizeColumns(1, CHART_CODE_TABLE[0].length);
  sheet.setFrozenRows(1);
}

function ensureFormTemplateSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
  const sheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.formTemplate);
  sheet.clear();
  const headerWidth = FORM_TEMPLATE_HEADER_ROWS.reduce((max, row) => Math.max(max, row.length), 0);
  const paddedHeaderRows = FORM_TEMPLATE_HEADER_ROWS.map((row) => {
    if (row.length === headerWidth) {
      return row;
    }
    const paddedRow = row.slice();
    while (paddedRow.length < headerWidth) {
      paddedRow.push('');
    }
    return paddedRow;
  });
  sheet.getRange(1, 1, paddedHeaderRows.length, headerWidth).setValues(paddedHeaderRows);
  const tableStartRow = FORM_TEMPLATE_HEADER_ROWS.length + 2;
  sheet.getRange(tableStartRow, 1, FORM_TEMPLATE_TABLE_ROWS.length, FORM_TEMPLATE_TABLE_ROWS[0].length).setValues(
    FORM_TEMPLATE_TABLE_ROWS,
  );
  sheet.getRange(tableStartRow, 1, 1, FORM_TEMPLATE_TABLE_ROWS[0].length).setFontWeight('bold');
  sheet.setColumnWidths(1, FORM_TEMPLATE_TABLE_ROWS[0].length, 200);
  sheet.getRange(tableStartRow + 1, 3, FORM_TEMPLATE_TABLE_ROWS.length - 1, 1).setNumberFormat('#,##0.00');
  sheet.setFrozenRows(tableStartRow - 1);
  sheet.autoResizeColumns(1, 2);
}

function getOrCreateSheet(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  name: string,
): GoogleAppsScript.Spreadsheet.Sheet {
  return spreadsheet.getSheetByName(name) ?? spreadsheet.insertSheet(name);
}

function ensureCheckboxColumn(sheet: GoogleAppsScript.Spreadsheet.Sheet, header: string, columnIndex: number): void {
  const headerValue = sheet.getRange(1, columnIndex).getValue();
  if (headerValue !== header) {
    sheet.getRange(1, columnIndex).setValue(header);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, columnIndex, lastRow - 1, 1).insertCheckboxes();
  }
}
