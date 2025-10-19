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

const SOURCE_TRANSACTION_HEADERS = [
  'Account Number',
  'Post Date',
  'Check',
  'Description',
  'Debit',
  'Credit',
  'Status',
  'Balance',
] as const;

const TRANSACTION_COLUMNS = {
  accountNumber: TRANSACTION_HEADERS.indexOf('Account Number') + 1,
  postDate: TRANSACTION_HEADERS.indexOf('Post Date') + 1,
  check: TRANSACTION_HEADERS.indexOf('Check') + 1,
  description: TRANSACTION_HEADERS.indexOf('Description') + 1,
  debit: TRANSACTION_HEADERS.indexOf('Debit') + 1,
  credit: TRANSACTION_HEADERS.indexOf('Credit') + 1,
  status: TRANSACTION_HEADERS.indexOf('Status') + 1,
  balance: TRANSACTION_HEADERS.indexOf('Balance') + 1,
  categoryCode: TRANSACTION_HEADERS.indexOf('Category Code') + 1,
  categoryDescription: TRANSACTION_HEADERS.indexOf('Category Description') + 1,
  preparedBy: TRANSACTION_HEADERS.indexOf('Prepared By') + 1,
  notes: TRANSACTION_HEADERS.indexOf('Notes') + 1,
  submitted: TRANSACTION_HEADERS.indexOf('Submitted') + 1,
  submittedAt: TRANSACTION_HEADERS.indexOf('Submitted At') + 1,
} as const;

type CsvImportPayload = {
  filename: string;
  content: string;
};

type ImportResult = {
  message: string;
  filename: string;
  totalRows: number;
  importedRows: number;
  duplicateRows: number;
  skippedRows: number;
  errors: string[];
};

type NormalizedTransactionRow = {
  key: string;
  values: (string | number | boolean | Date)[];
};

type HeaderIndexMap = Map<string, number>;

type ChartCodeEntry = {
  code: string;
  description: string;
  categoryType: string;
  isDeposit: boolean;
  notes: string;
};

type ChartCodeLookup = {
  mapByCode: Map<string, ChartCodeEntry>;
  mapByDescription: Map<string, ChartCodeEntry>;
};

type ChartCodeCache = ChartCodeLookup & {
  updatedAt: number;
};

const CHART_CODE_CACHE_TTL_MS = 5 * 60 * 1000;

let chartCodeCache: ChartCodeCache | null = null;

function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('WECP')
    .addItem('Launch Expense Tools', 'showWelcomeSidebar')
    .addItem('Import Transactions', 'showImportDialog')
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

function showImportDialog(): void {
  const html = createImportDialogHtml();
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Transactions CSV');
}

function createImportDialogHtml(): GoogleAppsScript.HTML.HtmlOutput {
  const html = HtmlService.createHtmlOutput(
    String.raw`<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
  </head>
  <body>
    <div class="container">
      <h2>Import Transactions CSV</h2>
      <p>Select the bank export CSV. New rows will be appended and duplicates skipped.</p>
      <input type="file" id="csvFile" accept=".csv" />
      <div class="actions">
        <button id="importBtn" disabled>Import</button>
        <button id="cancelBtn" type="button">Cancel</button>
      </div>
      <div id="status" class="status"></div>
    </div>
    <script>
      const fileInput = document.getElementById('csvFile');
      const importBtn = document.getElementById('importBtn');
      const cancelBtn = document.getElementById('cancelBtn');
      const status = document.getElementById('status');

      fileInput.addEventListener('change', () => {
        updateStatus('');
        importBtn.disabled = fileInput.files.length === 0;
      });

      cancelBtn.addEventListener('click', () => google.script.host.close());

      importBtn.addEventListener('click', () => {
        if (!fileInput.files.length) {
          updateStatus('Choose a CSV file first.', true);
          return;
        }
        const file = fileInput.files[0];
        const fileName = file && file.name ? file.name.trim() : '';
        if (!fileName.toLowerCase().endsWith('.csv')) {
          updateStatus('Please select a .csv file.', true);
          return;
        }
        importBtn.disabled = true;
        updateStatus('Reading file…', false);
        const reader = new FileReader();
        reader.onload = (event) => {
          updateStatus('Uploading to spreadsheet…', false);
          const content = event && event.target ? event.target.result : '';
          google.script.run
            .withSuccessHandler((result) => {
              const summary = result && result.message ? result.message : 'Import complete.';
              let detail = summary;
              if (result) {
                detail =
                  summary +
                  '\\nImported: ' +
                  result.importedRows +
                  ', Duplicates: ' +
                  result.duplicateRows +
                  ', Skipped: ' +
                  result.skippedRows;
              }
              updateStatus(detail, false);
              if (result && Array.isArray(result.errors) && result.errors.length) {
                const list = document.createElement('ul');
                list.className = 'error-list';
                result.errors.forEach((err) => {
                  const li = document.createElement('li');
                  li.textContent = err;
                  list.appendChild(li);
                });
                status.appendChild(list);
              }
              setTimeout(() => google.script.host.close(), 2500);
            })
            .withFailureHandler((err) => {
              const message = err && err.message ? err.message : String(err);
              updateStatus(message, true);
              importBtn.disabled = false;
            })
            .processImportedCsv({ filename: file.name, content: content || '' });
        };
        reader.onerror = () => {
          updateStatus('Could not read the selected file.', true);
          importBtn.disabled = false;
        };
        reader.readAsText(file);
      });

      function updateStatus(message, isError) {
        status.textContent = message;
        status.className = 'status' + (isError ? ' error' : '');
      }
    </script>
    <style>
      body {
        font-family: Roboto, Arial, sans-serif;
        margin: 0;
        padding: 16px;
      }
      .container {
        display: flex;
        flex-direction: column;
        gap: 10px;
      }
      h2 {
        font-size: 18px;
        margin: 0;
      }
      p {
        margin: 0;
        font-size: 13px;
        color: #333;
      }
      input[type="file"] {
        font-size: 13px;
      }
      .actions {
        display: flex;
        gap: 8px;
      }
      button {
        padding: 6px 12px;
        font-size: 13px;
      }
      .status {
        min-height: 24px;
        font-size: 12px;
        white-space: pre-line;
      }
      .status.error {
        color: #b00020;
      }
      .error-list {
        margin: 6px 0 0;
        padding-left: 18px;
        font-size: 12px;
        color: #b00020;
      }
    </style>
  </body>
</html>`,
  );
  html.setWidth(420);
  html.setHeight(260);
  html.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  return html;
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
  applyCategoryValidation(spreadsheet);
}

function ensureChartCodesSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
  const sheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.chartCodes);
  sheet.clearFormats();
  sheet.getRange(1, 1, CHART_CODE_TABLE.length, CHART_CODE_TABLE[0].length).setValues(CHART_CODE_TABLE);
  sheet.getRange(1, 1, 1, CHART_CODE_TABLE[0].length).setFontWeight('bold');
  sheet.autoResizeColumns(1, CHART_CODE_TABLE[0].length);
  sheet.setFrozenRows(1);
  invalidateChartCodeCache();
  applyCategoryValidation(spreadsheet);
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

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  if (!e || !e.range) {
    return;
  }
  const sheet = e.range.getSheet();
  const spreadsheet = sheet.getParent();
  const sheetName = sheet.getName();

  if (sheetName === SHEET_NAMES.chartCodes) {
    invalidateChartCodeCache();
    applyCategoryValidation(spreadsheet);
    return;
  }

  if (sheetName !== SHEET_NAMES.transactions) {
    return;
  }

  const firstColumn = e.range.getColumn();
  const lastColumn = firstColumn + e.range.getNumColumns() - 1;
  const affectsCategoryCode =
    firstColumn <= TRANSACTION_COLUMNS.categoryCode && lastColumn >= TRANSACTION_COLUMNS.categoryCode;
  const affectsCategoryDescription =
    firstColumn <= TRANSACTION_COLUMNS.categoryDescription &&
    lastColumn >= TRANSACTION_COLUMNS.categoryDescription;

  if (!affectsCategoryCode && !affectsCategoryDescription) {
    return;
  }

  try {
    handleCategoryFieldsEdit(sheet, e.range);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    spreadsheet.toast(message, 'WECP Expenses', 5);
  }
}

function handleCategoryFieldsEdit(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  editedRange: GoogleAppsScript.Spreadsheet.Range,
): void {
  const spreadsheet = sheet.getParent();
  const chartCodeLookup = getChartCodeLookup(spreadsheet);
  const editedColumns = new Set<number>();
  for (let offset = 0; offset < editedRange.getNumColumns(); offset += 1) {
    editedColumns.add(editedRange.getColumn() + offset);
  }
  const codeEdited = editedColumns.has(TRANSACTION_COLUMNS.categoryCode);
  const descriptionEdited = editedColumns.has(TRANSACTION_COLUMNS.categoryDescription);
  if (!codeEdited && !descriptionEdited) {
    return;
  }
  const startRow = editedRange.getRow();
  const numRows = editedRange.getNumRows();
  const messages = new Set<string>();

  for (let offset = 0; offset < numRows; offset += 1) {
    const rowIndex = startRow + offset;
    const codeCell = sheet.getRange(rowIndex, TRANSACTION_COLUMNS.categoryCode);
    const descriptionCell = sheet.getRange(rowIndex, TRANSACTION_COLUMNS.categoryDescription);
    const rawCode = toTrimmedString(codeCell.getValue());
    const rawDescription = toTrimmedString(descriptionCell.getValue());

    if (codeEdited && !rawCode && (!descriptionEdited || !rawDescription)) {
      descriptionCell.clearContent();
      continue;
    }

    if (descriptionEdited && !rawDescription && (!codeEdited || !rawCode)) {
      codeCell.clearContent();
      continue;
    }

    const entryFromCode = rawCode ? chartCodeLookup.mapByCode.get(rawCode) : undefined;
    if (codeEdited && rawCode && !entryFromCode) {
      codeCell.clearContent();
      descriptionCell.clearContent();
      messages.add(`Row ${rowIndex}: Unknown category code "${rawCode}".`);
      continue;
    }

    const entryFromDescription = rawDescription
      ? chartCodeLookup.mapByDescription.get(normalizeDescriptionKey(rawDescription))
      : undefined;
    if (descriptionEdited && rawDescription && !entryFromDescription) {
      codeCell.clearContent();
      descriptionCell.clearContent();
      messages.add(`Row ${rowIndex}: Unknown category description "${rawDescription}".`);
      continue;
    }

    let entry: ChartCodeEntry | undefined;
    if (codeEdited && entryFromCode) {
      entry = entryFromCode;
    } else if (descriptionEdited && entryFromDescription) {
      entry = entryFromDescription;
    } else if (entryFromCode) {
      entry = entryFromCode;
    } else if (entryFromDescription) {
      entry = entryFromDescription;
    } else if (!rawCode && !rawDescription) {
      continue;
    } else {
      codeCell.clearContent();
      descriptionCell.clearContent();
      continue;
    }

    const debitValue = coerceNumber(sheet.getRange(rowIndex, TRANSACTION_COLUMNS.debit).getValue());
    const creditValue = coerceNumber(sheet.getRange(rowIndex, TRANSACTION_COLUMNS.credit).getValue());
    const isDepositRow = isDepositTransaction(debitValue, creditValue);

    if (isDepositRow && !entry.isDeposit) {
      codeCell.clearContent();
      descriptionCell.clearContent();
      messages.add(`Row ${rowIndex}: Deposits must use a deposit-specific category.`);
      continue;
    }

    if (!isDepositRow && entry.isDeposit) {
      codeCell.clearContent();
      descriptionCell.clearContent();
      messages.add(`Row ${rowIndex}: Deposit categories can only be used on deposit transactions.`);
      continue;
    }

    if (codeCell.getValue() !== entry.code) {
      codeCell.setValue(entry.code);
    }
    if (descriptionCell.getValue() !== entry.description) {
      descriptionCell.setValue(entry.description || '');
    }
  }

  if (messages.size > 0) {
    const combined = Array.from(messages).join('\n');
    spreadsheet.toast(combined, 'WECP Expenses', 6);
  }
}

function applyCategoryValidation(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
  const transactionsSheet = spreadsheet.getSheetByName(SHEET_NAMES.transactions);
  if (!transactionsSheet) {
    return;
  }

  const maxRows = transactionsSheet.getMaxRows();
  if (maxRows <= 1) {
    return;
  }

  const codeColumnRange = transactionsSheet.getRange(
    2,
    TRANSACTION_COLUMNS.categoryCode,
    maxRows - 1,
    1,
  );
  const descriptionColumnRange = transactionsSheet.getRange(
    2,
    TRANSACTION_COLUMNS.categoryDescription,
    maxRows - 1,
    1,
  );

  const chartCodesSheet = spreadsheet.getSheetByName(SHEET_NAMES.chartCodes);
  if (!chartCodesSheet || chartCodesSheet.getLastRow() <= 1) {
    codeColumnRange.clearDataValidations();
    descriptionColumnRange.clearDataValidations();
    return;
  }

  const recordCount = chartCodesSheet.getLastRow() - 1;
  const codesRange = chartCodesSheet.getRange(2, 1, recordCount, 1);
  const descriptionsRange = chartCodesSheet.getRange(2, 2, recordCount, 1);

  const codeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(codesRange, true)
    .setAllowInvalid(false)
    .setHelpText('Select a category code from the ChartCodes tab.')
    .build();

  const descriptionValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(descriptionsRange, true)
    .setAllowInvalid(false)
    .setHelpText('Select a category description from the ChartCodes tab.')
    .build();

  codeColumnRange.setDataValidation(codeValidation);
  descriptionColumnRange.setDataValidation(descriptionValidation);
}

function invalidateChartCodeCache(): void {
  chartCodeCache = null;
}

function getChartCodeLookup(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
): ChartCodeLookup {
  const now = Date.now();
  if (chartCodeCache && now - chartCodeCache.updatedAt < CHART_CODE_CACHE_TTL_MS) {
    return chartCodeCache;
  }
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.chartCodes);
  if (!sheet) {
    throw new Error('ChartCodes sheet not found.');
  }
  const lookup = loadChartCodeEntries(sheet);
  chartCodeCache = { ...lookup, updatedAt: now };
  return chartCodeCache;
}

function loadChartCodeEntries(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): ChartCodeLookup {
  const mapByCode = new Map<string, ChartCodeEntry>();
  const mapByDescription = new Map<string, ChartCodeEntry>();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { mapByCode, mapByDescription };
  }
  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  values.forEach((row) => {
    const code = toTrimmedString(row[0]);
    if (!code) {
      return;
    }
    const entry: ChartCodeEntry = {
      code,
      description: toTrimmedString(row[1]),
      categoryType: toTrimmedString(row[2]),
      isDeposit: toBoolean(row[3]),
      notes: toTrimmedString(row[4]),
    };
    mapByCode.set(code, entry);
    const descriptionKey = normalizeDescriptionKey(entry.description);
    if (descriptionKey && !mapByDescription.has(descriptionKey)) {
      mapByDescription.set(descriptionKey, entry);
    }
  });
  return { mapByCode, mapByDescription };
}

function isDepositTransaction(debit: number, credit: number): boolean {
  return Math.abs(credit) > 0 && Math.abs(debit) === 0;
}

function processImportedCsv(payload: CsvImportPayload): ImportResult {
  if (!payload || !payload.content) {
    throw new Error('No CSV content received.');
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.transactions);
  if (!sheet) {
    throw new Error('Transactions sheet not found. Run the setup first.');
  }

  const csvContent = payload.content.replace(/^\uFEFF/, '');
  const parsedRows = Utilities.parseCsv(csvContent);
  const filteredRows = parsedRows.filter((row) => !isRowEmpty(row));

  if (filteredRows.length === 0) {
    throw new Error('The CSV file was empty.');
  }

  const headerRow = filteredRows[0].map((cell) => cell.trim());
  const headerIndex = buildHeaderIndex(headerRow);
  validateSourceHeaders(headerIndex);

  const existingKeys = collectTransactionKeys(sheet);
  const seenKeys = new Set(existingKeys);
  const newRows: (string | number | boolean | Date)[][] = [];
  let duplicateRows = 0;
  let skippedRows = 0;
  const errors: string[] = [];

  for (let i = 1; i < filteredRows.length; i += 1) {
    const rawRow = filteredRows[i];
    try {
      const normalized = normalizeTransactionCsvRow(rawRow, headerIndex);
      if (!normalized) {
        skippedRows += 1;
        continue;
      }
      if (seenKeys.has(normalized.key)) {
        duplicateRows += 1;
        continue;
      }
      seenKeys.add(normalized.key);
      newRows.push(normalized.values);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      errors.push(`Row ${i + 1}: ${message}`);
      skippedRows += 1;
    }
  }

  if (newRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, TRANSACTION_HEADERS.length).setValues(newRows);
    sheet
      .getRange(startRow, TRANSACTION_COLUMNS.postDate, newRows.length, 1)
      .setNumberFormat('yyyy-mm-dd');
    sheet
      .getRange(startRow, TRANSACTION_COLUMNS.debit, newRows.length, 2)
      .setNumberFormat('#,##0.00');
    sheet
      .getRange(startRow, TRANSACTION_COLUMNS.balance, newRows.length, 1)
      .setNumberFormat('#,##0.00');
    sheet
      .getRange(startRow, TRANSACTION_COLUMNS.submittedAt, newRows.length, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm');
    sheet.getRange(startRow, TRANSACTION_COLUMNS.submitted, newRows.length, 1).insertCheckboxes();
  }
  applyCategoryValidation(spreadsheet);

  const totalRows = filteredRows.length - 1;
  const message = `${newRows.length} new row${newRows.length === 1 ? '' : 's'} imported from ${
    payload.filename
  }. ${duplicateRows} duplicate${duplicateRows === 1 ? '' : 's'} skipped.`;

  spreadsheet.toast(message, 'WECP Import', 5);

  return {
    message,
    filename: payload.filename,
    totalRows,
    importedRows: newRows.length,
    duplicateRows,
    skippedRows,
    errors,
  };
}

function normalizeTransactionCsvRow(
  row: string[],
  headerIndex: HeaderIndexMap,
): NormalizedTransactionRow | null {
  const getCell = (header: string): string => {
    const index = headerIndex.get(header);
    if (index === undefined) {
      return '';
    }
    return row[index] ?? '';
  };

  const accountNumber = toTrimmedString(getCell('Account Number'));
  const rawPostDate = getCell('Post Date');
  const postDate = parseDateValue(rawPostDate);
  if (!postDate) {
    throw new Error(`Invalid Post Date "${rawPostDate}"`);
  }

  const description = toTrimmedString(getCell('Description'));
  const checkNumber = toTrimmedString(getCell('Check'));
  const status = toTrimmedString(getCell('Status'));
  const debit = parseMoneyValue(getCell('Debit'), true);
  const credit = parseMoneyValue(getCell('Credit'), true);
  const balance = parseMoneyValue(getCell('Balance'), true);

  if (!description && debit === 0 && credit === 0) {
    return null;
  }

  const key = buildTransactionKey(accountNumber, postDate, description, debit, credit, checkNumber);
  const values: (string | number | boolean | Date)[] = [
    accountNumber,
    postDate,
    checkNumber,
    description,
    debit,
    credit,
    status,
    balance,
    '',
    '',
    '',
    '',
    false,
    '',
  ];

  return { key, values };
}

function buildHeaderIndex(headerRow: string[]): HeaderIndexMap {
  const map: HeaderIndexMap = new Map();
  headerRow.forEach((label, index) => {
    map.set(label.trim(), index);
  });
  return map;
}

function validateSourceHeaders(headerIndex: HeaderIndexMap): void {
  const missing = SOURCE_TRANSACTION_HEADERS.filter((column) => !headerIndex.has(column));
  if (missing.length > 0) {
    throw new Error(`Missing required column(s): ${missing.join(', ')}`);
  }
}

function collectTransactionKeys(sheet: GoogleAppsScript.Spreadsheet.Sheet): Set<string> {
  const keys = new Set<string>();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return keys;
  }

  const existingRows = sheet
    .getRange(2, 1, lastRow - 1, TRANSACTION_HEADERS.length)
    .getValues();

  existingRows.forEach((row) => {
    const key = buildTransactionKeyFromExistingRow(row);
    if (key) {
      keys.add(key);
    }
  });

  return keys;
}

function buildTransactionKeyFromExistingRow(row: unknown[]): string | null {
  const accountNumber = toTrimmedString(row[0]);
  const postDate = parseDateValue(row[1] as string | Date);
  if (!postDate) {
    return null;
  }
  const checkNumber = toTrimmedString(row[2]);
  const description = toTrimmedString(row[3]);
  const debit = coerceNumber(row[4]);
  const credit = coerceNumber(row[5]);

  return buildTransactionKey(accountNumber, postDate, description, debit, credit, checkNumber);
}

function buildTransactionKey(
  accountNumber: string,
  postDate: Date,
  description: string,
  debit: number,
  credit: number,
  checkNumber: string,
): string {
  const normalizedDate = normalizeDate(postDate);
  return [
    normalizeKeyPart(accountNumber),
    formatDateKey(normalizedDate),
    normalizeKeyPart(checkNumber),
    normalizeDescription(description),
    debit.toFixed(2),
    credit.toFixed(2),
  ].join('|');
}

function normalizeDescriptionKey(value: string): string {
  return normalizeDescription(value);
}

function normalizeDescription(value: string): string {
  return normalizeKeyPart(value).replace(/\s+/g, ' ');
}

function normalizeKeyPart(value: string): string {
  return value ? value.trim().toUpperCase() : '';
}

function normalizeDate(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatDateKey(date: Date): string {
  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, '0');
  const day = `${date.getDate()}`.padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function parseDateValue(input: string | Date): Date | null {
  if (input instanceof Date && !Number.isNaN(input.getTime())) {
    return normalizeDate(input);
  }

  const value = toTrimmedString(input);
  if (!value) {
    return null;
  }

  const direct = new Date(value);
  if (!Number.isNaN(direct.getTime())) {
    return normalizeDate(direct);
  }

  const match = value.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (match) {
    const month = Number(match[1]);
    const day = Number(match[2]);
    let year = Number(match[3]);
    if (year < 100) {
      year += year >= 70 ? 1900 : 2000;
    }
    return new Date(year, month - 1, day);
  }

  return null;
}

function parseMoneyValue(input: string | number, allowBlank = false): number {
  if (typeof input === 'number') {
    return input;
  }
  const value = toTrimmedString(input);
  if (!value) {
    return allowBlank ? 0 : 0;
  }

  let sanitized = value.replace(/[$,]/g, '');
  let isNegative = false;
  if (sanitized.startsWith('(') && sanitized.endsWith(')')) {
    sanitized = sanitized.slice(1, -1);
    isNegative = true;
  }
  if (sanitized.startsWith('-')) {
    isNegative = true;
  }

  const numeric = Number(sanitized.replace(/^[+-]/, ''));
  if (Number.isNaN(numeric)) {
    throw new Error(`Invalid currency value "${value}"`);
  }

  return isNegative ? -numeric : numeric;
}

function toTrimmedString(value: unknown): string {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function coerceNumber(value: unknown): number {
  if (typeof value === 'number') {
    return value;
  }
  const stringValue = toTrimmedString(value);
  if (!stringValue) {
    return 0;
  }
  return parseMoneyValue(stringValue, true);
}

function toBoolean(value: unknown): boolean {
  if (typeof value === 'boolean') {
    return value;
  }
  const stringValue = toTrimmedString(value).toLowerCase();
  if (!stringValue) {
    return false;
  }
  return ['true', 'yes', 'y', '1', 't', 'checked'].includes(stringValue);
}

function isRowEmpty(row: string[]): boolean {
  return row.every((cell) => toTrimmedString(cell) === '');
}
