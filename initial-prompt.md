# Initial prompt

I would like you to help me make a Google App Script application that assists a non-profit treasurer in creating expense allocation reports for printing. 

## Manual process overview

Let me begin by describe the process as it exists today. 

The treasurer is responsible for overseeing the non-profit's ("WECP") finances. Another position, vice chair, typically does the purchasing for the organization. 

The treasurer may also write checks for rent, utilities and member reimbursements. 

Each of these expenses must be given a chart code (e.g. 42000, or 48100) that corresponds to the expense category. 

Once a month, the treasurer must submit the expenses along with other financial documents to a business analyst at our parent organization, ECC. The analyst does WECP's bookkeeping and watches for potential fraud or abuse of our accounts.

Each transaction in our checking account requires an expense report. Expenses may split categories, but they typically do not. 

The treasurer exports the bank account transactions in to a CSV, copies them in to a Google Sheet, and uses the sheet to track which expenses have been submitted to ECC. 

---

Brainstorm how we might make this process easier using Google Apps Script

---

• Here are a few ways Google Apps Script could streamline the treasurer’s workflow:

  - Transaction import: Add a custom menu that ingests the bank’s CSV export, normalizes columns, flags duplicates, and drops the raw data into a “Transactions” sheet automatically.
  - Chart-code helpers: Maintain a master “Chart Codes” sheet and script smart suggestions—either via sidebar lookup or predictive dropdowns—to ensure expenses always get the correct code.
  - Submission tracking: Use a status column powered by checkboxes or a dropdown; Apps Script can stamp timestamps, author, and notes whenever status changes.
  - Supporting docs / receipts: Provide a simple form or sidebar upload that writes drive-file links into the expense row, so everything needed for ECC is in one place.
  - ECC packet generation: Build an “Export for ECC” action that compiles the pending rows into a formatted PDF or Google Doc with summary totals and optional receipt links, then emails it to the analyst with a consistent subject.
  - Audit aid: Keep a log sheet that records each submission batch (date, preparer, totals) automatically, making audits easier and providing quick history.
  - Validation & reminders: Add on-edit rules to ensure required fields are filled before submission and a timed trigger that emails weekly reminders for any open transactions.

---

## Future state

Here's the process I'd like to implement:

1. The treasurer semi-regularly uploads a CSV to a Google Sheet. The App Script ingests the sheet, and adds non-duplicate entries.
2. The treasurer starts at the top of the new transactions and starts categorizing them. As they begin to type the category, the category and number are auto-populated using some authoritative list of possible options.
3. The treasurer selects some number of rows to export in to the printed materials. The export sidebar menu allows the treasurer to select if the rows should be marked 'submitted' or not after they are exported. When exported, each transaction is turned in to a PDF of the expected expense allocation form, which is a template that can also live in another tab of the sheet. These PDFs are saved to a dated folder to be printed out later.
4. If the treasurer attempts to categorize a deposit, an error is shown. Deposits follow a different format. 

---

• Feasibility

  - Fully achievable with Google Apps Script: Sheets API handles CSV ingestion/deduping, Apps Script UI enables menus/sidebars, and Drive/Docs services generate templated PDFs.
  - Validation rules plus simple on-edit triggers cover category autocomplete, deposit blocking, and submitted-status stamping.
  - Apps Script quotas easily cover monthly volumes; only caution is keeping template-driven PDF generation efficient via caching (one doc template + per-row copies).

  Solution Shape

  - Core sheets: Transactions (imported data + status), ChartCodes (authoritative codes, descriptions, deposit flag), FormTemplate (one formatted expense form), ExportsLog (submission audit trail).
  - Custom menu: WECP → Import CSV, Categorize, Export Packet, Admin.
  - Sidebar workflow: category picker with autocomplete + metadata preview; export sidebar to choose rows, mark-as-submitted toggle, destination folder selector.
  - Script services: Drive folders dated YYYY-MM-DD, DocumentApp merge of template with transaction data, Sheets range metadata for status tracking, PropertiesService for persistent config (last import hash, default folder, etc.).

  Implementation Plan

  - Define sheet schema and data validation (headers, helper columns, submitted, isDeposit detection, protected ranges).
  - Build CSV importer (file picker → parse → normalize → hash on acct/date/amount → append only new entries → log results).
  - Implement category assist (dropdown driven by ChartCodes, on-edit trigger to auto-fill number/description, block deposits from expense categories).
  - Create sidebar UI (HTML service) for richer autocomplete, deposit warnings, quick links to docs/receipts.
  - Implement export workflow (collect selected rows, duplicate template doc per row, fill placeholders, combine to PDFs in dated Drive folder, mark submitted if chosen, log batch).
  - Add admin & QA tools (config dialog, refresh cache, audit log viewer) plus unit tests via Apps Script test framework or simple assertion functions.

  1. Confirm chart-of-accounts list structure and template fields so scripts can be parameterized instead of hard-coded.
  2. Decide on PDF packaging (single ZIP, merged PDF via Drive API, or per-transaction files) before implementation to avoid rework.

---

• Implementation Plan

  - Phase 1 – Sheet Scaffold: Define core tabs (Transactions, ChartCodes, FormTemplate, ExportsLog, optional Config),
    header schemas, named ranges, and initial validation rules; wire up script properties for defaults (template doc id,
    export root folder).
  - Phase 2 – CSV Importer: Build custom menu entry and dialog to accept Drive picker uploads, parse CSV with
    normalization (date, amount, memo fields), hash to detect duplicates, append new rows, and log import summary with
    row counts and rejects.
  - Phase 3 – Categorization UX: Implement sidebar with autocomplete using ChartCodes, on-edit triggers that stamp
    matched code/description, block edits on submitted rows, and auto-detect deposits to prevent expense categorization.
  - Phase 4 – Export Workflow: Create sidebar flow to gather selected rows, choose target folder, duplicate FormTemplate,
    merge transaction data, render PDFs (one per row) into a dated subfolder, and optionally mark rows submitted with
    timestamp/user in Transactions.
  - Phase 5 – Validation & Audit: Add data validation and error messaging for required fields, ensure deposit rows route
    to alternate handling, append export batch metadata (date, preparer, totals, folder link) to ExportsLog, and maintain
    change log for troubleshooting.
  - Phase 6 – Testing & Deployment: Write clasp-compatible unit helpers or simple assertion functions for import/export
    logic, add mock datasets for local validation, document setup steps (triggers, permissions, template linkage), and
    prepare release checklist for production rollout.

---

Great! Everything's working. What should we implement next?


• We’re in good shape. Next logical pieces:

  1. Build the initial Transactions, ChartCodes, and FormTemplate tabs with headers + sample data; that lets us start
     testing sheet-driven logic.
  2. Wire a typed Sheet service layer (range access helpers, schema enums) so import/export code stays consistent.
  3. Implement the actual CSV importer stub—file picker dialog, Drive file read, row dedupe, append.

  Any of those sound good; just let me know which one you want to tackle first.


› Let's start with number 1. What do you need from me before you start implementing?


• To scaffold those tabs cleanly, I need a bit of detail from you:

  - Transactions sheet: exact column headers/order you want (e.g., Date, Payee, Amount, Chart Code, Status, Notes, etc.)
    and any default formats or sample rows to seed.
  - ChartCodes sheet: columns to include (code, description, category type, isDeposit flag, maybe ECC notes?) plus a
    handful of starter entries if you have them.
  - FormTemplate sheet: what fields should appear on the expense form (and placeholder values) so we can lay out the
    template grid that future exports will clone.

  If you already have a legacy sheet or doc we should mirror, feel free to summarize or link the structure. Once I have
  that, I’ll wire up the initial tabs.

---

The bank exports come in CSV format and contain the following headers:

Account Number
Post Date
Check
Description
Debit
Credit
Status
Balance

I've added 'Prepared' and 'Notes' columns but those can be kept, changed, or discarded.

The expense allocation form is pictured here: examples/2025-10-19_12-56-23.png

I have already parsed the form, here in CSV format:

```csv
Expense Code,Expense Description,Amount
40300,Fee Refunds,
62200,Professional Development,
63000,"Fundraising: List Event- Read a thon",12.00
64000,School Dist Teacher Consult,
64100,In-Out Expense: Specify-,
65000,Professional Consult Fees,
65100,ICC Fees,
65200,CC Tuition,
65500,Legal/License Fees,
66000,Curriculum Expenses,20.00
66200,Other Supplies,9.00
66300,Class Enrichment Prof.,
66400,Toys/Equipment Expense,59.10
66500,Fieldtrips,
67000,Sunshine Fund,
67100,Membership Events,
67200,Meetings and Speakers,
68000,Rent,
68100,Utilities/Storage,
68200,Telephone,
68300,Maintenance Expenses,
69000,Office Supplies,
69100,Promotion & Marketing,
69200,Website/Publ. Expense,
69300,Insurance,
Reserve,Specify-,
Other,Specify-,
```