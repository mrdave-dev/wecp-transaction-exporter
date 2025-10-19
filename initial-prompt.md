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