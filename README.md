# WECP Expense Allocation Tools

TypeScript/Apps Script project that supports WECP's expense allocation workflow.

## Development

- `npm install` – install dependencies.
- `npm run build` – compile TypeScript and copy the manifest into `dist/`.

## clasp Setup

1. Authenticate with clasp (one-time):
   - `npm run clasp login -- --no-localhost` (omit the flag if you prefer the browser flow).
2. Retrieve your Apps Script `scriptId` (open the target Spreadsheet → Extensions → Apps Script → Project Settings).
3. Update `.clasp.json` and replace `REPLACE_WITH_YOUR_SCRIPT_ID` with the ID from step 2.
4. Deploy the current build to Apps Script:
   - `npm run deploy`
   - or use `npm run clasp push` variants, such as `npm run clasp -- push --watch` for iterative development.
5. To pull remote script changes back into `dist/`, run `npm run clasp -- pull`.

`dist/` is the clasp root; always run `npm run build` before pushing.
