
# Google Sheets Budget Tracker

Mobile-first budget tracker web app for Google Sheets and Apps Script.

## Files

- `Code.gs` contains the Google Sheets setup and Apps Script functions.
- `Index.html` contains the iOS-style dashboard and add-expense dialog.
- `appsscript.json` is the Apps Script manifest for V8 and Asia/Manila time.

## Setup

1. Create a Google Sheet.
2. Open **Extensions > Apps Script**.
3. Copy `Code.gs` into the Apps Script editor.
4. Create an HTML file named `Index` and copy `Index.html` into it.
5. Run `setupBudgetTracker` once and approve permissions.
6. Deploy with **Deploy > New deployment > Web app**.
7. Set access to the audience you want, then open the web app URL.

The app creates `Expenses` and `Income` sheets with these columns: `ID`, `Date`, `Title`, `Category`, `Amount`, `Note`, and `Created At`.

You can add, edit, and delete expenses and income from the web app. Edit and delete controls appear beside each latest transaction row.
