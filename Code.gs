const EXPENSE_SHEET_NAME = 'Expenses';
const INCOME_SHEET_NAME = 'Income';
const DEBT_SHEET_NAME = 'Debts';
const HEADERS = ['ID', 'Date', 'Title', 'Category', 'Amount', 'Note', 'Created At'];
const DEBT_HEADERS = ['ID', 'Date', 'Title', 'Category', 'Amount', 'Note', 'Status', 'Created At'];

function doGet() {
  ensureSheet_(EXPENSE_SHEET_NAME);
  ensureSheet_(INCOME_SHEET_NAME);
  ensureSheet_(DEBT_SHEET_NAME, DEBT_HEADERS);
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Budget Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupBudgetTracker() {
  const expenseSheet = ensureSheet_(EXPENSE_SHEET_NAME);
  const incomeSheet = ensureSheet_(INCOME_SHEET_NAME);
  ensureSheet_(DEBT_SHEET_NAME, DEBT_HEADERS);
  const now = new Date();

  if (expenseSheet.getLastRow() === 1) {
    expenseSheet.appendRow([Utilities.getUuid(), new Date(2026, 4, 9), 'Ice candy', 'Food', 20, '', now]);
    expenseSheet.appendRow([Utilities.getUuid(), new Date(2026, 4, 9), 'Lunch', 'Food', 200, '', now]);
  }

  if (incomeSheet.getLastRow() === 1) {
    incomeSheet.appendRow([Utilities.getUuid(), new Date(2026, 4, 8), 'Allowance', 'Allowance', 1000, '', now]);
  }

  return getDashboardData();
}

function getDashboardData() {
  const expenses = readTransactions_(EXPENSE_SHEET_NAME, 'expense');
  const incomes = readTransactions_(INCOME_SHEET_NAME, 'income');
  const debts = readTransactions_(DEBT_SHEET_NAME, 'debt');

  const total = expenses.reduce((sum, expense) => sum + expense.amount, 0);
  const incomeTotal = incomes.reduce((sum, income) => sum + income.amount, 0);

  const totalBorrowed = debts
    .filter(d => d.category === 'Borrowed' && d.status === 'Open')
    .reduce((sum, d) => sum + d.amount, 0);

  const totalLent = debts
    .filter(d => d.category === 'Lent' && d.status === 'Open')
    .reduce((sum, d) => sum + d.amount, 0);

  const transactions = expenses.concat(incomes).concat(debts)
    .sort((a, b) => new Date(b.date) - new Date(a.date));

  return {
    total,
    incomeTotal,
    totalBorrowed,
    totalLent,
    balance: incomeTotal - total,
    transactions,
    latest: transactions,
    categories: getCategories_(),
    incomeCategories: getIncomeCategories_(),
    debtCategories: getDebtCategories_()
  };
}

function addExpense(expense) {
  appendTransaction_(EXPENSE_SHEET_NAME, normalizeTransaction_(expense, 'Expense'));
  return getDashboardData();
}

function addIncome(income) {
  appendTransaction_(INCOME_SHEET_NAME, normalizeTransaction_(income, 'Income'));
  return getDashboardData();
}

function updateExpense(expense) {
  updateTransaction_(EXPENSE_SHEET_NAME, expense, 'Expense');
  return getDashboardData();
}

function updateIncome(income) {
  updateTransaction_(INCOME_SHEET_NAME, income, 'Income');
  return getDashboardData();
}

function deleteExpense(id) {
  deleteTransaction_(EXPENSE_SHEET_NAME, id);
  return getDashboardData();
}

function deleteIncome(id) {
  deleteTransaction_(INCOME_SHEET_NAME, id);
  return getDashboardData();
}

function addDebt(debt) {
  appendTransaction_(DEBT_SHEET_NAME, normalizeTransaction_(debt, 'Debt'));
  return getDashboardData();
}

function updateDebt(debt) {
  updateTransaction_(DEBT_SHEET_NAME, debt, 'Debt');
  return getDashboardData();
}

function deleteDebt(id) {
  deleteTransaction_(DEBT_SHEET_NAME, id);
  return getDashboardData();
}

function toggleDebtStatus(id) {
  const sheet = ensureSheet_(DEBT_SHEET_NAME, DEBT_HEADERS);
  const rowNumber = findTransactionRow_(sheet, id);
  if (!rowNumber) {
    throw new Error('Debt was not found.');
  }

  const currentStatus = sheet.getRange(rowNumber, 7).getValue();
  const newStatus = currentStatus === 'Paid' ? 'Open' : 'Paid';
  sheet.getRange(rowNumber, 7).setValue(newStatus);

  return getDashboardData();
}

function appendTransaction_(sheetName, transaction) {
  const isDebt = sheetName === DEBT_SHEET_NAME;
  const sheet = isDebt ? ensureSheet_(sheetName, DEBT_HEADERS) : ensureSheet_(sheetName);
  const row = [
    Utilities.getUuid(),
    transaction.date,
    transaction.title,
    transaction.category,
    transaction.amount,
    transaction.note
  ];

  if (isDebt) {
    row.push(transaction.status || 'Open');
  }

  row.push(new Date());
  sheet.appendRow(row);
}

function updateTransaction_(sheetName, transaction, label) {
  const id = String(transaction && transaction.id || '').trim();
  if (!id) {
    throw new Error(label + ' ID is required.');
  }

  const normalized = normalizeTransaction_(transaction, label);
  const isDebt = sheetName === DEBT_SHEET_NAME;
  const sheet = isDebt ? ensureSheet_(sheetName, DEBT_HEADERS) : ensureSheet_(sheetName);
  const rowNumber = findTransactionRow_(sheet, id);
  if (!rowNumber) {
    throw new Error(label + ' was not found.');
  }

  const values = [
    normalized.date,
    normalized.title,
    normalized.category,
    normalized.amount,
    normalized.note
  ];

  if (isDebt) {
    values.push(normalized.status || 'Open');
    sheet.getRange(rowNumber, 2, 1, 6).setValues([values]);
  } else {
    sheet.getRange(rowNumber, 2, 1, 5).setValues([values]);
  }
}

function deleteTransaction_(sheetName, id) {
  const sheet = ensureSheet_(sheetName);
  const rowNumber = findTransactionRow_(sheet, id);
  if (rowNumber) {
    sheet.deleteRow(rowNumber);
  }
}

function normalizeTransaction_(transaction, label) {
  const title = String(transaction && transaction.title || '').trim();
  const amount = Number(transaction && transaction.amount);
  if (!title) {
    throw new Error(label + ' title is required.');
  }
  if (!amount || amount < 0) {
    throw new Error('Amount must be greater than 0.');
  }

  return {
    title,
    amount,
    date: transaction.date ? new Date(transaction.date) : new Date(),
    category: String(transaction.category || 'Other').trim(),
    note: String(transaction.note || '').trim(),
    status: String(transaction.status || 'Open').trim()
  };
}

function readTransactions_(sheetName, type) {
  const isDebt = sheetName === DEBT_SHEET_NAME;
  const sheet = isDebt ? ensureSheet_(sheetName, DEBT_HEADERS) : ensureSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  return values.slice(1).filter(row => row[0]).map(row => {
    const transaction = {
      id: row[0],
      type,
      date: normalizeDate_(row[1]),
      title: row[2],
      category: row[3] || 'Other',
      amount: Number(row[4]) || 0,
      note: row[5] || ''
    };

    if (isDebt) {
      transaction.status = row[6] || 'Open';
      transaction.createdAt = normalizeDate_(row[7]);
    } else {
      transaction.createdAt = normalizeDate_(row[6]);
    }

    return transaction;
  });
}

function findTransactionRow_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      return i + 1;
    }
  }
  return 0;
}

function ensureSheet_(sheetName, headers = HEADERS) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = headers.every((header, index) => existingHeaders[index] === header);
  if (!hasHeaders) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function getCategories_() {
  return ['Food', 'Transport', 'Bills', 'Shopping', 'School', 'Health', 'Savings', 'Other'];
}

function getIncomeCategories_() {
  return ['Salary', 'Allowance', 'Freelance', 'Gift', 'Business', 'Savings', 'Other'];
}

function getDebtCategories_() {
  return ['Borrowed', 'Lent'];
}

function normalizeDate_(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
