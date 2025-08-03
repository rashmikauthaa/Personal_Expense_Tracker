/**
 * Personal Expense Tracker - Google Apps Script
 * Built according to README specifications
 * Features: Setup columns, update transactions, monthly summaries, next month sheet creation
 */

// Constants for the expense tracker
const COLUMNS = {
  DATE: 1,
  CATEGORY: 2,
  DESCRIPTION: 3,
  AMOUNT: 4,
  TYPE: 5,
  MODE: 6,
  BALANCE: 7
};

const HEADERS = ['Date', 'Category', 'Description', 'Amount', 'Type', 'Mode', 'Balance'];

const CATEGORIES = ['Food', 'Rent', 'Travel', 'Family', 'Misc', 'Groceries', 'Shopping'];
const TRANSACTION_TYPES = ['In pocket', 'Expense', 'Gift'];
const PAYMENT_MODES = ['Cash', 'UPI', 'Card', 'Bank'];

/**
 * Creates the custom menu when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Expense Tracker')
    .addItem('Setup Columns', 'setupColumns')
    .addSeparator()
    .addItem('Update Balances', 'updateBalances')
    .addSeparator()
    .addItem('Get Monthly Summary', 'getCurrentMonthlySummary')
    .addSeparator()
    .addItem('Initialise Next Month Sheet', 'generateNextMonthSheet')
    .addSeparator()
    .addItem('Clear Results', 'clearResults')
    .addToUi();
}

/**
 * Feature 1: Setup Columns
 * Formats the active sheet with the required column structure and data validation
 */
function setupColumns() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Clear existing content and start fresh
    sheet.clear();

    // ✅ Determine the month name based on the sheet name first
    let sheetName = sheet.getName();
    let currentMonth = '';
    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];

    // If sheet name starts with a month name, use it
    const firstWord = sheetName.split(' ')[0];
    if (months.includes(firstWord)) {
      currentMonth = sheetName; // Full sheet name like "February 2025"
    } else {
      currentMonth = getCurrentMonthName(); // fallback
    }

    // Rename sheet to proper month name if it's a default name
    const isDefaultName = sheetName.match(/^Sheet\d*$/);
    if (isDefaultName || sheetName === 'Sheet1' || sheetName === 'Untitled spreadsheet') {
      if (!ss.getSheetByName(currentMonth)) {
        sheet.setName(currentMonth);
      } else {
        let counter = 1;
        let newName = currentMonth + ' (' + counter + ')';
        while (ss.getSheetByName(newName)) {
          counter++;
          newName = currentMonth + ' (' + counter + ')';
        }
        sheet.setName(newName);
      }
      sheetName = sheet.getName();
    }

    // ✅ Now everything else stays the same
    const ui = SpreadsheetApp.getUi();
    const incomeResponse = ui.prompt('Enter your monthly stipend/salary for ' + currentMonth + ' (₹):', ui.ButtonSet.OK_CANCEL);

    if (incomeResponse.getSelectedButton() !== ui.Button.OK) return;

    const balanceResponse = ui.prompt('Enter your current account balance (₹):', ui.ButtonSet.OK_CANCEL);
    if (balanceResponse.getSelectedButton() !== ui.Button.OK) return;

    const monthlyIncome = Number(incomeResponse.getResponseText());
    const startingBalance = Number(balanceResponse.getResponseText()) || monthlyIncome;

    // Validate income > 0
    if (isNaN(monthlyIncome) || monthlyIncome <= 0) {
      SpreadsheetApp.getUi().alert(
        'Invalid Input',
        'Please enter a valid monthly income greater than 0.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return; // Stop execution
    }

    createMonthHeaderBlock(sheet, currentMonth, monthlyIncome, startingBalance);
    setupDataValidation(sheet);

    const dataRange = sheet.getRange(5, 1, 995, HEADERS.length);
    dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

    sheet.getRange(4, COLUMNS.DATE).setNote('Enter date in DD/MM/YYYY format');
    sheet.getRange(4, COLUMNS.DESCRIPTION).setNote('Enter details for your reference');
    sheet.getRange(4, COLUMNS.BALANCE).setNote('Auto-calculated after Update Transactions');

    SpreadsheetApp.getUi().alert('Success!', `Sheet "${sheetName}" has been set up successfully for ${currentMonth}.`, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to setup columns: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Sets up data validation for dropdown columns
 */
function setupDataValidation(sheet) {
  const maxRows = 1000; // Apply validation to first 1000 rows
  const startRow = 5; // Data starts from row 5 (after header block)

  // Category dropdown validation
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CATEGORIES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(startRow, COLUMNS.CATEGORY, maxRows - startRow + 1, 1).setDataValidation(categoryRule);

  // Type dropdown validation
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(TRANSACTION_TYPES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(startRow, COLUMNS.TYPE, maxRows - startRow + 1, 1).setDataValidation(typeRule);

  // Mode dropdown validation
  const modeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PAYMENT_MODES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(startRow, COLUMNS.MODE, maxRows - startRow + 1, 1).setDataValidation(modeRule);
}

/**
 * Updates the current account balance in the header
 */
function updateCurrentAccountBalance() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const ui = SpreadsheetApp.getUi();
    
    // Get current balance from header
    const currentBalance = sheet.getRange(2, 4).getValue();
    
    const response = ui.prompt(
      'Update Current Account Balance', 
      `Current balance: ₹${currentBalance || 0}\n\nEnter your new current account balance (₹):`, 
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) return;
    
    const newBalance = Number(response.getResponseText());
    
    if (isNaN(newBalance) || newBalance < 0) {
      ui.alert('Invalid Input', 'Please enter a valid positive number.', ui.ButtonSet.OK);
      return;
    }
    
    // Update the balance in the header
    sheet.getRange(2, 4).setValue(newBalance).setNumberFormat('₹#,##0.00');
    
    ui.alert('Success!', `Current account balance updated to ₹${newBalance.toLocaleString('en-IN')}`, ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to update current account balance: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Feature 2: Update Balances
 * Calculates running balance for all transactions with empty balance cells
 */
function updateBalances() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const startRow = 5; // Data starts from row 5 (after header block)

    if (lastRow < startRow) {
      SpreadsheetApp.getUi().alert('No Data', 'No transactions found to update.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Get all data at once for performance
    const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, HEADERS.length);
    const data = dataRange.getValues();
    const balances = data.map(row => row[COLUMNS.BALANCE - 1]); // Current balances

    const initialBalance = Number(sheet.getRange(2, 4).getValue()) || 0;
    let runningBalance = initialBalance;
    let updates = [];
    let updatedCount = 0;
    let finalBalance = initialBalance; // Track the final balance

    // Compute balances in memory
    for (let i = 0; i < data.length; i++) {
      const amount = Number(data[i][COLUMNS.AMOUNT - 1]) || 0;
      const type = data[i][COLUMNS.TYPE - 1];

      if (!amount || !type) continue; // Skip incomplete rows

      if (type === 'In pocket' || type === 'Gift') {
        runningBalance += amount;
      } else if (type === 'Expense') {
        runningBalance -= amount;
      }

      // Track the final balance from the last valid transaction
      if (amount > 0 && type) {
        finalBalance = runningBalance;
      }

      // Update the balance in the array
      balances[i] = runningBalance;
      updatedCount++;
    }

    // Write all updated balances in one go
    const balanceRange = sheet.getRange(startRow, COLUMNS.BALANCE, balances.length, 1);
    balanceRange.setValues(balances.map(v => [v]));
    balanceRange.setNumberFormat('₹#,##0.00');

    // Highlight updated cells
    for (let i = 0; i < balances.length; i++) {
      if (data[i][COLUMNS.BALANCE - 1] !== balances[i]) {
        sheet.getRange(startRow + i, COLUMNS.BALANCE).setBackground('#d4f4dd');
      }
    }

    // ✅ Always update the current account balance in header with the final balance
    sheet.getRange(2, 6).setValue(finalBalance).setNumberFormat('₹#,##0.00');

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to update transactions: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Feature 3: Get Current/Monthly Summary
 * Creates two summary tables on the right side of the current sheet
 */
function getCurrentMonthlySummary() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const startRow = 5; // Data starts from row 5 (after header block)

    // Check if sheet has proper structure
    if (lastRow < 4) {
      SpreadsheetApp.getUi().alert(
        'No Setup',
        'Please run "Setup Columns" first to format the sheet.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Check if there are any transactions
    const transactionCount = Math.max(0, lastRow - startRow + 1);
    if (transactionCount === 0) {
      SpreadsheetApp.getUi().alert(
        'No Transactions',
        'No transactions found. Please add some transactions first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // ✅ First, clear any existing summary tables completely
    clearExistingSummaryTables(sheet);

    // Calculate summary data first
    const summaryData = calculateSummaryData(sheet, lastRow);

    // Position for summary tables - fixed position for better visibility
    const summaryStartCol = 9; // Fixed column I (9th column)
    const summaryStartRow = 8; // Fixed position - always start at row 8

    // ✅ Ensure previous operations are committed
    SpreadsheetApp.flush();

    // Create Monthly Expenses Summary
    createMonthlyExpensesSummary(sheet, summaryStartCol, summaryStartRow, summaryData);

    // Create Account Status Summary
    createAccountStatusSummary(sheet, summaryStartCol, summaryStartRow, summaryData);

    // ✅ Commit all formatting/writes before showing alert
    SpreadsheetApp.flush();

    // ✅ Alert shown only after everything is ready
    SpreadsheetApp.getUi().alert(
      'Success!',
      `Monthly summary generated successfully!\n(${transactionCount} transactions processed).`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    console.error('Summary generation error:', error);
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to generate summary: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Clears any existing summary tables completely
 */
function clearExistingSummaryTables(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    // Clear a large area in columns I and J (9-10) to ensure we catch any existing summary tables
    // Start from row 1 and go down to the last row + 50 to be safe
    const clearRange = sheet.getRange(1, 9, Math.max(lastRow + 50, 100), 2);
    
    // Clear everything comprehensively
    clearRange.clearContent(); // Clear all content
    clearRange.clearFormat(); // Clear all formatting
    clearRange.clearNote(); // Clear notes/comments
    clearRange.setBorder(false, false, false, false, false, false); // Remove all borders
    clearRange.setBackground(null); // Remove background colors
    clearRange.setFontWeight('normal'); // Reset font weight
    clearRange.setFontColor('black'); // Reset font color
    clearRange.setHorizontalAlignment('left'); // Reset alignment
    
    // Also clear any merged cells in the summary area
    try {
      clearRange.breakApart(); // Break any merged cells
    } catch (e) {
      // Ignore errors if no merged cells exist
    }
    
    SpreadsheetApp.flush(); // Commit the clearing operations
    
  } catch (error) {
    console.error('Error clearing existing summary tables:', error);
    // Don't throw error - continue with summary generation even if clearing fails
  }
}

/**
 * Calculates summary data from transaction records
 */
function calculateSummaryData(sheet, lastRow) {
  const startRow = 5; // Data starts from row 5 (after header block)

  if (lastRow < startRow) {
    return {
      totalIncome: 0,
      totalExpense: 0,
      savings: 0,
      startBalance: 0,
      endBalance: 0,
      currentSavings: 0,
      currentAccountBalance: 0
    };
  }

  try {
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, HEADERS.length).getValues();

    let totalIncome = 0;  // In pocket + Gift
    let totalExpense = 0;
    let startBalance = 0;
    let endBalance = 0;

    // Get starting balance from header block (row 2, column 4)
    try {
      startBalance = Number(sheet.getRange(2, 4).getValue()) || 0;
    } catch (e) {
      console.warn('Could not get start balance from header:', e);
      startBalance = 0;
    }

    // Process each transaction row
    data.forEach((row, index) => {
      const amount = Number(row[COLUMNS.AMOUNT - 1]) || 0;
      const type = row[COLUMNS.TYPE - 1];
      const balance = Number(row[COLUMNS.BALANCE - 1]) || 0;

      // Only process rows with valid amount and type
      if (amount > 0 && type && ['In pocket', 'Gift', 'Expense'].includes(type)) {
        if (type === 'In pocket' || type === 'Gift') {
          totalIncome += amount;
        } else if (type === 'Expense') {
          totalExpense += amount;
        }
      }

      // Get end balance from last transaction with valid balance
      if (balance > 0 && index === data.length - 1) {
        endBalance = balance;
      }
    });

    // If no valid end balance found, calculate it
    if (endBalance === 0) {
      endBalance = startBalance + totalIncome - totalExpense;
    }

    const savings = totalIncome - totalExpense;
    const currentAccountBalance = startBalance + savings; // Starting balance + current month net

    return {
      totalIncome,
      totalExpense,
      savings,
      startBalance,
      endBalance,
      currentSavings: savings,
      currentAccountBalance: currentAccountBalance
    };

  } catch (error) {
    console.error('Error calculating summary data:', error);
    return {
      totalIncome: 0,
      totalExpense: 0,
      savings: 0,
      startBalance: 0,
      endBalance: 0,
      currentSavings: 0,
      currentAccountBalance: 0
    };
  }
}

/**
 * Creates the Monthly Expenses Summary table
 */
function createMonthlyExpensesSummary(sheet, startCol, startRow, summaryData) {
  try {
    // Title
    const titleRange = sheet.getRange(startRow, startCol, 1, 2);
    titleRange.merge().setValue('💰 Monthly Expenses');
    titleRange.setFontWeight('bold').setFontSize(12).setBackground('#28a745').setFontColor('white').setHorizontalAlignment('center');

    // Summary data with color coding
    const summaryTable = [
      ['Income:', summaryData.totalIncome],
      ['Spent:', summaryData.totalExpense],
      ['Left/Savings:', summaryData.savings]
    ];

    const summaryRange = sheet.getRange(startRow + 1, startCol, summaryTable.length, 2);
    summaryRange.setValues(summaryTable);
    summaryRange.setBorder(true, true, true, true, true, true);

    // Apply different background colors for each row
    sheet.getRange(startRow + 1, startCol, 1, 2).setBackground('#d4f4dd'); // Light green for income
    sheet.getRange(startRow + 2, startCol, 1, 2).setBackground('#ffe6e6'); // Light red for expenses  
    sheet.getRange(startRow + 3, startCol, 1, 2).setBackground('#fff3cd'); // Light yellow for savings

    // Make labels bold and centered
    sheet.getRange(startRow + 1, startCol, summaryTable.length, 1).setFontWeight('bold').setHorizontalAlignment('center');

    // Format amounts and center them
    sheet.getRange(startRow + 1, startCol + 1, summaryTable.length, 1).setNumberFormat('₹#,##0.00').setFontWeight('bold').setHorizontalAlignment('center');

    // Add notes
    sheet.getRange(startRow + 1, startCol).setNote('Sum of all In pocket and Gift transactions');
    sheet.getRange(startRow + 2, startCol).setNote('Sum of all Expense transactions');
    sheet.getRange(startRow + 3, startCol).setNote('Income minus Expenses');

    // Set column widths for better layout
    sheet.setColumnWidth(startCol, 220);      // Label column
    sheet.setColumnWidth(startCol + 1, 140);  // Amount column

  } catch (error) {
    console.error('Error creating monthly expenses summary:', error);
    throw error;
  }
}

/**
 * Creates the Account Status Summary table
 */
function createAccountStatusSummary(sheet, startCol, startRow, summaryData) {
  try {
    const accountStartRow = startRow + 6; // Start below the first summary with some spacing

    // Title
    const titleRange = sheet.getRange(accountStartRow, startCol, 1, 2);
    titleRange.merge().setValue('📊 Account Status');
    titleRange.setFontWeight('bold').setFontSize(12).setBackground('#6f42c1').setFontColor('white').setHorizontalAlignment('center');

    // Status data with different colors for different balance types
    const statusTable = [
      ['Starting Balance:', summaryData.startBalance],
      ['Current Month Net:', summaryData.savings],
      ['Current Account Balance:', summaryData.currentAccountBalance]
    ];

    const statusRange = sheet.getRange(accountStartRow + 1, startCol, statusTable.length, 2);
    statusRange.setValues(statusTable);
    statusRange.setBorder(true, true, true, true, true, true);

    sheet.setColumnWidth(startCol, 220);      // Label column
    sheet.setColumnWidth(startCol + 1, 140);  // Amount column

    // Apply different background colors for each balance row
    sheet.getRange(accountStartRow + 1, startCol, 1, 2).setBackground('#cce5ff'); // Light blue for start balance
    sheet.getRange(accountStartRow + 2, startCol, 1, 2).setBackground('#fff3cd'); // Light yellow for current month net
    sheet.getRange(accountStartRow + 3, startCol, 1, 2).setBackground('#d4f4dd'); // Light green for current account balance

    // Make labels bold and centered
    sheet.getRange(accountStartRow + 1, startCol, statusTable.length, 1).setFontWeight('bold').setHorizontalAlignment('center');

    // Format amounts and center them
    sheet.getRange(accountStartRow + 1, startCol + 1, statusTable.length, 1).setNumberFormat('₹#,##0.00').setFontWeight('bold').setHorizontalAlignment('center');

    // Add notes
    sheet.getRange(accountStartRow + 1, startCol).setNote('Balance at the beginning of this month');
    sheet.getRange(accountStartRow + 2, startCol).setNote('Income minus expenses for current month only');
    sheet.getRange(accountStartRow + 3, startCol).setNote('Current total account balance (Starting + Current Month Net)');

  } catch (error) {
    console.error('Error creating account status summary:', error);
    throw error;
  }
}

/**
 * Feature 4: Create Next Month Sheet
 * Creates a new blank sheet with next month name
 */
function generateNextMonthSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentSheet = ss.getActiveSheet();
    const currentSheetName = currentSheet.getName();

    // Generate next month name
    const nextMonthName = generateNextMonthName(currentSheetName);

    // Check if next month sheet already exists
    if (ss.getSheetByName(nextMonthName)) {
      SpreadsheetApp.getUi().alert('Sheet Exists', `Sheet "${nextMonthName}" already exists.`, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Create new blank sheet with next month name
    const newSheet = ss.insertSheet(nextMonthName);

    // Move new sheet to end
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(ss.getNumSheets());

    SpreadsheetApp.getUi().alert('Success!', `New sheet "${nextMonthName}" created successfully!\n\nPlease click "Setup Columns" to configure the sheet and enter your monthly salary and starting balance.`, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to create next month sheet: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Gets the current month name in "Month Year" format
 */
function getCurrentMonthName() {
  const now = new Date();
  const months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  return months[now.getMonth()] + ' ' + now.getFullYear();
}

/**
 * Generates the next month name from current month name
 */
function generateNextMonthName(currentName) {
  const months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  // Try to parse current name (format: "Month Year" or just use current date)
  let currentMonth, currentYear;

  const nameParts = currentName.split(' ');
  if (nameParts.length >= 2 && months.includes(nameParts[0])) {
    currentMonth = months.indexOf(nameParts[0]);
    currentYear = parseInt(nameParts[1]) || new Date().getFullYear();
  } else {
    // Fallback to current date
    const now = new Date();
    currentMonth = now.getMonth();
    currentYear = now.getFullYear();
  }

  // Calculate next month
  let nextMonth = currentMonth + 1;
  let nextYear = currentYear;

  if (nextMonth > 11) {
    nextMonth = 0;
    nextYear++;
  }

  return `${months[nextMonth]} ${nextYear}`;
}

/**
 * Creates a header block with month details
 */
function createMonthHeaderBlock(sheet, monthName, income, startBalance) {
  // Insert rows at the top for header block
  sheet.insertRows(1, 3);

  // Month title - fill all 7 columns with improved color
  sheet.getRange(1, 1, 1, HEADERS.length).merge().setValue(`📊 ${monthName} - Expense Tracker`);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontSize(14).setFontWeight('bold').setBackground('#e8f5e8').setFontColor('#2d5016');

  // Month details row - with different colors for each section
  sheet.getRange(2, 1).setValue('Monthly Income:').setBackground('#fff3cd').setFontWeight('bold');
  sheet.getRange(2, 2).setValue(income).setNumberFormat('₹#,##0.00').setBackground('#fff3cd').setFontWeight('bold');

  sheet.getRange(2, 3).setValue('Initial Balance:').setBackground('#cce5ff').setFontWeight('bold');
  sheet.getRange(2, 4).setValue(startBalance).setNumberFormat('₹#,##0.00').setBackground('#cce5ff').setFontWeight('bold');

  sheet.getRange(2, 5).setValue('Current Balance:').setBackground('#d4f4dd').setFontWeight('bold');
  sheet.getRange(2, 6).setValue(startBalance).setNumberFormat('₹#,##0.00').setBackground('#d4f4dd').setFontWeight('bold');

  // Fill remaining cell in row 2 with light background
  sheet.getRange(2, 7, 1, 1).setBackground('#f8f9fa');

  // Add empty row for spacing
  sheet.getRange(3, 1, 1, HEADERS.length).setBackground('#f8f9fa');

  // Column headers with improved color coordination
  sheet.getRange(4, 1, 1, HEADERS.length).setValues([HEADERS]);

  // Apply different colors to specific header columns
  sheet.getRange(4, COLUMNS.DATE).setBackground('#6c757d').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center'); // Grey for date
  sheet.getRange(4, COLUMNS.CATEGORY).setBackground('#6c757d').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center'); // Grey for category
  sheet.getRange(4, COLUMNS.DESCRIPTION).setBackground('#ffc107').setFontColor('black').setFontWeight('bold').setHorizontalAlignment('center'); // Yellow for description/reference
  sheet.getRange(4, COLUMNS.AMOUNT).setBackground('#6c757d').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center'); // Grey for amount
  sheet.getRange(4, COLUMNS.TYPE).setBackground('#6c757d').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center'); // Grey for type
  sheet.getRange(4, COLUMNS.MODE).setBackground('#6c757d').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center'); // Grey for mode
  sheet.getRange(4, COLUMNS.BALANCE).setBackground('#28a745').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center'); // Green for balance

  // Set column widths for better visibility
  sheet.setColumnWidth(COLUMNS.DATE, 110);
  sheet.setColumnWidth(COLUMNS.CATEGORY, 140);
  sheet.setColumnWidth(COLUMNS.DESCRIPTION, 250);
  sheet.setColumnWidth(COLUMNS.AMOUNT, 120);
  sheet.setColumnWidth(COLUMNS.TYPE, 160);
  sheet.setColumnWidth(COLUMNS.MODE, 120);
  sheet.setColumnWidth(COLUMNS.BALANCE, 130);

  // Update frozen rows
  sheet.setFrozenRows(4);
}

function clearResults() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    const range = sheet.getActiveRange();
    
    if (!range) {
      ui.alert('No Selection', 'Please select the summary tables you want to clear.', ui.ButtonSet.OK);
      return;
    }

    // Clear everything comprehensively
    range.breakApart(); // In case there are merged header cells
    range.clearContent(); // Clear all content
    range.clearFormat(); // Clear all formatting
    range.clearNote(); // Clear notes/comments
    range.clearDataValidations(); // Clear dropdown menus
    range.setBackground(null); // Remove background colors
    range.setBorder(false, false, false, false, false, false); // Remove all borders
    range.setFontWeight('normal'); // Reset font weight
    range.setFontColor('black'); // Reset font color
    range.setHorizontalAlignment('left'); // Reset alignment
    SpreadsheetApp.flush();
    
    ui.alert('Success!', 'Selected range cleared completely!', ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('Error', 'Failed to clear selected area: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Debug function to help troubleshoot summary issues
 */
function debugSummaryData() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const startRow = 5;

    console.log('Debug Info:');
    console.log('Last Row:', lastRow);
    console.log('Last Column:', lastCol);
    console.log('Start Row:', startRow);

    if (lastRow >= startRow) {
      const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, HEADERS.length).getValues();
      console.log('Transaction Data:', data);
      
      const summaryData = calculateSummaryData(sheet, lastRow);
      console.log('Summary Data:', summaryData);
    }

    // Check header block
    try {
      const monthlyIncome = sheet.getRange(2, 2).getValue();
      const startBalance = sheet.getRange(2, 4).getValue();
      console.log('Monthly Income:', monthlyIncome);
      console.log('Start Balance:', startBalance);
    } catch (e) {
      console.log('Header block error:', e);
    }

  } catch (error) {
    console.error('Debug error:', error);
  }
}