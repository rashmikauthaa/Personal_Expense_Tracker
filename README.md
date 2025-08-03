Create a Google Apps Script for a Google Sheet-based Personal Expense Tracker with the following requirements:

### 1️⃣ General Overview
- Add a **custom menu** called `Expense Tracker` with the following items:
  1. Setup Columns
  2. Update Transactions
  3. Get Current/Monthly Summary
  4. Create Next Month Sheet

- The script should work in Google Sheets and handle user interaction via dialogs, prompts, and Sheets automation.
- All code should be modular, well-commented, and visually formatted where needed.

---

### 2️⃣ Feature 1: Setup Columns
When the user clicks **Setup Columns**, the script should:

1. Format the active sheet into a structured table with these columns:

   | Column        | Details                                                      |
   |---------------|--------------------------------------------------------------|
   | Date          | Date of transaction                                          |
   | Category      | Dropdown with a few default categories (Food, Rent, Salary)  |
   | Description   | Free text for personal notes                                 |
   | Type          | Dropdown: `In pocket`, `Expense`, `Gift`                     |
   | Mode          | Dropdown: `Cash`, `UPI`, `Card`, `Bank`                      |
   | Balance       | To be calculated by the script                               |

2. Apply **data validation** for dropdown columns:
   - **Category**: A default list (Food, Rent, Travel, Salary, Misc)
   - **Type**: `In pocket`, `Expense`, `Gift`
   - **Mode**: `Cash`, `UPI`, `Card`, `Bank`

3. Apply **basic formatting**:
   - Header row: Bold, center aligned, colored background (light blue or green)
   - Alternate row shading
   - Auto-resize columns
   - Notes on hover for columns:
     - Description → “Enter details for your reference”
     - Balance → “Auto-calculated after Update Transactions”

---

### 3️⃣ Feature 2: Update Transactions
When the user clicks **Update Transactions**, the script should:

1. Assume the user has entered all transactions but **left the Balance column empty**.
2. Calculate the running **Balance** using the following logic:
   - Take the **previous row balance** as the base.
   - For **Type**:
     - `In pocket` → Add the amount
     - `Expense` → Subtract the amount
     - `Gift` → Add the amount
3. Update all empty balance cells automatically.
4. Visually highlight the updated cells (light green).
5. Should work for multiple new rows at once.

---

### 4️⃣ Feature 3: Get Current/Monthly Summary
When the user clicks **Get Current/Monthly Summary**, the script should:

1. Generate **two small summary tables** at the right side of the current sheet (starting 2 columns after the last column):

#### a) Monthly Expenses Summary
Monthly Expenses
Income: <sum of 'In pocket'>
Spent: <sum of 'Expense'>
Left/Savings: <Income - Expense>



#### b) Account Status Summary
Account Status
Start Balance (previous cumulative):
Savings (current month):
Total Balance / To Be Next Month Balance: <calculated>

2. Apply **borders, background colors, and bold titles**.
3. Add hover notes explaining each field.

---

### 4️⃣ Feature: Create Next Month Sheet

**Behavior:**
1. Duplicates the **current sheet structure** (headers, dropdowns, formatting).
2. Clears old transactions, retains only **header and formatting**.
3. **Sheet name auto-generated**:
   - If current sheet is `July 2025`, next = `August 2025`.
   - Handle year rollover (December → January next year).
4. **Popup prompt** again asks:
   - “This Month Income / Stipend”
   - “Current Account Balance” (pre-fill with last month closing balance if available)
5. Writes header block at the top with new month details.

---

### 5️⃣ Additional Notes

- Use `onOpen()` to automatically create the menu.
- Use **SpreadsheetApp** methods for styling, data validation, and cell notes.
- Maintain **cumulative account balance logic**:
  - First month: user enters both manually
  - Next months: starting balance = previous month closing balance + new month income
- Ensure **robust error handling** for empty or invalid rows.

---

### ✅ End Goal
Deliver a single Apps Script file that:
- Adds the `Expense Tracker` menu
- Handles setup, balance calculation, summaries, and month-to-month tracking
- Is visually clean, user-friendly, and optimized for Google Sheets

---

## 🎉 **IMPLEMENTATION COMPLETE!**

The Google Apps Script has been successfully built according to all specifications above. 

### 📋 **Quick Setup Guide:**

1. **Open Google Sheets** → Create a new spreadsheet
2. **Go to Extensions** → Apps Script  
3. **Delete default code** → Paste the code from `code.gs`
4. **Save the script** (Ctrl+S)
5. **Refresh your Google Sheet** → You'll see the "Expense Tracker" menu

### 🚀 **Usage Workflow:**

1. **First Time**: Click `Setup Columns` → Enter monthly salary & account balance
2. **Add Transactions**: Enter data manually using dropdown menus (starts from row 5)
3. **Calculate Balances**: Click `Update Balances` → Auto-calculates running balance
4. **View Summary**: Click `Generate Monthly Summary` → Creates summary tables on right side
5. **Clear Results**: Select any range → Click `Clear Results` → Removes everything
6. **Next Month**: Click `Generate Next Month Sheet` → Creates blank sheet for next month

### 📊 **Transaction Types:**
- **In pocket** → Adds to balance (salary, income)
- **Expense** → Subtracts from balance (spending)  
- **Gift** → Adds to balance (received money)

### 🎨 **Visual Features:**
- **Color-coded headers**: Date/Category/Amount/Type/Mode (grey), Description (yellow), Balance (green)
- **Summary tables**: Color-coded rows with different backgrounds for Income/Expense/Savings
- **Hover notes**: Helpful tips on Date, Description, and Balance columns
- **Auto-renaming**: Sheets automatically named to current month (e.g., "January 2025")

### 📋 **Updated Categories:**
Food, Rent, Travel, Family, Misc, Groceries

See `sample-config.json` for comprehensive setup instructions and troubleshooting tips.# Personal_Expense_Tracker
