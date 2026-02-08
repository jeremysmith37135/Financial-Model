# CFO Financial Model Generator
## User Manual v3.2.13

---

# Table of Contents

1. [Introduction](#1-introduction)
2. [System Requirements](#2-system-requirements)
3. [Getting Started](#3-getting-started)
4. [Main Interface](#4-main-interface)
5. [Single Entity Mode](#5-single-entity-mode)
6. [Multi-Division Mode](#6-multi-division-mode)
7. [Update Existing Model](#7-update-existing-model)
8. [Generated Excel Workbook](#8-generated-excel-workbook)
9. [Working with the Output](#9-working-with-the-output)
10. [Troubleshooting](#10-troubleshooting)
11. [FAQ](#11-faq)

---

# 1. Introduction

## What is the CFO Financial Model Generator?

The CFO Financial Model Generator is a desktop application designed for CFOs, controllers, and financial professionals to create comprehensive financial models from QuickBooks or other accounting system exports. The application transforms your P&L (Profit & Loss) and Balance Sheet data into a fully-functional Excel workbook with:

- **Interactive P&L and Balance Sheet reports** with YTD calculations
- **Cash Flow Statement** automatically generated from P&L and Balance Sheet data
- **Executive Dashboard** with key financial metrics and charts
- **Forecast module** comparing Actuals vs Budget with variance analysis
- **Multi-division consolidation** for companies with multiple entities

## Key Features

| Feature | Description |
|---------|-------------|
| **Auto-Detection** | Automatically detects date ranges and account structures from your files |
| **File Validation** | Validates that P&L and Balance Sheet files are correctly assigned |
| **Dynamic Formulas** | All reports use formulas - change source data and reports update automatically |
| **Period Selection** | Select any period from the dropdown to view historical data |
| **Multi-Division** | Consolidate multiple divisions/departments/branches into one model |
| **Update Capability** | Add new months to existing models without rebuilding |

---

# 2. System Requirements

## Minimum Requirements

- **Operating System**: Windows 10 or Windows 11
- **Microsoft Excel**: Microsoft Excel 2016 or later (required)
- **Memory**: 4 GB RAM minimum (8 GB recommended)
- **Disk Space**: 200 MB for application + space for generated files
- **Screen Resolution**: 1280 x 720 minimum

## Important Notes

- **Excel must be installed** - The application uses Excel's COM interface to create macro-enabled workbooks
- **Close Excel before running** - For best results, close any open Excel workbooks before generating
- **Enable macros** - Generated workbooks contain VBA macros for navigation features

---

# 3. Getting Started

## Installation

1. Download `CFO_Financial_Model_Generator_v3.2.13.exe`
2. Place the EXE file in a permanent location (e.g., `C:\Program Files\CFO Tools\`)
3. Create a desktop shortcut if desired
4. Double-click to run - no installation required

## Preparing Your Data

### P&L (Profit & Loss) Export

Export your P&L from QuickBooks or your accounting system with:
- **Columns**: Account names in column A, months as column headers
- **Rows**: All income and expense accounts
- **Format**: Excel file (.xlsx or .xls)

Example P&L structure:
```
|          | Oct 2024 | Nov 2024 | Dec 2024 |
|----------|----------|----------|----------|
| Income   |          |          |          |
|  Sales   | 50,000   | 55,000   | 60,000   |
| Expenses |          |          |          |
|  Rent    | 5,000    | 5,000    | 5,000    |
```

### Balance Sheet Export

Export your Balance Sheet with:
- **Columns**: Account names in column A, months as column headers
- **Rows**: All asset, liability, and equity accounts
- **Format**: Excel file (.xlsx or .xls)

Example Balance Sheet structure:
```
|                    | Oct 2024 | Nov 2024 | Dec 2024 |
|--------------------|----------|----------|----------|
| Assets             |          |          |          |
|  Cash              | 100,000  | 120,000  | 150,000  |
| Liabilities        |          |          |          |
|  Accounts Payable  | 20,000   | 25,000   | 22,000   |
```

---

# 4. Main Interface

When you launch the application, you'll see the main window:

```
+--------------------------------------------------+
|        CFO Financial Model Generator              |
|              Version 3.2.13                       |
+--------------------------------------------------+
|                                                   |
|  Mode: [Single Entity]                            |
|        [Configure Divisions...]                   |
|                                                   |
|  Company Name: [________________________]         |
|                                                   |
|  -- OR Update Existing Model --                   |
|  Model File: [________________________] [Browse]  |
|                                                   |
|  P&L File:       [____________________] [Browse]  |
|  Balance Sheet:  [____________________] [Browse]  |
|                                                   |
|  Date range will be auto-detected from files.     |
|                                                   |
|        [Generate Financial Model]                 |
|                                                   |
|  Status: Ready                                    |
|  [========================] 0%                    |
|                                                   |
|                              [Close]              |
+--------------------------------------------------+
```

## Interface Elements

| Element | Description |
|---------|-------------|
| **Mode** | Shows "Single Entity" or "X Divisions configured" |
| **Configure Divisions** | Opens the multi-division setup dialog |
| **Company Name** | Enter the company name for report headers |
| **Model File** | (Optional) Select existing model to update |
| **P&L File** | Select your P&L export file |
| **Balance Sheet** | Select your Balance Sheet export file |
| **Generate** | Creates the financial model |
| **Progress Bar** | Shows generation progress with time estimate |
| **Status** | Displays current operation or errors |

---

# 5. Single Entity Mode

Single Entity mode is for companies with one set of books - no divisions, departments, or branches to consolidate.

## Step-by-Step Guide

### Step 1: Enter Company Name
Type your company name in the "Company Name" field. This appears on all report headers.

### Step 2: Select P&L File
1. Click **Browse** next to "P&L File"
2. Navigate to your P&L export
3. Select the file and click **Open**
4. The application validates it's a P&L file (looks for income/expense keywords)

### Step 3: Select Balance Sheet File
1. Click **Browse** next to "Balance Sheet"
2. Navigate to your Balance Sheet export
3. Select the file and click **Open**
4. The application validates it's a Balance Sheet (looks for assets/liabilities keywords)

### Step 4: Generate Model
1. Click **Generate Financial Model**
2. Choose where to save the output file
3. Wait for generation to complete (typically 30-90 seconds)
4. The generated file opens automatically in Excel

## What Gets Created

The generated workbook contains these sheets:

| Sheet | Description |
|-------|-------------|
| **Menu** | Navigation hub with company info and period selector |
| **Dashboard** | Executive summary with KPIs and charts |
| **PL** | Profit & Loss statement with YTD columns |
| **Balance_Sheet** | Balance Sheet with comparative periods |
| **Cash_Flow** | Statement of Cash Flows (auto-generated) |
| **Forecast** | 12-month forecast with Actual vs Budget |
| **Forecast_Summary** | Condensed forecast view |
| **Notes** | Annotation sheet for comments |
| **Settings** | Sheet visibility controls |
| **Source_PL** | Raw P&L data (hidden) |
| **Source_BS** | Raw Balance Sheet data (hidden) |
| **Source_Budget** | Budget data template (hidden) |

---

# 6. Multi-Division Mode

Multi-Division mode is for companies with multiple entities, departments, branches, or divisions that need consolidated reporting.

## When to Use Multi-Division Mode

- Parent company with subsidiaries
- Company with multiple departments
- Franchise with multiple locations
- Any organization needing consolidated financials

## Setting Up Divisions

### Step 1: Open Division Setup
Click **Configure Divisions...** to open the Division Setup dialog.

### Step 2: Add Divisions
1. Click **+ Add Division** for each entity
2. Enter a name for each division (e.g., "North Region", "South Region")
3. For each division, browse to select:
   - P&L file for that division
   - Balance Sheet file for that division

### Step 3: Save Configuration
Click **Save & Continue** to return to the main window.

The mode indicator will show "X Divisions configured".

## Generated Output

Multi-division models include additional sheets:

| Sheet | Description |
|-------|-------------|
| **Consolidated_PL** | Combined P&L for all divisions |
| **Division1_PL** | P&L for first division |
| **Division2_PL** | P&L for second division |
| **Consolidated_Forecast** | Combined forecast |
| **Division1_Forecast** | Forecast for first division |

## Division Selector

Reports include a division selector dropdown that lets you switch between:
- **Consolidated** - Shows all divisions combined
- **Division Name** - Shows only that division's data

---

# 7. Update Existing Model

Instead of creating a new model each month, you can update an existing model with new data.

## When to Use Update Mode

- Adding a new month of data to an existing model
- Updating a model with corrected data
- Maintaining a continuous financial model

## How to Update

### Step 1: Select Existing Model
1. Click **Browse** next to "Model File"
2. Select your existing financial model (.xlsm file)

### Step 2: Select New Data Files
1. Select your NEW P&L file (containing the new month)
2. Select your NEW Balance Sheet file (containing the new month)

### Step 3: Generate Update
1. Click **Generate Financial Model**
2. Choose where to save (can overwrite or save as new)
3. Wait for update to complete

## What Gets Updated

- **Source_PL**: New month data added
- **Source_BS**: New month data added
- **All Reports**: New column inserted with formulas
- **Menu**: Dropdown updated with new month
- **Forecast**: Actual formulas updated for new data
- **Dashboard**: Automatically reflects new period

---

# 8. Generated Excel Workbook

## Menu Sheet

The Menu sheet is your navigation hub:

```
+------------------------------------------+
|     COMPANY NAME                         |
|     Financial Model                      |
+------------------------------------------+
|                                          |
|  Current Period: [Dec 2024 ▼]            |
|  Data Range: Oct 2024 - Dec 2024         |
|  Actuals Through: Dec 2024               |
|                                          |
|  Quick Links:                            |
|  • Dashboard                             |
|  • P&L                                   |
|  • Balance Sheet                         |
|  • Cash Flow                             |
|  • Forecast                              |
+------------------------------------------+
```

### Period Selector (C7)
- Click the dropdown to select any available month
- All reports automatically update to show data through that period
- YTD calculations adjust to the selected period

## Dashboard Sheet

The Dashboard provides an executive overview:

### Key Metrics
- Revenue (Current Month and YTD)
- Gross Profit and Margin %
- Operating Expenses
- Net Income and Margin %

### Charts
- Revenue trend over time
- Expense breakdown
- Net Income visualization
- Key ratios

### Division Selector (Multi-Division Only)
- Dropdown in cell C6 to switch between divisions

## P&L Sheet

Profit & Loss statement with:
- Monthly columns for each period
- YTD columns (Prior Year, Current Year)
- Annual totals by year
- Variance columns ($ and %)

### Row Structure
- **Headers**: Bold, section titles
- **Detail rows**: Indented account names
- **Totals**: Bold with borders (Gross Profit, Net Income, etc.)

### Hidden Helper Row
- Row 3 contains YYYYMM values (e.g., 202412) for formula lookups
- This row is hidden but essential for formulas

## Balance Sheet

Balance Sheet with:
- Monthly snapshot columns
- Asset, Liability, and Equity sections
- Totals with accounting-style formatting

## Cash Flow Statement

Automatically generated from P&L and Balance Sheet:

### Operating Activities
- Net Income (from P&L)
- Adjustments for non-cash items
- Changes in working capital

### Investing Activities
- Changes in fixed assets

### Financing Activities
- Changes in debt
- Changes in equity

### Summary
- Net change in cash
- Beginning cash balance
- Ending cash balance

## Forecast Sheet

12-month forecast with 5 columns per month:

| Column | Description |
|--------|-------------|
| **Actual** | Actual data from Source_PL (past months) |
| **Budget** | Budget data from Source_Budget |
| **Adj** | Manual adjustments (user input) |
| **Note** | Comments/explanations |
| **Forecast** | = Actual (past) or Budget + Adj (future) |

### How Forecast Works
- **Past/Current months**: Forecast = Actual
- **Future months**: Forecast = Budget + Adjustment

### Using the Forecast
1. Enter budget data in Source_Budget sheet
2. Add adjustments in Adj column as needed
3. Add notes to explain variances
4. Forecast column automatically calculates

---

# 9. Working with the Output

## Changing the Viewing Period

1. Go to the **Menu** sheet
2. Click the dropdown in cell **C7**
3. Select any available month
4. All sheets automatically update

## Adding Budget Data

1. Go to the **Source_Budget** sheet
2. Enter budget amounts for each account
3. Budget columns are B (Jan) through M (Dec)
4. Forecast sheet automatically uses this data

## Viewing Different Divisions

1. Go to any report sheet (Dashboard, P&L, etc.)
2. Find the division dropdown (usually cell C6)
3. Select "Consolidated" or a specific division

## Printing Reports

Each sheet is pre-formatted for printing:
- Appropriate margins and scaling
- Headers repeat on each page
- Portrait or landscape as appropriate

To print:
1. Go to the desired sheet
2. File > Print
3. Review preview and print

## Protecting Your Work

Recommended workflow:
1. Generate the model
2. Save a backup copy
3. Work in the primary copy
4. Re-generate when needed (update mode)

---

# 10. Troubleshooting

## Common Issues

### "The file is still open"
**Solution**: Close Excel completely and try again.

### "Could not determine date columns"
**Solution**:
- Ensure your Excel file has date headers (e.g., "Oct 2024", "Nov 2024")
- Check that dates are in the first few rows
- The application will prompt you to enter dates manually if needed

### "Wrong file type detected"
**Solution**: You may have selected the P&L file for Balance Sheet or vice versa. The application checks for keywords:
- P&L: income, revenue, expense, cost of
- Balance Sheet: assets, liabilities, equity, cash

### Model takes too long to generate
**Typical times**:
- Single entity: 30-90 seconds
- Multi-division (3 divisions): 2-3 minutes
- Multi-division (5 divisions): 3-5 minutes

### Formulas show #REF! errors
**Solution**: This usually means the source data structure doesn't match expected format. Regenerate the model.

### Charts don't show data
**Solution**: Ensure the hidden helper columns (row 3 YYYYMM values) are present but hidden (not deleted).

## Getting Help

If you encounter issues:
1. Note the exact error message
2. Note what step you were on
3. Contact Focus CFO support with details

---

# 11. FAQ

**Q: Can I edit the generated reports?**
A: Yes, but be careful with formula cells. Adding rows/columns may break formulas.

**Q: How do I add a new month?**
A: Use Update mode - select your existing model and new data files.

**Q: Can I use this with accounting systems other than QuickBooks?**
A: Yes, as long as your export has account names in column A and months as column headers.

**Q: Why are some sheets hidden?**
A: Source sheets (Source_PL, Source_BS, Source_Budget) are hidden to keep the workbook clean. Use Settings sheet to show/hide sheets.

**Q: Can I customize the reports?**
A: Yes, but save a backup first. The formulas are designed to work together.

**Q: What happens if my account names change?**
A: Regenerate the model with the new data. The formulas look up accounts by name.

**Q: How do I enter my budget?**
A: Go to Source_Budget sheet and enter amounts. The Forecast sheet reads from there.

**Q: Can multiple people use the same model?**
A: Yes, but not simultaneously. Save and close before another user opens.

---

# Support

For questions, issues, or feature requests:
- Email: support@focuscfo.com
- Website: www.focuscfo.com

---

*CFO Financial Model Generator v3.2.13*
*Copyright 2024-2026 Focus CFO. All rights reserved.*
