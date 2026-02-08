# CFO Financial Model Generator
## Beta Testing Guide v3.2.13

---

# Introduction

Thank you for helping test the CFO Financial Model Generator! Your feedback is essential for ensuring a quality product. This guide provides specific test scenarios, expected results, and feedback forms.

**Important**: Please complete all applicable tests and provide honest feedback. Both passes AND failures are valuable information.

---

# Before You Begin

## Test Environment Setup

1. **Windows PC** with Excel 2016 or later installed
2. **Close all Excel windows** before testing
3. **Download test files** from the provided location
4. **Create a test folder** (e.g., `C:\CFO_Testing\`) for generated files

## Test Data Files Needed

You should have received:
- [ ] Sample P&L files (single entity)
- [ ] Sample Balance Sheet files (single entity)
- [ ] Multi-division P&L files (3+ divisions)
- [ ] Multi-division Balance Sheet files (3+ divisions)
- [ ] Previous month's generated model (for update testing)

---

# Section 1: Installation & Launch Tests

## Test 1.1: Application Launch
| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Double-click EXE file | Application window opens | | |
| Window displays correctly | All buttons/labels visible | | |
| Version shows 3.2.13 | Check bottom left corner | | |
| No error messages on launch | Clean startup | | |

## Test 1.2: Interface Responsiveness
| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Window can be moved | Drag title bar | | |
| Buttons respond to hover | Visual feedback | | |
| Browse buttons work | File dialog opens | | |
| Close button works | Application closes cleanly | | |

---

# Section 2: Single Entity Mode Tests

## Test 2.1: File Selection

### P&L File Selection
| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Browse opens file dialog | Dialog appears | | |
| Select P&L file | Path shows in text field | | |
| File validation runs | No error (correct file type) | | |
| Select BS file as P&L | Should show warning | | |

### Balance Sheet Selection
| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Browse opens file dialog | Dialog appears | | |
| Select BS file | Path shows in text field | | |
| File validation runs | No error (correct file type) | | |
| Select P&L file as BS | Should show warning | | |

## Test 2.2: Model Generation

### Test with 3-Month Data (Oct-Dec)
| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Enter company name | Text accepted | | |
| Click Generate | Save dialog appears | | |
| Choose save location | File saves | | |
| Progress bar updates | Shows progress 0-100% | | |
| Progress messages show | Step descriptions visible | | |
| Generation completes | Success message | | |
| File opens in Excel | Workbook opens | | |
| **Time to complete** | Record: ___ seconds | | |

### Generated File Validation
| Sheet | Expected | Pass/Fail | Notes |
|-------|----------|-----------|-------|
| Menu exists | Sheet present | | |
| Dashboard exists | Sheet present | | |
| PL exists | Sheet present | | |
| Balance_Sheet exists | Sheet present | | |
| Cash_Flow exists | Sheet present | | |
| Forecast exists | Sheet present | | |
| Source_PL hidden | Sheet hidden | | |
| Source_BS hidden | Sheet hidden | | |

## Test 2.3: Menu Sheet Functionality

| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Company name displays | Your entered name shows | | |
| Period dropdown works (C7) | Click shows all months | | |
| Select different month | Dropdown changes | | |
| Data range correct (C8) | Shows "Oct 2024 - Dec 2024" | | |
| Quick links work | Clicking navigates to sheet | | |
| Hyperlinks are blue | Standard link color | | |

## Test 2.4: P&L Sheet Validation

| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Row 3 is hidden | YYYYMM row not visible | | |
| Column headers show months | Oct, Nov, Dec visible | | |
| YTD columns present | PY YTD, CY YTD columns | | |
| Data values match source | Spot check 3 accounts | | |
| Total Income row exists | Shows sum | | |
| Gross Profit row exists | Calculated correctly | | |
| Net Income row exists | Shows correct value | | |
| Formulas update with Menu | Change Menu C7, P&L updates | | |
| **Number format** | No cents showing | | |
| **Column widths** | Data fits without clipping | | |

## Test 2.5: Balance Sheet Validation

| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Row 3 is hidden | YYYYMM row not visible | | |
| Assets section present | Correct accounts listed | | |
| Liabilities section present | Correct accounts listed | | |
| Equity section present | Correct accounts listed | | |
| Total Assets = Total L&E | Balanced | | |
| Formulas use YYYYMM lookup | Not static values | | |

## Test 2.6: Cash Flow Validation

| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Row 3 is hidden | YYYYMM row not visible | | |
| Operating section present | Net Income + adjustments | | |
| Investing section present | Asset changes | | |
| Financing section present | Liability/equity changes | | |
| Ending cash = BS Cash | Cross-check with Balance Sheet | | |
| Monthly columns present | Same months as P&L | | |

## Test 2.7: Dashboard Validation

| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Company name shows | Header displays correctly | | |
| Current Period shows correctly | Matches Menu C7 | | |
| **NOT showing number** | Should show "Dec 2024" not "46381" | | |
| Revenue displays | Current month and YTD | | |
| Net Income displays | Current month and YTD | | |
| Charts render | Visual charts visible | | |
| Data updates with Menu | Change Menu C7, Dashboard updates | | |

## Test 2.8: Forecast Validation

| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| 12 months displayed | Jan through Dec | | |
| 5 columns per month | Actual, Budget, Adj, Note, Forecast | | |
| Past months show Actual data | Oct, Nov, Dec have values | | |
| Future months show 0 in Actual | Jan-Sep show 0 | | |
| Future Forecast = Budget + Adj | Formula correct | | |
| Past Forecast = Actual | Formula correct | | |
| Row groups collapsible | Click +/- to expand/collapse | | |

---

# Section 3: Update Existing Model Tests

## Test 3.1: Basic Update

Use a previously generated model and new data files with one additional month.

### Setup
- Existing model: Oct-Nov data
- New files: Oct-Dec data (adds December)

| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Select existing model | File loads | | |
| Select new P&L file | File validates | | |
| Select new BS file | File validates | | |
| Click Generate | Update starts | | |
| Progress shows update steps | Different from new build | | |
| Update completes | Success message | | |
| **Time to complete** | Record: ___ seconds | | |

### Updated File Validation
| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| December column added to PL | New column present | | |
| December column added to BS | New column present | | |
| December column added to CF | New column present | | |
| Menu dropdown has December | C7 dropdown includes Dec | | |
| Menu shows "Dec 2024" | Current period updated | | |
| Data range updated | Shows "Oct 2024 - Dec 2024" | | |
| Dashboard shows December | Current period correct | | |
| Forecast Dec has Actual data | Not 0 anymore | | |
| Row 3 still hidden | YYYYMM helper row | | |
| Existing data unchanged | Oct, Nov data intact | | |

## Test 3.2: Overwrite Same Month

Test updating with same data (should overwrite, not duplicate).

| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Update with same month data | No duplicate columns | | |
| Values may change | If source data changed | | |
| No #REF errors | Formulas intact | | |

---

# Section 4: Multi-Division Mode Tests

## Test 4.1: Division Setup

| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Click Configure Divisions | Dialog opens | | |
| Add Division button works | New division row appears | | |
| Can add 3+ divisions | No limit hit | | |
| Can name each division | Text entry works | | |
| Can browse P&L per division | File dialog works | | |
| Can browse BS per division | File dialog works | | |
| Remove button works | Division row removed | | |
| Save & Continue closes dialog | Returns to main | | |
| Status shows "X Divisions" | Count is correct | | |

## Test 4.2: Multi-Division Generation

| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Generate with 3 divisions | Process starts | | |
| Progress bar updates | Shows division progress | | |
| **Time to complete** | Record: ___ seconds | | |
| File generates successfully | No errors | | |

### Generated Sheets
| Sheet | Expected | Pass/Fail | Notes |
|-------|----------|-----------|-------|
| Consolidated_PL exists | Combined P&L | | |
| Division1_PL exists | First division | | |
| Division2_PL exists | Second division | | |
| Division3_PL exists | Third division | | |
| Consolidated_Forecast exists | Combined forecast | | |
| Dashboard exists | With division selector | | |

## Test 4.3: Division Selector

| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Dashboard has dropdown | Cell C6 | | |
| Dropdown lists Consolidated | First option | | |
| Dropdown lists all divisions | All names present | | |
| Selecting division updates data | Values change | | |
| Charts update with selection | Visuals change | | |
| P&L selector works | If present | | |

## Test 4.4: Consolidation Accuracy

| Item | Expected | Pass/Fail | Notes |
|------|----------|-----------|-------|
| Consolidated Revenue = Sum | Add all division revenues | | |
| Consolidated Expenses = Sum | Add all division expenses | | |
| Account names consolidated | Matching accounts combined | | |
| Division-specific accounts shown | With "(Division)" suffix | | |

---

# Section 5: Edge Cases & Error Handling

## Test 5.1: Missing Data

| Scenario | Expected | Pass/Fail | Notes |
|----------|----------|-----------|-------|
| Generate with empty company name | Should work (blank header) | | |
| P&L file with missing months | Handles gracefully | | |
| BS file with different date range | Warning or adapts | | |

## Test 5.2: File Format Issues

| Scenario | Expected | Pass/Fail | Notes |
|----------|----------|-----------|-------|
| .xls file (old format) | Should work | | |
| .xlsx file | Should work | | |
| File with spaces in path | Should work | | |
| File with special characters | Should work | | |

## Test 5.3: Error Recovery

| Scenario | Expected | Pass/Fail | Notes |
|----------|----------|-----------|-------|
| Cancel during generation | Application remains stable | | |
| Excel opens during generation | Handles or waits | | |
| Disk full error | Clear error message | | |
| Invalid file selected | Clear error message | | |

---

# Section 6: Performance Tests

## Test 6.1: Generation Times

Record actual times for your system:

| Scenario | Expected | Actual | Pass/Fail |
|----------|----------|--------|-----------|
| Single entity (3 months) | 30-90 sec | ___ sec | |
| Single entity (12 months) | 45-120 sec | ___ sec | |
| Multi-division (3 div, 3 months) | 2-3 min | ___ min | |
| Multi-division (5 div, 3 months) | 3-5 min | ___ min | |
| Update existing model | 15-45 sec | ___ sec | |

## Test 6.2: Large Data Sets

| Scenario | Expected | Pass/Fail | Notes |
|----------|----------|-----------|-------|
| 100+ accounts | Should complete | | |
| 24 months of data | Should complete | | |
| 5 divisions | Should complete | | |
| Very long account names | Should display | | |

---

# Section 7: Feedback Questions

Please answer these questions after completing the tests:

## Usability

1. **Was the interface intuitive?**
   - [ ] Very intuitive
   - [ ] Somewhat intuitive
   - [ ] Confusing
   - [ ] Very confusing

2. **Were error messages helpful?**
   - [ ] Very helpful
   - [ ] Somewhat helpful
   - [ ] Not helpful
   - [ ] Did not encounter errors

3. **Was the progress indicator useful?**
   - [ ] Very useful
   - [ ] Somewhat useful
   - [ ] Not useful
   - Comments: _______________________

## Output Quality

4. **Did the generated reports meet your expectations?**
   - [ ] Exceeded expectations
   - [ ] Met expectations
   - [ ] Below expectations
   - Comments: _______________________

5. **Were the formulas correct?**
   - [ ] All correct
   - [ ] Minor issues
   - [ ] Significant issues
   - Describe issues: _______________________

6. **Was the formatting professional?**
   - [ ] Very professional
   - [ ] Acceptable
   - [ ] Needs improvement
   - Suggestions: _______________________

## Features

7. **What features were most valuable?**
   ___________________________________

8. **What features were missing?**
   ___________________________________

9. **What would you change?**
   ___________________________________

## Performance

10. **Was generation time acceptable?**
    - [ ] Fast enough
    - [ ] Acceptable
    - [ ] Too slow
    - Comments: _______________________

## Overall

11. **Would you use this tool in production?**
    - [ ] Yes, definitely
    - [ ] Yes, with improvements
    - [ ] Not sure
    - [ ] No
    - Why: _______________________

12. **Would you recommend this tool?**
    - [ ] Yes
    - [ ] Maybe
    - [ ] No
    - Why: _______________________

---

# Section 8: Bug Report Template

If you encounter any bugs, please copy and complete this template:

```
BUG REPORT
===========

Date:
Tester Name:
Version: 3.2.13

DESCRIPTION:
[What happened?]

STEPS TO REPRODUCE:
1.
2.
3.

EXPECTED BEHAVIOR:
[What should have happened?]

ACTUAL BEHAVIOR:
[What actually happened?]

ERROR MESSAGE (if any):
[Copy exact text]

SCREENSHOT:
[Attach if possible]

DATA FILES:
[Names of files used]

SEVERITY:
[ ] Critical - Cannot use application
[ ] Major - Feature doesn't work
[ ] Minor - Works but not ideal
[ ] Cosmetic - Visual issue only

WORKAROUND (if found):
[Did you find a way around it?]
```

---

# Section 9: Test Sign-Off

## Tester Information

- **Tester Name**: _______________________
- **Date Tested**: _______________________
- **Version Tested**: 3.2.13
- **Operating System**: _______________________
- **Excel Version**: _______________________

## Test Summary

| Section | Tests Passed | Tests Failed | Tests Skipped |
|---------|--------------|--------------|---------------|
| 1. Installation | /4 | | |
| 2. Single Entity | /8 | | |
| 3. Update Model | /2 | | |
| 4. Multi-Division | /4 | | |
| 5. Edge Cases | /3 | | |
| 6. Performance | /2 | | |

## Final Comments

_________________________________________
_________________________________________
_________________________________________

## Sign-Off

- [ ] I have completed all applicable tests
- [ ] I have documented all failures
- [ ] I have provided feedback

**Signature**: _______________________
**Date**: _______________________

---

*Please return completed test results to: [Your Email/Location]*

*Thank you for your valuable feedback!*
