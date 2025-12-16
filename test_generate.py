"""
Test script to generate DNA model with xlwings
"""
import sys
sys.path.insert(0, 'webapp')

import pandas as pd
import xlwings as xw
from desktop_app import DNAModelApp

MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December']

def parse_financial_data(df):
    header_row = None
    for idx in range(min(15, len(df))):
        for col_idx in range(len(df.columns)):
            val = df.iloc[idx, col_idx]
            if pd.notna(val):
                val_str = str(val).strip().lower()
                for month in MONTHS:
                    if val_str.startswith(month.lower()) and '-' not in val_str:
                        parts = str(df.iloc[idx, col_idx]).split()
                        if len(parts) == 2:
                            header_row = idx
                            break
            if header_row is not None:
                break
        if header_row is not None:
            break

    if header_row is None:
        raise ValueError('No header row found')

    months = []
    for col_idx in range(1, len(df.columns)):
        val = df.iloc[header_row, col_idx]
        if pd.notna(val) and 'total' not in str(val).lower():
            val_str = str(val).strip()
            for i, month in enumerate(MONTHS, 1):
                if val_str.lower().startswith(month.lower()):
                    parts = val_str.split()
                    if len(parts) == 2:
                        try:
                            year = int(parts[1])
                            months.append((i, year, val_str))
                        except:
                            pass
                    break

    accounts = []
    for row_idx in range(header_row + 1, len(df)):
        account_name = df.iloc[row_idx, 0]
        if pd.isna(account_name) or not str(account_name).strip():
            continue
        account_name = str(account_name).strip()
        if 'cash basis' in account_name.lower():
            continue

        values = {}
        for col_idx in range(1, min(len(df.columns), len(months) + 1)):
            val = df.iloc[row_idx, col_idx]
            if pd.notna(val):
                try:
                    values[col_idx - 1] = float(val)
                except:
                    values[col_idx - 1] = 0
            else:
                values[col_idx - 1] = 0

        is_total = account_name.lower().startswith('total') or account_name in ['Net Income', 'Gross Profit']
        accounts.append({
            'name': account_name,
            'values': values,
            'is_total': is_total,
            'is_header': account_name in ['Income', 'Expenses', 'Cost of Sales']
        })

    return accounts, months

def col_letter(col_num):
    result = ''
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result

def main():
    print('Loading sample files...')
    pl_data = pd.read_excel('SAMPLE PL.xlsx', header=None)
    bs_data = pd.read_excel('SAMPLE BS.xlsx', header=None)

    pl_accounts, months = parse_financial_data(pl_data)
    bs_accounts, _ = parse_financial_data(bs_data)

    print(f'Parsed {len(pl_accounts)} P&L accounts')
    print(f'Parsed {len(bs_accounts)} BS accounts')
    print(f'Found {len(months)} months: {months[0][2]} to {months[-1][2]}')

    print('\nCreating Excel model with xlwings...')
    app = xw.App(visible=False)

    try:
        wb = app.books.add()

        # Create sheets
        menu_sheet = wb.sheets[0]
        menu_sheet.name = 'Menu'

        source_pl = wb.sheets.add('Source_PL', after=menu_sheet)
        source_bs = wb.sheets.add('Source_BS', after=source_pl)
        pl_sheet = wb.sheets.add('PL', after=source_bs)
        bs_sheet = wb.sheets.add('Balance_Sheet', after=pl_sheet)
        cf_sheet = wb.sheets.add('Cash_Flow', after=bs_sheet)

        # Populate Source P&L
        print('Populating Source P&L...')
        source_pl.range('A1').value = 'Account'
        for i, (m, y, name) in enumerate(months):
            source_pl.range((1, i + 2)).value = name

        for row_idx, account in enumerate(pl_accounts, 2):
            source_pl.range((row_idx, 1)).value = account['name']
            for col_idx, val in account['values'].items():
                source_pl.range((row_idx, col_idx + 2)).value = val

        # Populate Source BS
        print('Populating Source BS...')
        source_bs.range('A1').value = 'Account'
        for i, (m, y, name) in enumerate(months):
            source_bs.range((1, i + 2)).value = name

        for row_idx, account in enumerate(bs_accounts, 2):
            source_bs.range((row_idx, 1)).value = account['name']
            for col_idx, val in account['values'].items():
                source_bs.range((row_idx, col_idx + 2)).value = val

        # Create named ranges
        pl_last_row = len(pl_accounts) + 1
        bs_last_row = len(bs_accounts) + 1
        last_col = len(months) + 1

        wb.names.add('SourcePL', f'=Source_PL!$A$1:${col_letter(last_col)}${pl_last_row}')
        wb.names.add('SourceBS', f'=Source_BS!$A$1:${col_letter(last_col)}${bs_last_row}')

        # Create P&L report with SUMIF
        print('Creating P&L report with SUMIF formulas...')
        pl_sheet.range('A1').value = 'AMP Quality Energy Services'
        pl_sheet.range('A1').font.size = 14
        pl_sheet.range('A1').font.bold = True

        pl_sheet.range('A2').value = 'Profit & Loss Statement'
        pl_sheet.range('A2').font.bold = True

        pl_sheet.range('A4').value = 'Account'
        for i, (m, y, name) in enumerate(months):
            pl_sheet.range((4, i + 2)).value = f'{MONTHS[m-1][:3]} {y}'

        ytd_col = len(months) + 3
        pl_sheet.range((4, ytd_col)).value = 'YTD'
        pl_sheet.range('A4').expand('right').font.bold = True

        for row_idx, account in enumerate(pl_accounts, 5):
            pl_sheet.range((row_idx, 1)).value = account['name']

            if account['is_total']:
                pl_sheet.range((row_idx, 1)).font.bold = True
                for col in range(1, ytd_col + 1):
                    pl_sheet.range((row_idx, col)).color = (211, 211, 211)

            # SUMIF formulas
            for i in range(len(months)):
                col = i + 2
                formula = f'=SUMIF(Source_PL!$A:$A,$A{row_idx},Source_PL!{col_letter(col)}:{col_letter(col)})'
                pl_sheet.range((row_idx, col)).formula = formula
                pl_sheet.range((row_idx, col)).number_format = '#,##0'

            # YTD
            first_col = col_letter(2)
            last_month_col = col_letter(len(months) + 1)
            pl_sheet.range((row_idx, ytd_col)).formula = f'=SUM({first_col}{row_idx}:{last_month_col}{row_idx})'
            pl_sheet.range((row_idx, ytd_col)).number_format = '#,##0'

        # Create BS report
        print('Creating Balance Sheet report...')
        bs_sheet.range('A1').value = 'AMP Quality Energy Services'
        bs_sheet.range('A1').font.size = 14
        bs_sheet.range('A1').font.bold = True

        bs_sheet.range('A2').value = 'Balance Sheet'
        bs_sheet.range('A2').font.bold = True

        bs_sheet.range('A4').value = 'Account'
        for i, (m, y, name) in enumerate(months):
            bs_sheet.range((4, i + 2)).value = f'{MONTHS[m-1][:3]} {y}'
        bs_sheet.range('A4').expand('right').font.bold = True

        for row_idx, account in enumerate(bs_accounts, 5):
            bs_sheet.range((row_idx, 1)).value = account['name']

            if account['is_total']:
                bs_sheet.range((row_idx, 1)).font.bold = True
                for col in range(1, len(months) + 2):
                    bs_sheet.range((row_idx, col)).color = (211, 211, 211)

            for i in range(len(months)):
                col = i + 2
                formula = f'=SUMIF(Source_BS!$A:$A,$A{row_idx},Source_BS!{col_letter(col)}:{col_letter(col)})'
                bs_sheet.range((row_idx, col)).formula = formula
                bs_sheet.range((row_idx, col)).number_format = '#,##0'

        # Create Cash Flow
        print('Creating Cash Flow statement...')
        cf_sheet.range('A1').value = 'AMP Quality Energy Services'
        cf_sheet.range('A1').font.size = 14
        cf_sheet.range('A1').font.bold = True

        cf_sheet.range('A2').value = 'Cash Flow Statement'
        cf_sheet.range('A2').font.bold = True

        cf_sheet.range('A4').value = 'Description'
        for i, (m, y, name) in enumerate(months):
            cf_sheet.range((4, i + 2)).value = f'{MONTHS[m-1][:3]} {y}'
        cf_ytd_col = len(months) + 3
        cf_sheet.range((4, cf_ytd_col)).value = 'YTD'
        cf_sheet.range('A4').expand('right').font.bold = True

        # Cash flow items
        cf_items = [
            ('OPERATING ACTIVITIES', True),
            ('Net Income', False),
            ('Depreciation', False),
            ('Change in Working Capital', True),
            ('  (Inc)/Dec in Accounts Receivable', False),
            ('  Inc/(Dec) in Accounts Payable', False),
            ('Net Cash from Operations', False),
            ('', False),
            ('INVESTING ACTIVITIES', True),
            ('Capital Expenditures', False),
            ('Net Cash from Investing', False),
            ('', False),
            ('FINANCING ACTIVITIES', True),
            ('Debt Changes', False),
            ('Distributions', False),
            ('Net Cash from Financing', False),
            ('', False),
            ('NET CHANGE IN CASH', False),
            ('Beginning Cash', False),
            ('Ending Cash', False),
        ]

        row = 5
        for item_name, is_header in cf_items:
            cf_sheet.range((row, 1)).value = item_name
            if is_header:
                cf_sheet.range((row, 1)).font.bold = True
            elif 'Net Cash' in item_name or 'NET CHANGE' in item_name or 'Ending Cash' in item_name:
                cf_sheet.range((row, 1)).font.bold = True
                for col in range(1, cf_ytd_col + 1):
                    cf_sheet.range((row, col)).color = (211, 211, 211)

            # Add formulas
            if item_name == 'Net Income':
                for i in range(len(months)):
                    col = i + 2
                    formula = f'=SUMIF(Source_PL!$A:$A,"Net Income",Source_PL!{col_letter(col)}:{col_letter(col)})'
                    cf_sheet.range((row, col)).formula = formula
                    cf_sheet.range((row, col)).number_format = '#,##0'
            elif not is_header and item_name:
                for col in range(2, len(months) + 2):
                    cf_sheet.range((row, col)).value = 0
                    cf_sheet.range((row, col)).number_format = '#,##0'

            # YTD for non-headers
            if item_name and not is_header:
                first = col_letter(2)
                last = col_letter(len(months) + 1)
                cf_sheet.range((row, cf_ytd_col)).formula = f'=SUM({first}{row}:{last}{row})'
                cf_sheet.range((row, cf_ytd_col)).number_format = '#,##0'

            row += 1

        # Menu sheet
        print('Creating Menu sheet...')
        menu_sheet.range('B2').value = 'AMP Quality Energy Services'
        menu_sheet.range('B2').font.size = 24
        menu_sheet.range('B2').font.bold = True

        menu_sheet.range('B3').value = 'DNA Financial Model'
        menu_sheet.range('B3').font.size = 14

        menu_sheet.range('B5').value = 'Current Month:'
        menu_sheet.range('C5').value = months[-1][2]

        menu_sheet.range('B7').value = 'MACROS (Press Alt+F8)'
        menu_sheet.range('B7').font.bold = True

        menu_sheet.range('B8').value = 'UploadPLFile - Import new P&L'
        menu_sheet.range('B9').value = 'UploadBSFile - Import new Balance Sheet'
        menu_sheet.range('B10').value = 'RefreshAll - Recalculate formulas'

        # Add VBA
        print('Adding VBA macros...')
        vba_code = DNAModelApp.VBA_CODE
        vba_module = wb.api.VBProject.VBComponents.Add(1)
        vba_module.Name = 'DNAModel'
        vba_module.CodeModule.AddFromString(vba_code)

        # Autofit columns
        for sheet in [source_pl, source_bs, pl_sheet, bs_sheet, cf_sheet, menu_sheet]:
            sheet.autofit()

        # Save
        output_path = 'AMP_Quality_Energy_DNA_Model.xlsm'
        wb.save(output_path)
        print(f'\nSaved: {output_path}')

        wb.close()

    finally:
        app.quit()

    print('\nDone! The model is ready to use.')
    print('- Open the .xlsm file in Excel')
    print('- Macros are already embedded')
    print('- Press Alt+F8 to run macros')

if __name__ == '__main__':
    main()
