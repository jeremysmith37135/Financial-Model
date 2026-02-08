"""
Automated conversion script: xlwings → openpyxl
Handles the most common patterns. Edge cases get marked with # CONVERT: for manual review.
"""
import re
import sys

def convert_file(input_path, output_path):
    with open(input_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    converted = []
    i = 0
    stats = {'total_changes': 0, 'manual_review': 0}

    while i < len(lines):
        line = lines[i]
        original = line
        indent = len(line) - len(line.lstrip())
        ind = ' ' * indent
        stripped = line.strip()

        # Skip already-converted imports and helper section
        if 'import xlwings' in line or 'from xlwings' in line:
            i += 1
            continue

        # =============================================
        # xw.App() and app-level operations → remove/replace
        # =============================================
        if re.match(r'\s*app\s*=\s*xw\.App\(', line):
            line = f'{ind}# openpyxl: no Excel app needed\n'
            stats['total_changes'] += 1
        elif re.match(r'\s*app\.display_alerts\s*=', line):
            line = f'{ind}# openpyxl: no display_alerts needed\n'
            stats['total_changes'] += 1
        elif re.match(r'\s*app\.screen_updating\s*=', line):
            line = f'{ind}# openpyxl: no screen_updating needed\n'
            stats['total_changes'] += 1
        elif re.match(r'\s*app\.calculation\s*=', line):
            line = f'{ind}# openpyxl: no calculation mode needed\n'
            stats['total_changes'] += 1
        elif re.match(r'\s*app\.quit\(\)', line):
            line = f'{ind}# openpyxl: no app to quit\n'
            stats['total_changes'] += 1
        elif re.match(r'\s*app\.books\.add\(\)', stripped) or 'app.books.add()' in line:
            line = line.replace('app.books.add()', 'Workbook()')
            stats['total_changes'] += 1
        elif 'app.books.open(' in line:
            line = re.sub(r'app\.books\.open\(([^)]+)\)', r'load_workbook(\1, keep_vba=True)', line)
            stats['total_changes'] += 1

        # =============================================
        # wb.sheets['name'] → wb['name']
        # =============================================
        if "wb.sheets['" in line or 'wb.sheets["' in line:
            line = re.sub(r'wb\.sheets\[([\'"][^\'"]+[\'"])\]', r'wb[\1]', line)
            stats['total_changes'] += 1

        # wb.sheets.add('Name', after=...) → wb.create_sheet('Name')
        if 'wb.sheets.add(' in line:
            m = re.search(r'wb\.sheets\.add\(([\'"][^\'"]+[\'"])[^)]*\)', line)
            if m:
                line = re.sub(r'wb\.sheets\.add\(([\'"][^\'"]+[\'"])[^)]*\)', r'wb.create_sheet(\1)', line)
                stats['total_changes'] += 1

        # wb.sheets[0] → wb.worksheets[0]  (but also handle .name = assignment)
        if re.search(r'wb\.sheets\[\d+\]', line):
            line = re.sub(r'wb\.sheets\[(\d+)\]', r'wb.worksheets[\1]', line)
            stats['total_changes'] += 1

        # for sheet in wb.sheets → for sheet in wb.worksheets
        if 'in wb.sheets' in line or 'wb.sheets:' in line:
            line = line.replace('wb.sheets', 'wb.worksheets')
            stats['total_changes'] += 1

        # [s.name for s in wb.sheets] → wb.sheetnames
        if 'for s in wb.sheets]' in line:
            line = line.replace('[s.name for s in wb.sheets]', 'wb.sheetnames')
            stats['total_changes'] += 1

        # wb.names.add(...) → wb.defined_names.new(...)
        if 'wb.names.add(' in line:
            line = line.replace('wb.names.add(', 'wb.defined_names.new(')
            stats['total_changes'] += 1
        if "wb.names['" in line and '.delete()' in line:
            m = re.search(r"wb\.names\['([^']+)'\]\.delete\(\)", line)
            if m:
                name = m.group(1)
                line = line.replace(m.group(0), f"wb.defined_names.delete('{name}')")
                stats['total_changes'] += 1

        # sheet.name property access
        if '.sheets.count' in line:
            line = line.replace('.sheets.count', '.worksheets.__len__()')
            stats['total_changes'] += 1

        # =============================================
        # sheet.range((row, col)).value → sheet.cell(row=row, column=col).value
        # Handle both assignment and read
        # =============================================
        # Pattern: var.range((expr, expr)) - with tuple of expressions
        if '.range((' in line and '.api.' not in line:
            # sheet.range((r, c)).value = val
            m = re.search(r'(\w+)\.range\(\(([^,]+),\s*([^)]+)\)\)\.value\s*=\s*(.*)', line)
            if m:
                sheet_var, row_expr, col_expr, val = m.groups()
                line = f'{ind}{sheet_var}.cell(row={row_expr.strip()}, column={col_expr.strip()}).value = {val.strip()}\n'
                stats['total_changes'] += 1
            else:
                # sheet.range((r, c)).value (read)
                line = re.sub(
                    r'(\w+)\.range\(\(([^,]+),\s*([^)]+)\)\)\.value',
                    lambda m: f'{m.group(1)}.cell(row={m.group(2).strip()}, column={m.group(3).strip()}).value',
                    line
                )
                if line != original:
                    stats['total_changes'] += 1

                # sheet.range((r, c)).formula = val → .value = val (openpyxl uses .value for formulas)
                m2 = re.search(r'(\w+)\.range\(\(([^,]+),\s*([^)]+)\)\)\.formula\s*=\s*(.*)', line)
                if m2:
                    sheet_var, row_expr, col_expr, val = m2.groups()
                    line = f'{ind}{sheet_var}.cell(row={row_expr.strip()}, column={col_expr.strip()}).value = {val.strip()}\n'
                    stats['total_changes'] += 1

                # Remaining range((r,c)) patterns - just convert to cell()
                line = re.sub(
                    r'(\w+)\.range\(\(([^,]+),\s*([^)]+)\)\)',
                    lambda m: f'{m.group(1)}.cell(row={m.group(2).strip()}, column={m.group(3).strip()})',
                    line
                )
                if line != original:
                    stats['total_changes'] += 1

        # =============================================
        # sheet.range('A1').value → sheet['A1'].value  (string cell references)
        # =============================================
        if ".range('" in line and '.api.' not in line and '.range((' not in line:
            # .formula = → .value =
            line = re.sub(r"\.range\((['\"][^'\"]+['\"])\)\.formula\s*=", r"[\1].value =", line)
            if line != original:
                stats['total_changes'] += 1
            # .value = or .value (read)
            line = re.sub(r"\.range\((['\"][^'\"]+['\"])\)", r'[\1]', line)
            if line != original:
                stats['total_changes'] += 1

        # sheet.range(f'...').value → sheet[f'...'].value
        if ".range(f'" in line and '.api.' not in line and '.range((' not in line:
            line = re.sub(r"\.range\((f'[^']*')\)\.formula\s*=", r'[\1].value =', line)
            line = re.sub(r"\.range\((f'[^']*')\)", r'[\1]', line)
            if line != original:
                stats['total_changes'] += 1
        if '.range(f"' in line and '.api.' not in line and '.range((' not in line:
            line = re.sub(r'\.range\((f"[^"]*")\)\.formula\s*=', r'[\1].value =', line)
            line = re.sub(r'\.range\((f"[^"]*")\)', r'[\1]', line)
            if line != original:
                stats['total_changes'] += 1

        # =============================================
        # .formula = ... → .value = ... for any remaining
        # =============================================
        if '.formula =' in line and '.value =' not in line:
            line = line.replace('.formula =', '.value =')
            if line != original:
                stats['total_changes'] += 1
        if '.formula=' in line and '.value=' not in line:
            line = line.replace('.formula=', '.value=')
            if line != original:
                stats['total_changes'] += 1

        # =============================================
        # Font properties: .font.bold = True, .font.size = X, etc.
        # These need Font() objects in openpyxl. Mark for manual review if complex.
        # Simple single-property cases we can handle:
        # =============================================
        if '.font.bold = True' in line and '.api.' not in line:
            # cell.font.bold = True → cell.font = Font(bold=True)
            line = line.replace('.font.bold = True', '.font = Font(bold=True)')
            stats['total_changes'] += 1
        elif '.font.bold = False' in line and '.api.' not in line:
            line = line.replace('.font.bold = False', '.font = Font(bold=False)')
            stats['total_changes'] += 1
        elif '.font.italic = True' in line and '.api.' not in line:
            line = line.replace('.font.italic = True', '.font = Font(italic=True)')
            stats['total_changes'] += 1
        elif '.font.underline = True' in line and '.api.' not in line:
            line = line.replace('.font.underline = True', '.font = Font(underline="single")')
            stats['total_changes'] += 1

        # .font.size = X
        m_size = re.search(r'\.font\.size\s*=\s*(\d+)', line)
        if m_size and '.api.' not in line and '.font = Font' not in line:
            size = m_size.group(1)
            line = re.sub(r'\.font\.size\s*=\s*\d+', f'.font = Font(size={size})', line)
            stats['total_changes'] += 1

        # .font.name = 'X'
        m_name = re.search(r"\.font\.name\s*=\s*('[^']+')", line)
        if m_name and '.api.' not in line and '.font = Font' not in line:
            name = m_name.group(1)
            line = re.sub(r"\.font\.name\s*=\s*'[^']+'", f'.font = Font(name={name})', line)
            stats['total_changes'] += 1

        # .font.color = (r, g, b) → .font = Font(color="RRGGBB")
        m_fc = re.search(r'\.font\.color\s*=\s*\((\d+),\s*(\d+),\s*(\d+)\)', line)
        if m_fc and '.api.' not in line:
            r, g, b = int(m_fc.group(1)), int(m_fc.group(2)), int(m_fc.group(3))
            hex_color = f'{r:02X}{g:02X}{b:02X}'
            line = re.sub(r'\.font\.color\s*=\s*\(\d+,\s*\d+,\s*\d+\)', f'.font = Font(color="{hex_color}")', line)
            stats['total_changes'] += 1

        # .font.color = DARK_BLUE (variable) → .font = FONT_DARK_BLUE
        if '.font.color = DARK_BLUE' in line:
            line = line.replace('.font.color = DARK_BLUE', '.font = FONT_DARK_BLUE')
            stats['total_changes'] += 1
        if '.font.color = ACCENT_BLUE' in line:
            line = line.replace('.font.color = ACCENT_BLUE', '.font = FONT_ACCENT_BLUE')
            stats['total_changes'] += 1
        if '.font.color = GRAY' in line:
            line = line.replace('.font.color = GRAY', '.font = FONT_GRAY_TEXT')
            stats['total_changes'] += 1
        if '.font.color = LINK_BLUE' in line:
            line = line.replace('.font.color = LINK_BLUE', '.font = FONT_LINK_BLUE')
            stats['total_changes'] += 1

        # =============================================
        # Background color: .color = (r, g, b) → .fill = PatternFill(...)
        # =============================================
        m_bg = re.search(r'\.color\s*=\s*\((\d+),\s*(\d+),\s*(\d+)\)', line)
        if m_bg and '.font' not in line and '.api.' not in line:
            r, g, b = int(m_bg.group(1)), int(m_bg.group(2)), int(m_bg.group(3))
            hex_color = f'{r:02X}{g:02X}{b:02X}'
            line = re.sub(
                r'\.color\s*=\s*\(\d+,\s*\d+,\s*\d+\)',
                f'.fill = PatternFill(start_color="{hex_color}", end_color="{hex_color}", fill_type="solid")',
                line
            )
            stats['total_changes'] += 1

        # .color = DARK_BLUE → .fill = FILL_DARK_BLUE
        if re.search(r'\.color\s*=\s*DARK_BLUE', line) and '.font' not in line:
            line = re.sub(r'\.color\s*=\s*DARK_BLUE', '.fill = FILL_DARK_BLUE', line)
            stats['total_changes'] += 1
        if re.search(r'\.color\s*=\s*SUBTOTAL_GRAY', line) and '.font' not in line:
            line = re.sub(r'\.color\s*=\s*SUBTOTAL_GRAY', '.fill = FILL_SUBTOTAL_GRAY', line)
            stats['total_changes'] += 1
        if re.search(r'\.color\s*=\s*ACCENT_BLUE', line) and '.font' not in line:
            line = re.sub(r'\.color\s*=\s*ACCENT_BLUE', '.fill = FILL_ACCENT_BLUE', line)
            stats['total_changes'] += 1

        # =============================================
        # number_format stays the same in openpyxl
        # .number_format = fmt → .number_format = fmt (no change needed!)
        # =============================================

        # =============================================
        # column_width: sheet.range('A:A').column_width = X
        # =============================================
        m_cw = re.search(r"(\w+)\[?'([A-Z]):([A-Z])'\]?\.column_width\s*=\s*(\d+)", line)
        if m_cw:
            sheet_var = m_cw.group(1)
            col_start = m_cw.group(2)
            col_end = m_cw.group(3)
            width = m_cw.group(4)
            if col_start == col_end:
                line = f"{ind}{sheet_var}.column_dimensions['{col_start}'].width = {width}\n"
            else:
                line = f"{ind}for _c in range(ord('{col_start}'), ord('{col_end}')+1):\n"
                line += f"{ind}    {sheet_var}.column_dimensions[chr(_c)].width = {width}\n"
            stats['total_changes'] += 1

        # =============================================
        # .api patterns - mark complex ones for review
        # =============================================
        if '.api.' in line:
            # .api.Validation.Delete() + .api.Validation.Add()
            if 'Validation.Delete()' in line:
                # We'll handle this by removing old validation and adding new
                line = f'{ind}# Validation handled via DataValidation object\n'
                stats['total_changes'] += 1
            elif 'Validation.Add(' in line:
                m_val = re.search(r"Formula1=(.+?)(?:\)|,)", line)
                formula = m_val.group(1) if m_val else '""'
                line = f'{ind}# CONVERT: DataValidation - see manual fix section\n'
                stats['manual_review'] += 1
            elif 'EntireColumn.Hidden = True' in line:
                line = f'{ind}# CONVERT: hide columns - use hide_column() or hide_columns_range()\n'
                stats['manual_review'] += 1
            elif 'EntireRow.Hidden = True' in line:
                line = f'{ind}# CONVERT: hide row - use hide_row()\n'
                stats['manual_review'] += 1
            elif 'Hyperlinks.Add' in line:
                line = f'{ind}# CONVERT: add hyperlink - use cell.hyperlink\n'
                stats['manual_review'] += 1
            elif 'Borders(' in line:
                line = f'{ind}# CONVERT: borders - use Border()/Side() objects\n'
                stats['manual_review'] += 1
            elif 'HorizontalAlignment' in line:
                line = f'{ind}# CONVERT: alignment - use Alignment(horizontal=...)\n'
                stats['manual_review'] += 1
            elif 'Outline.ShowLevels' in line:
                line = f'{ind}# CONVERT: outline levels - use group_rows()/group_cols()\n'
                stats['manual_review'] += 1
            elif 'UsedRange' in line:
                line = f'{ind}# CONVERT: UsedRange - use sheet.max_row/max_column\n'
                stats['manual_review'] += 1
            elif 'Tab.Color' in line:
                m_tc = re.search(r'\.api\.Tab\.Color\s*=\s*0x([0-9A-Fa-f]+)', line)
                if m_tc:
                    color = m_tc.group(1).upper().zfill(6)
                    sheet_var = line.split('.api')[0].strip()
                    line = f'{ind}{sheet_var}.sheet_properties.tabColor = "{color}"\n'
                    stats['total_changes'] += 1
                else:
                    line = f'{ind}# CONVERT: tab color\n'
                    stats['manual_review'] += 1
            elif 'Cells.Replace' in line:
                line = f'{ind}# CONVERT: Cells.Replace - use manual find/replace\n'
                stats['manual_review'] += 1
            elif 'EntireColumn.Insert' in line:
                line = f'{ind}# CONVERT: column insert - use sheet.insert_cols()\n'
                stats['manual_review'] += 1
            elif 'Rows.Group' in line:
                line = f'{ind}# CONVERT: row grouping - use group_rows()\n'
                stats['manual_review'] += 1
            elif 'VBProject' in line or 'VBComponents' in line or 'CodeModule' in line:
                line = f'{ind}# VBA: handled via keep_vba=True template approach\n'
                stats['total_changes'] += 1
            elif 'ActiveWindow' in line:
                line = f'{ind}# CONVERT: window settings - openpyxl sheet views\n'
                stats['manual_review'] += 1
            else:
                line = f'{ind}# CONVERT: .api call needs manual conversion: {stripped}\n'
                stats['manual_review'] += 1

        # =============================================
        # .end('right').column / .end('down').row → max_column/max_row
        # =============================================
        if ".end('right').column" in line or '.end("right").column' in line:
            line = f'{ind}# CONVERT: .end("right") → use sheet.max_column\n'
            stats['manual_review'] += 1
        if ".end('down').row" in line or '.end("down").row' in line:
            line = f'{ind}# CONVERT: .end("down") → use sheet.max_row\n'
            stats['manual_review'] += 1

        # =============================================
        # .autofit() → manual width (openpyxl doesn't have autofit)
        # =============================================
        if '.autofit()' in line:
            line = f'{ind}# openpyxl: autofit not available - using fixed widths\n'
            stats['total_changes'] += 1

        # =============================================
        # .add_hyperlink() → cell.hyperlink
        # =============================================
        if '.add_hyperlink(' in line and '.api.' not in line:
            line = f'{ind}# CONVERT: add_hyperlink → cell.hyperlink = ...\n'
            stats['manual_review'] += 1

        # =============================================
        # .clear() on ranges → delete rows or set values to None
        # =============================================
        if ".clear()" in line:
            line = f'{ind}# CONVERT: .clear() → iterate and set None, or delete_rows\n'
            stats['manual_review'] += 1

        # =============================================
        # sheet.delete() → wb.remove(sheet)
        # =============================================
        if re.search(r'(\w+)\.delete\(\)', line) and 'sheet' in line.lower():
            line = f'{ind}# CONVERT: sheet.delete() → wb.remove(sheet)\n'
            stats['manual_review'] += 1

        # =============================================
        # _col_letter helper → get_column_letter (already imported)
        # =============================================
        if 'self._col_letter(' in line:
            line = line.replace('self._col_letter(', 'get_column_letter(')
            stats['total_changes'] += 1

        # =============================================
        # Range with two tuples: sheet.range((r1,c1),(r2,c2))
        # =============================================
        if re.search(r'\.range\(\([^)]+\),\s*\([^)]+\)\)', line) and '.api.' not in line:
            line = line.rstrip() + '  # CONVERT: range with two tuples → iterate cells\n'
            stats['manual_review'] += 1

        # =============================================
        # Remaining .range( patterns that weren't caught
        # =============================================
        if '.range(' in line and '.api.' not in line and '# CONVERT' not in line:
            # Check if it's something we missed
            if not any(x in line for x in ['.cell(', "['", '[f"', "[f'", 'CONVERT', '# openpyxl']):
                if '.range(' in line:
                    line = line.rstrip() + '  # CONVERT: remaining .range() call\n'
                    stats['manual_review'] += 1

        converted.append(line)
        i += 1

    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(converted)

    print(f"\nConversion complete!")
    print(f"  Automated changes: {stats['total_changes']}")
    print(f"  Manual review needed: {stats['manual_review']}")
    print(f"  Output: {output_path}")
    return stats

if __name__ == '__main__':
    input_file = r"C:\Users\jerju\OneDrive\Focus CFO\AI\DNA Model\webapp\desktop_app_365.py"
    output_file = r"C:\Users\jerju\OneDrive\Focus CFO\AI\DNA Model\webapp\desktop_app_365.py"
    convert_file(input_file, output_file)
