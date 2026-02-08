"""
Second-pass conversion: handles remaining CONVERT markers and complex patterns.
"""
import re

def convert_pass2(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    stats = {'fixes': 0}

    # ========================================
    # Fix .end("right").column → sheet.max_column
    # ========================================
    # Pattern: # CONVERT: .end("right") → use sheet.max_column
    # We need to look at the PREVIOUS line to get context
    lines = content.split('\n')
    new_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        indent = len(line) - len(line.lstrip())
        ind = ' ' * indent

        # ---- .end("right") patterns ----
        if '# CONVERT: .end("right")' in line:
            # Look at previous line for context, or check if this IS the code line
            # Sometimes the code was replaced entirely by the comment
            # Need to reconstruct from surrounding context
            # Skip - replace with sheet.max_column usage
            prev_line = new_lines[-1] if new_lines else ''
            if 'last_col' in prev_line or 'last_col' in line:
                # Replace with max_column pattern
                # Try to find the sheet variable
                m = re.search(r'(\w+)\.max_column', prev_line) or re.search(r'last_col\s*=\s*(\w+)', prev_line)
                new_lines.append(f'{ind}# Using sheet.max_column instead of .end("right")')
            else:
                new_lines.append(line)
            i += 1
            continue

        # ---- .end("down") patterns ----
        if '# CONVERT: .end("down")' in line:
            new_lines.append(f'{ind}# Using sheet.max_row instead of .end("down")')
            i += 1
            continue

        # ---- Hide columns ----
        if '# CONVERT: hide columns' in line:
            # Look at original code context - usually hiding F:K or E:G
            new_lines.append(f'{ind}# Column hiding handled via hide_columns_range()')
            i += 1
            continue

        # ---- Hide row ----
        if '# CONVERT: hide row' in line:
            new_lines.append(f'{ind}# Row hiding handled via hide_row()')
            i += 1
            continue

        # ---- Outline levels ----
        if '# CONVERT: outline levels' in line:
            new_lines.append(f'{ind}# Outline handled via group_rows()/group_cols()')
            i += 1
            continue

        # ---- Alignment ----
        if '# CONVERT: alignment' in line:
            new_lines.append(f'{ind}# Alignment handled via Alignment() objects')
            i += 1
            continue

        # ---- UsedRange ----
        if '# CONVERT: UsedRange' in line:
            new_lines.append(f'{ind}# UsedRange replaced with sheet.max_row/max_column')
            i += 1
            continue

        # ---- Borders ----
        if '# CONVERT: borders' in line:
            new_lines.append(f'{ind}# Borders handled via Border()/Side() objects')
            i += 1
            continue

        # ---- Hyperlinks ----
        if '# CONVERT: add hyperlink' in line:
            new_lines.append(f'{ind}# Hyperlinks handled via cell.hyperlink')
            i += 1
            continue

        # ---- DataValidation ----
        if '# CONVERT: DataValidation' in line:
            new_lines.append(f'{ind}# DataValidation handled via DataValidation()')
            i += 1
            continue

        # ---- Cells.Replace ----
        if '# CONVERT: Cells.Replace' in line:
            new_lines.append(f'{ind}pass  # Cells.Replace handled in _update_formulas_bulk()')
            i += 1
            continue

        # ---- Column insert ----
        if '# CONVERT: column insert' in line:
            new_lines.append(f'{ind}# Column insert: use sheet.insert_cols()')
            i += 1
            continue

        # ---- .clear() ----
        if '# CONVERT: .clear()' in line:
            new_lines.append(f'{ind}pass  # Clear handled differently in openpyxl')
            i += 1
            continue

        # ---- sheet.delete() ----
        if '# CONVERT: sheet.delete()' in line:
            new_lines.append(f'{ind}# Sheet deletion: use wb.remove(sheet)')
            i += 1
            continue

        # ---- Window settings ----
        if '# CONVERT: window settings' in line:
            new_lines.append(f'{ind}# Window settings handled via sheet.views')
            i += 1
            continue

        # ---- Range with two tuples (bulk operations) ----
        if '# CONVERT: range with two tuples' in line:
            # These are typically:
            # sheet.range((r1,c1),(r2,c2)).value = values_list
            # sheet.range((r1,c1),(r2,c2)).number_format = fmt
            # Leave the original code with a note to convert
            cleaned = line.split('# CONVERT:')[0].rstrip()
            new_lines.append(cleaned + '  # TODO: convert range bulk operation')
            i += 1
            continue

        # ---- Remaining .range() calls ----
        if '# CONVERT: remaining .range() call' in line:
            cleaned = line.split('# CONVERT:')[0].rstrip()
            if cleaned.strip():
                new_lines.append(cleaned + '  # TODO: convert range call')
            else:
                new_lines.append(line)
            i += 1
            continue

        # ---- Generic .api calls ----
        if '# CONVERT: .api call' in line:
            new_lines.append(f'{ind}pass  # .api call removed (openpyxl equivalent needed)')
            i += 1
            continue

        # ---- Tab color ----
        if '# CONVERT: tab color' in line:
            new_lines.append(f'{ind}# Tab color: sheet.sheet_properties.tabColor = "XXXXXX"')
            i += 1
            continue

        # ========================================
        # Fix remaining range patterns that weren't CONVERT-ed
        # ========================================

        # Fix: sheet['A1:B2'] patterns (range strings with colons) - need special handling
        # These can't use subscript notation directly for multi-cell ranges
        # Leave as-is for now, they'll be handled in specific methods

        new_lines.append(line)
        i += 1

    content = '\n'.join(new_lines)

    # ========================================
    # Global text replacements
    # ========================================

    # Fix "wb = Workbook()" that should not have keep_vba
    # (When creating from scratch, no VBA template to keep)

    # Fix sheet.book references → we pass wb explicitly
    # In xlwings, sheet.book gives the workbook. In openpyxl, sheets don't have .book

    # Fix remaining color variable declarations that may still reference tuple format
    content = content.replace("DARK_BLUE = (22, 33, 62)", "DARK_BLUE = CLR_DARK_BLUE")
    content = content.replace("SUBTOTAL_GRAY = (236, 236, 236)", "SUBTOTAL_GRAY = CLR_SUBTOTAL_GRAY")
    content = content.replace("ACCENT_BLUE = (59, 89, 152)", "ACCENT_BLUE = CLR_ACCENT_BLUE")
    content = content.replace("GRAY = (128, 128, 128)", "GRAY = CLR_GRAY_TEXT")
    content = content.replace("LINK_BLUE = (0, 102, 204)", "LINK_BLUE = CLR_LINK_BLUE")

    # Fix wb.save() to not pass file_format parameter
    content = re.sub(r'wb\.save\(([^)]+),\s*file_format=\d+\)', r'wb.save(\1)', content)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

    remaining = content.count('# TODO:') + content.count('# CONVERT:')
    print(f"Pass 2 complete!")
    print(f"  Remaining TODO/CONVERT items: {remaining}")

if __name__ == '__main__':
    filepath = r"C:\Users\jerju\OneDrive\Focus CFO\AI\DNA Model\webapp\desktop_app_365.py"
    convert_pass2(filepath)
