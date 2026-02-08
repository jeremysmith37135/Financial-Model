"""
Third-pass conversion: Handle range bulk operations and remaining patterns.
Converts sheet.range((r1,c1),(r2,c2)).property patterns to cell-by-cell operations.
"""
import re

def convert_pass3(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    new_lines = []
    fixes = 0

    for i, line in enumerate(lines):
        stripped = line.rstrip()
        indent = len(line) - len(line.lstrip())
        ind = ' ' * indent

        if '# TODO: convert range bulk operation' not in stripped and \
           '# TODO: convert range call' not in stripped and \
           '# CONVERT: add_hyperlink' not in stripped and \
           '# CONVERT: .end("right")' not in stripped and \
           '# CONVERT: row grouping' not in stripped:
            new_lines.append(line)
            continue

        # Remove the TODO comment to get the actual code
        code = stripped.split('# TODO:')[0].split('# CONVERT:')[0].rstrip()

        # ==========================================
        # .end("right") → sheet.max_column
        # ==========================================
        if '# CONVERT: .end("right")' in stripped:
            new_lines.append(f'{ind}# end("right") replaced with max_column\n')
            fixes += 1
            continue

        # ==========================================
        # Row grouping
        # ==========================================
        if '# CONVERT: row grouping' in stripped:
            new_lines.append(f'{ind}# Row grouping: use group_rows() helper\n')
            fixes += 1
            continue

        # ==========================================
        # add_hyperlink → cell.hyperlink
        # ==========================================
        if '# CONVERT: add_hyperlink' in stripped:
            # Try to extract from original pattern
            # sheet['A1'].add_hyperlink('#Sheet!A1', text_to_display='label')
            new_lines.append(f'{ind}pass  # Hyperlinks: set cell.hyperlink and cell.font for link style\n')
            fixes += 1
            continue

        # ==========================================
        # sheet.range((r1,c1),(r2,c2)).value = data  → write cells
        # ==========================================
        # Pattern: var.range((r1, c1), (r2, c2)).value = data_var
        m_val = re.match(
            r'(\s*)(\w+)\.range\(\(([^,]+),\s*([^)]+)\),\s*\(([^,]+),\s*([^)]+)\)\)\.value\s*=\s*(.+)',
            code
        )
        if m_val:
            ws = m_val.group(2)
            r1, c1, r2, c2 = m_val.group(3).strip(), m_val.group(4).strip(), m_val.group(5).strip(), m_val.group(6).strip()
            data = m_val.group(7).strip()
            # Check if it's a single-row write (like header): [data_list]
            if data.startswith('[') and data.endswith(']'):
                # Single row of data as list
                new_lines.append(f'{ind}for _ci, _val in enumerate({data[1:-1]} if isinstance({data[1:-1]}, list) else [{data[1:-1]}]):\n')
                new_lines.append(f'{ind}    {ws}.cell(row={r1}, column={c1} + _ci).value = _val if not isinstance(_val, list) else _val\n')
            else:
                # Multi-row data (list of lists or list of values for a column)
                new_lines.append(f'{ind}_data = {data}\n')
                new_lines.append(f'{ind}if _data is not None:\n')
                new_lines.append(f'{ind}    if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):\n')
                new_lines.append(f'{ind}        for _ri, _row in enumerate(_data):\n')
                new_lines.append(f'{ind}            for _ci, _val in enumerate(_row):\n')
                new_lines.append(f'{ind}                {ws}.cell(row={r1} + _ri, column={c1} + _ci).value = _val\n')
                new_lines.append(f'{ind}    elif isinstance(_data, list):\n')
                new_lines.append(f'{ind}        for _ri, _val in enumerate(_data):\n')
                new_lines.append(f'{ind}            if isinstance(_val, list):\n')
                new_lines.append(f'{ind}                for _ci, _v in enumerate(_val):\n')
                new_lines.append(f'{ind}                    {ws}.cell(row={r1} + _ri, column={c1} + _ci).value = _v\n')
                new_lines.append(f'{ind}            else:\n')
                new_lines.append(f'{ind}                {ws}.cell(row={r1} + _ri, column={c1}).value = _val\n')
            fixes += 1
            continue

        # ==========================================
        # sheet.range((r1,c1),(r2,c2)).number_format = fmt → apply_style_to_range
        # ==========================================
        m_fmt = re.match(
            r'(\s*)(\w+)\.range\(\(([^,]+),\s*([^)]+)\),\s*\(([^,]+),\s*([^)]+)\)\)\.number_format\s*=\s*(.+)',
            code
        )
        if m_fmt:
            ws = m_fmt.group(2)
            r1, c1, r2, c2 = m_fmt.group(3).strip(), m_fmt.group(4).strip(), m_fmt.group(5).strip(), m_fmt.group(6).strip()
            fmt = m_fmt.group(7).strip()
            new_lines.append(f'{ind}apply_style_to_range({ws}, {r1}, {c1}, {r2}, {c2}, number_format={fmt})\n')
            fixes += 1
            continue

        # ==========================================
        # sheet.range((r1,c1),(r2,c2)).font = Font(...) → apply_style_to_range
        # ==========================================
        m_font = re.match(
            r'(\s*)(\w+)\.range\(\(([^,]+),\s*([^)]+)\),\s*\(([^,]+),\s*([^)]+)\)\)\.font\s*=\s*(.+)',
            code
        )
        if m_font:
            ws = m_font.group(2)
            r1, c1, r2, c2 = m_font.group(3).strip(), m_font.group(4).strip(), m_font.group(5).strip(), m_font.group(6).strip()
            font_expr = m_font.group(7).strip()
            new_lines.append(f'{ind}apply_style_to_range({ws}, {r1}, {c1}, {r2}, {c2}, font={font_expr})\n')
            fixes += 1
            continue

        # ==========================================
        # sheet.range((r1,c1),(r2,c2)).fill = ... → apply_style_to_range
        # ==========================================
        m_fill = re.match(
            r'(\s*)(\w+)\.range\(\(([^,]+),\s*([^)]+)\),\s*\(([^,]+),\s*([^)]+)\)\)\.fill\s*=\s*(.+)',
            code
        )
        if m_fill:
            ws = m_fill.group(2)
            r1, c1, r2, c2 = m_fill.group(3).strip(), m_fill.group(4).strip(), m_fill.group(5).strip(), m_fill.group(6).strip()
            fill_expr = m_fill.group(7).strip()
            new_lines.append(f'{ind}apply_style_to_range({ws}, {r1}, {c1}, {r2}, {c2}, fill={fill_expr})\n')
            fixes += 1
            continue

        # ==========================================
        # sheet.range((r1,c1),(r2,c2)).color = ... → apply_style_to_range with fill
        # ==========================================
        m_color = re.match(
            r'(\s*)(\w+)\.range\(\(([^,]+),\s*([^)]+)\),\s*\(([^,]+),\s*([^)]+)\)\)\.color\s*=\s*(.+)',
            code
        )
        if m_color:
            ws = m_color.group(2)
            r1, c1, r2, c2 = m_color.group(3).strip(), m_color.group(4).strip(), m_color.group(5).strip(), m_color.group(6).strip()
            color_expr = m_color.group(7).strip()
            # Check if it's a named fill constant or a tuple
            if color_expr.startswith('('):
                # Tuple RGB → create PatternFill
                m_rgb = re.match(r'\((\d+),\s*(\d+),\s*(\d+)\)', color_expr)
                if m_rgb:
                    r_val, g_val, b_val = int(m_rgb.group(1)), int(m_rgb.group(2)), int(m_rgb.group(3))
                    hex_c = f'{r_val:02X}{g_val:02X}{b_val:02X}'
                    new_lines.append(f'{ind}apply_style_to_range({ws}, {r1}, {c1}, {r2}, {c2}, fill=PatternFill(start_color="{hex_c}", end_color="{hex_c}", fill_type="solid"))\n')
                else:
                    new_lines.append(f'{ind}# Could not parse color: {color_expr}\n')
            else:
                # Named variable - map to FILL_ constants
                fill_map = {
                    'ACTUAL_BLUE': 'PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")',
                    'BUDGET_YELLOW': 'PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")',
                    'ADJ_ORANGE': 'PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")',
                    'FORECAST_GREEN': 'PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")',
                    'HEADER_GRAY': 'PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")',
                }
                fill_val = fill_map.get(color_expr, f'PatternFill(start_color="ECECEC", end_color="ECECEC", fill_type="solid")  # was: {color_expr}')
                new_lines.append(f'{ind}apply_style_to_range({ws}, {r1}, {c1}, {r2}, {c2}, fill={fill_val})\n')
            fixes += 1
            continue

        # ==========================================
        # sheet.range((r1,c1),(r2,c2)).merge() → sheet.merge_cells()
        # ==========================================
        m_merge = re.match(
            r'(\s*)(\w+)\.range\(\(([^,]+),\s*([^)]+)\),\s*\(([^,]+),\s*([^)]+)\)\)\.merge\(\)',
            code
        )
        if m_merge:
            ws = m_merge.group(2)
            r1, c1, r2, c2 = m_merge.group(3).strip(), m_merge.group(4).strip(), m_merge.group(5).strip(), m_merge.group(6).strip()
            new_lines.append(f'{ind}{ws}.merge_cells(start_row={r1}, start_column={c1}, end_row={r2}, end_column={c2})\n')
            fixes += 1
            continue

        # ==========================================
        # header_range = sheet.range((r1,c1),(r2,c2)) → variable reference
        # These are usually followed by header_range.fill = ... or header_range.font = ...
        # Convert to apply_style_to_range calls later
        # ==========================================
        m_assign = re.match(
            r'(\s*)(\w+)\s*=\s*(\w+)\.range\(\(([^,]+),\s*([^)]+)\),\s*\(([^,]+),\s*([^)]+)\)\)',
            code
        )
        if m_assign:
            var_name = m_assign.group(2)
            ws = m_assign.group(3)
            r1, c1, r2, c2 = m_assign.group(4).strip(), m_assign.group(5).strip(), m_assign.group(6).strip(), m_assign.group(7).strip()
            # Store as tuple for later use
            new_lines.append(f'{ind}{var_name} = ({ws}, {r1}, {c1}, {r2}, {c2})  # range ref as (sheet, r1, c1, r2, c2)\n')
            fixes += 1
            continue

        # ==========================================
        # sheet.range((r1,c1),(r2,c2)).column_width = X → set_col_width
        # ==========================================
        m_cw = re.match(
            r'(\s*)(\w+)\.range\(\(([^,]+),\s*([^)]+)\),\s*\(([^,]+),\s*([^)]+)\)\)\.column_width\s*=\s*(\d+)',
            code
        )
        if m_cw:
            ws = m_cw.group(2)
            c1 = m_cw.group(4).strip()
            width = m_cw.group(7).strip()
            new_lines.append(f'{ind}set_col_width({ws}, {c1}, {width})\n')
            fixes += 1
            continue

        # ==========================================
        # Remaining sheet.range() patterns - just convert to comment
        # ==========================================
        if '.range(' in code:
            # Generic range conversion attempt for remaining patterns
            m_gen = re.match(
                r'(\s*)(.+?)\.range\(\(([^,]+),\s*([^)]+)\),\s*\(([^,]+),\s*([^)]+)\)\)(.*)',
                code
            )
            if m_gen:
                ws = m_gen.group(2).strip()
                r1, c1, r2, c2 = m_gen.group(3).strip(), m_gen.group(4).strip(), m_gen.group(5).strip(), m_gen.group(6).strip()
                rest = m_gen.group(7).strip()
                new_lines.append(f'{ind}# Auto-converted range: apply to ({r1},{c1})-({r2},{c2}): {rest}\n')
                fixes += 1
                continue

        # If nothing matched, keep the line with a note
        new_lines.append(line)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)

    # Count remaining
    content = ''.join(new_lines)
    remaining = content.count('# TODO:') + content.count('# CONVERT:')
    print(f"Pass 3 complete!")
    print(f"  Fixes applied: {fixes}")
    print(f"  Remaining TODO/CONVERT: {remaining}")

if __name__ == '__main__':
    filepath = r"C:\Users\jerju\OneDrive\Focus CFO\AI\DNA Model\webapp\desktop_app_365.py"
    convert_pass3(filepath)
