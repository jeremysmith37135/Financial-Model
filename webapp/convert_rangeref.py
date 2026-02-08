"""
Fix all range ref tuples and their property accesses.
Converts patterns like:
    header_range = (sheet, r1, c1, r2, c2)  # range ref
    header_range.font = Font(bold=True)
    header_range.fill = FILL_DARK_BLUE
Into:
    apply_style_to_range(sheet, r1, c1, r2, c2, font=Font(bold=True))
    apply_style_to_range(sheet, r1, c1, r2, c2, fill=FILL_DARK_BLUE)
"""
import re

def fix_rangerefs(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # First pass: collect all range ref variable names and their definitions
    range_refs = {}  # var_name -> (line_num, sheet, r1, c1, r2, c2)

    for i, line in enumerate(lines):
        if '# range ref as' in line:
            s = line.strip()
            # Parse: var = (sheet, r1, c1, r2, c2)  # range ref...
            m = re.match(r'(\w+)\s*=\s*\((\w+),\s*(.+?),\s*(.+?),\s*(.+?),\s*(.+?)\)\s*#', s)
            if m:
                var_name = m.group(1)
                range_refs[var_name] = {
                    'line': i,
                    'sheet': m.group(2),
                    'r1': m.group(3).strip(),
                    'c1': m.group(4).strip(),
                    'r2': m.group(5).strip(),
                    'c2': m.group(6).strip()
                }
            else:
                # Try without closing paren (some are missing it)
                m2 = re.match(r'(\w+)\s*=\s*\((\w+),\s*(.+?),\s*(.+?),\s*(.+?),\s*(.+?)\s*#', s)
                if m2:
                    var_name = m2.group(1)
                    range_refs[var_name] = {
                        'line': i,
                        'sheet': m2.group(2),
                        'r1': m2.group(3).strip(),
                        'c1': m2.group(4).strip(),
                        'r2': m2.group(5).strip(),
                        'c2': m2.group(6).strip().rstrip(')')
                    }

    print(f"Found {len(range_refs)} range ref variables")

    # Second pass: convert all lines
    new_lines = []
    skip_set = set()

    for i, line in enumerate(lines):
        s = line.strip()
        indent = len(line) - len(line.lstrip())
        ind = ' ' * indent

        # Skip range ref definition lines - replace with comment
        if '# range ref as' in line:
            # Check if any subsequent line uses this variable for styling
            # If so, we'll inline the apply_style_to_range calls
            var_match = re.match(r'\s*(\w+)\s*=\s*\(', line)
            if var_match:
                var_name = var_match.group(1)
                if var_name in range_refs:
                    ref = range_refs[var_name]
                    new_lines.append(f'{ind}# Range: {var_name} = ({ref["sheet"]}, {ref["r1"]}, {ref["c1"]}, {ref["r2"]}, {ref["c2"]})\n')
                    continue
            new_lines.append(line)
            continue

        # Check if this line uses a range ref variable
        converted = False
        for var_name, ref in range_refs.items():
            sheet = ref['sheet']
            r1, c1, r2, c2 = ref['r1'], ref['c1'], ref['r2'], ref['c2']

            # var.font = Font(...)
            m = re.match(rf'\s*{re.escape(var_name)}\.font\s*=\s*(.+)', s)
            if m:
                font_val = m.group(1).strip()
                new_lines.append(f'{ind}apply_style_to_range({sheet}, {r1}, {c1}, {r2}, {c2}, font={font_val})\n')
                converted = True
                break

            # var.fill = ...
            m = re.match(rf'\s*{re.escape(var_name)}\.fill\s*=\s*(.+)', s)
            if m:
                fill_val = m.group(1).strip()
                new_lines.append(f'{ind}apply_style_to_range({sheet}, {r1}, {c1}, {r2}, {c2}, fill={fill_val})\n')
                converted = True
                break

            # var.color = ... (background color)
            m = re.match(rf'\s*{re.escape(var_name)}\.color\s*=\s*(.+)', s)
            if m and '.font' not in s:
                color_val = m.group(1).strip()
                # Convert color value to fill
                rgb_m = re.match(r'\((\d+),\s*(\d+),\s*(\d+)\)', color_val)
                if rgb_m:
                    r_v, g_v, b_v = int(rgb_m.group(1)), int(rgb_m.group(2)), int(rgb_m.group(3))
                    hex_c = f'{r_v:02X}{g_v:02X}{b_v:02X}'
                    fill_expr = f'PatternFill(start_color="{hex_c}", end_color="{hex_c}", fill_type="solid")'
                else:
                    # Named constant - map to FILL_ constants
                    fill_map = {
                        'DARK_BLUE': 'FILL_DARK_BLUE', 'CLR_DARK_BLUE': 'FILL_DARK_BLUE',
                        'SUBTOTAL_GRAY': 'FILL_SUBTOTAL_GRAY', 'CLR_SUBTOTAL_GRAY': 'FILL_SUBTOTAL_GRAY',
                        'ACCENT_BLUE': 'FILL_ACCENT_BLUE', 'CLR_ACCENT_BLUE': 'FILL_ACCENT_BLUE',
                        'HEADER_GRAY': 'PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")',
                    }
                    fill_expr = fill_map.get(color_val, f'PatternFill(start_color="ECECEC", end_color="ECECEC", fill_type="solid")  # was: {color_val}')
                new_lines.append(f'{ind}apply_style_to_range({sheet}, {r1}, {c1}, {r2}, {c2}, fill={fill_expr})\n')
                converted = True
                break

            # var.number_format = ...
            m = re.match(rf'\s*{re.escape(var_name)}\.number_format\s*=\s*(.+)', s)
            if m:
                fmt_val = m.group(1).strip()
                new_lines.append(f'{ind}apply_style_to_range({sheet}, {r1}, {c1}, {r2}, {c2}, number_format={fmt_val})\n')
                converted = True
                break

            # var.alignment = ... or var.horizontal_alignment = ...
            m = re.match(rf'\s*{re.escape(var_name)}\.(alignment|horizontal_alignment)\s*=\s*(.+)', s)
            if m:
                align_val = m.group(2).strip()
                new_lines.append(f'{ind}apply_style_to_range({sheet}, {r1}, {c1}, {r2}, {c2}, alignment={align_val})\n')
                converted = True
                break

            # var.border = ...
            m = re.match(rf'\s*{re.escape(var_name)}\.border\s*=\s*(.+)', s)
            if m:
                border_val = m.group(1).strip()
                new_lines.append(f'{ind}apply_style_to_range({sheet}, {r1}, {c1}, {r2}, {c2}, border={border_val})\n')
                converted = True
                break

            # var.value = ... (reading/writing a range of values)
            m = re.match(rf'\s*{re.escape(var_name)}\.value\s*=\s*(.+)', s)
            if m:
                val = m.group(1).strip()
                new_lines.append(f'{ind}# Write data to range ({r1},{c1})-({r2},{c2})\n')
                new_lines.append(f'{ind}_data = {val}\n')
                new_lines.append(f'{ind}if isinstance(_data, list):\n')
                new_lines.append(f'{ind}    for _ri, _row in enumerate(_data):\n')
                new_lines.append(f'{ind}        if isinstance(_row, list):\n')
                new_lines.append(f'{ind}            for _ci, _v in enumerate(_row):\n')
                new_lines.append(f'{ind}                {sheet}.cell(row={r1}+_ri, column={c1}+_ci).value = _v\n')
                new_lines.append(f'{ind}        else:\n')
                new_lines.append(f'{ind}            {sheet}.cell(row={r1}+_ri, column={c1}).value = _row\n')
                converted = True
                break

            # var.value (read - just accessing the property)
            if f'{var_name}.value' in s and '=' not in s.split('.value')[1][:3]:
                # This is a read - convert to list comprehension
                new_lines.append(f'{ind}# Read range ({r1},{c1})-({r2},{c2}) - convert to list\n')
                new_lines.append(f'{ind}{s.replace(f"{var_name}.value", f"[[{sheet}.cell(row=_r, column=_c).value for _c in range({c1}, {c2}+1)] for _r in range({r1}, {r2}+1)]")}\n')
                converted = True
                break

            # var.merge() → sheet.merge_cells()
            if f'{var_name}.merge()' in s:
                new_lines.append(f'{ind}{sheet}.merge_cells(start_row={r1}, start_column={c1}, end_row={r2}, end_column={c2})\n')
                converted = True
                break

        if not converted:
            new_lines.append(line)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)

    remaining = sum(1 for l in new_lines if '# range ref as' in l)
    print(f"Converted range refs. Remaining 'range ref' markers: {remaining}")

if __name__ == '__main__':
    filepath = r"C:\Users\jerju\OneDrive\Focus CFO\AI\DNA Model\webapp\desktop_app_365.py"
    fix_rangerefs(filepath)
