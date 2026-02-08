"""
Script to fix font-overwriting patterns in desktop_app_365.py.
In openpyxl, Font objects are immutable, so each `.font = Font(...)` assignment
replaces the entire font object. This script consolidates consecutive font
assignments to the same cell into a single Font() call.
"""

import re

def fix_font_overwrites(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # Pattern to match: something.font = Font(...)
    font_pattern = re.compile(
        r'^(\s+)'                    # Leading whitespace (capture group 1)
        r'(.+?)'                     # Cell reference (capture group 2)
        r'\.font\s*=\s*'            # .font =
        r'(?:Font\((.+?)\)'         # Font(...) - capture args (group 3)
        r'|'
        r'(FONT_\w+))'              # or FONT_XXX constant (group 4)
        r'\s*$'                      # End of line
    )

    # Also match the constant reference pattern
    font_const_pattern = re.compile(
        r'^(\s+)(.+?)\.font\s*=\s*(FONT_\w+)\s*$'
    )

    result = []
    i = 0
    fixes = 0

    while i < len(lines):
        line = lines[i]

        # Try to match a font assignment
        m = font_pattern.match(line)
        if not m:
            result.append(line)
            i += 1
            continue

        indent = m.group(1)
        cell_ref = m.group(2).strip()

        # Check if it's a FONT_XXX constant
        if m.group(4):
            # It's a constant like FONT_DARK_BLUE - can't merge, skip
            result.append(line)
            i += 1
            continue

        font_args_str = m.group(3)

        # Collect all consecutive font assignments to the SAME cell
        group_lines = [(i, line, font_args_str)]
        j = i + 1

        while j < len(lines):
            next_line = lines[j]
            nm = font_pattern.match(next_line)
            if nm and nm.group(2).strip() == cell_ref and nm.group(3):
                group_lines.append((j, next_line, nm.group(3)))
                j += 1
            elif next_line.strip() == '' or next_line.strip().startswith('#'):
                # Skip blank lines and comments between font assignments
                # But only if the next non-blank line is also a font assignment to same cell
                k = j + 1
                while k < len(lines) and (lines[k].strip() == '' or lines[k].strip().startswith('#')):
                    k += 1
                if k < len(lines):
                    km = font_pattern.match(lines[k])
                    if km and km.group(2).strip() == cell_ref and km.group(3):
                        group_lines.append((j, next_line, None))  # Keep blank/comment
                        j += 1
                        continue
                break
            else:
                break

        if len(group_lines) <= 1:
            # Only one font assignment, no merging needed
            result.append(line)
            i += 1
            continue

        # Merge all font args into one
        all_kwargs = {}
        for _, _, args_str in group_lines:
            if args_str is None:
                continue  # blank line/comment
            # Parse the Font() arguments
            # Handle: name='Calibri Light', size=16, bold=True, color="FFFFFF", etc.
            for arg in _split_font_args(args_str):
                arg = arg.strip()
                if '=' in arg:
                    key, val = arg.split('=', 1)
                    all_kwargs[key.strip()] = val.strip()

        # Build merged Font() call
        merged_args = ', '.join(f'{k}={v}' for k, v in all_kwargs.items())
        merged_line = f'{indent}{cell_ref}.font = Font({merged_args})\n'

        result.append(merged_line)
        fixes += 1

        # Skip all the lines we merged
        i = j

    print(f"Fixed {fixes} font-overwriting groups")

    with open(filepath, 'w', encoding='utf-8') as f:
        f.writelines(result)

    print(f"Saved updated file: {filepath}")


def _split_font_args(args_str):
    """Split Font() arguments respecting nested parentheses and quotes."""
    result = []
    depth = 0
    current = ''
    in_string = False
    string_char = None

    for ch in args_str:
        if in_string:
            current += ch
            if ch == string_char:
                in_string = False
        elif ch in ('"', "'"):
            in_string = True
            string_char = ch
            current += ch
        elif ch == '(':
            depth += 1
            current += ch
        elif ch == ')':
            depth -= 1
            current += ch
        elif ch == ',' and depth == 0:
            result.append(current)
            current = ''
        else:
            current += ch

    if current.strip():
        result.append(current)

    return result


if __name__ == '__main__':
    import sys
    filepath = sys.argv[1] if len(sys.argv) > 1 else 'desktop_app_365.py'
    fix_font_overwrites(filepath)
