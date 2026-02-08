"""
Cleanup script: Remove all dangling keyword arguments and closing parens
left behind by the .api conversion. These are fragments from calls like:
    sheet.api.Cells.Replace(
        What=...,
        Replacement=...,
        MatchCase=False
    )
Where the first line was replaced with 'pass' but the kwargs and ) were left.

Also handles dangling Validation.Add, Hyperlinks.Add fragments.
"""
import re

def cleanup(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    new_lines = []
    skip_until_close = False
    paren_depth = 0
    fixes = 0
    i = 0

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # If we're in skip mode (removing dangling kwargs)
        if skip_until_close:
            # Count parens to find the matching close
            for ch in stripped:
                if ch == '(':
                    paren_depth += 1
                elif ch == ')':
                    if paren_depth > 0:
                        paren_depth -= 1
                    else:
                        skip_until_close = False
                        break
            i += 1
            fixes += 1
            continue

        # Detect dangling kwargs after 'pass' comments
        # Pattern: 'pass  # Cells.Replace...' followed by What=..., Replacement=..., etc.
        if stripped.startswith('pass') and ('Cells.Replace' in stripped or 'handled' in stripped):
            # Check if next line is a dangling kwarg
            if i + 1 < len(lines):
                next_s = lines[i + 1].strip()
                if next_s.startswith('What=') or next_s.startswith('Anchor=') or \
                   next_s.startswith('Address=') or next_s.startswith('Formula1=') or \
                   next_s.startswith('Type=') or next_s.startswith('Replacement='):
                    # Keep the pass line, skip the kwargs until closing )
                    new_lines.append(line)
                    skip_until_close = True
                    paren_depth = 0
                    i += 1
                    fixes += 1
                    continue

        # Detect standalone dangling kwargs (not after pass)
        if stripped.startswith(('What=', 'Replacement=', 'LookAt=', 'MatchCase=',
                                'Anchor=', 'Address=', 'SubAddress=', 'TextToDisplay=',
                                'Formula1=', 'AlertStyle=', 'Type=')):
            # Check if this is inside a function call (has matching open paren before)
            # or if it's a dangling fragment
            # Look backward to see if there's an unclosed paren
            paren_check = 0
            is_dangling = True
            for prev_line in reversed(new_lines[-5:]):
                for ch in prev_line:
                    if ch == '(':
                        paren_check += 1
                    elif ch == ')':
                        paren_check -= 1
                if paren_check > 0:
                    is_dangling = False
                    break

            if is_dangling:
                # Skip this line and any following kwargs + closing paren
                # But check if this is part of a legitimate function call
                i += 1
                fixes += 1
                continue

        # Handle orphan closing parens
        if stripped == ')' or stripped == '),':
            # Check if this is a valid close by looking at previous non-empty lines
            # Count open/close parens in recent context
            context_parens = 0
            for j in range(max(0, len(new_lines) - 10), len(new_lines)):
                for ch in new_lines[j]:
                    if ch == '(':
                        context_parens += 1
                    elif ch == ')':
                        context_parens -= 1
            if context_parens <= 0:
                # This is an orphan - skip it
                i += 1
                fixes += 1
                continue

        # Also fix Validation.Add dangling fragments
        # Pattern: "# Validation handled..." followed by "Formula1=..." lines
        if stripped.startswith('# Validation handled') or stripped.startswith('# DataValidation'):
            if i + 1 < len(lines):
                next_s = lines[i + 1].strip()
                if next_s.startswith('Type=') or next_s.startswith('Formula1='):
                    new_lines.append(line)
                    skip_until_close = True
                    paren_depth = 0
                    i += 1
                    fixes += 1
                    continue

        new_lines.append(line)
        i += 1

    with open(filepath, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)

    print(f"Cleanup complete! Removed {fixes} dangling fragments")

if __name__ == '__main__':
    filepath = r"C:\Users\jerju\OneDrive\Focus CFO\AI\DNA Model\webapp\desktop_app_365.py"
    cleanup(filepath)
