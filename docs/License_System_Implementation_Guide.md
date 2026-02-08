# License/Activation System Implementation Guide

## Overview

This document describes a simple license code system for Python desktop applications. It provides:
- First-run activation requirement
- Time-limited codes for beta testers
- Permanent codes for production users
- Per-machine licensing (code doesn't travel with the EXE)

## How It Works

1. **First Run**: App checks for license file in user's AppData
2. **No License**: Shows activation dialog requiring a code
3. **Valid Code**: Saves expiration date to license file, app opens
4. **Subsequent Runs**: Checks if license expired, opens directly if valid
5. **Expired**: Shows dialog asking for extension code

## Code Algorithm

Codes are generated using XOR obfuscation + base36 encoding + checksum:

```
Date (YYYYMMDD) → XOR with secret key → Base36 → Add checksum → Format as CFO-XXXX-XXX
```

Example: March 31, 2026 → `20260331` → XOR `0x4F43` → `C2MJ...` → `CFO-C2MJ-RM`

---

## Implementation

### Step 1: Add Constants

```python
# At top of your main app file
from datetime import datetime
import os

# License Settings - CHANGE THESE FOR YOUR PROJECT
APP_VERSION = "1.0.0"
DEFAULT_EXPIRATION = "2026-02-23"  # Default expiration if no code entered
LICENSE_SECRET_KEY = 0x4F43  # Change this to your own hex value (keep secret!)
```

### Step 2: Add LicenseManager Class

```python
class LicenseManager:
    """Simple license/expiration system"""

    LICENSE_FILE = "app_license.dat"  # Change filename for your app

    @staticmethod
    def get_license_path():
        """Get path to license file in user's AppData"""
        appdata = os.environ.get('APPDATA', os.path.expanduser('~'))
        license_dir = os.path.join(appdata, 'YourAppName')  # Change folder name
        if not os.path.exists(license_dir):
            os.makedirs(license_dir)
        return os.path.join(license_dir, LicenseManager.LICENSE_FILE)

    @staticmethod
    def generate_code(year, month, day):
        """Generate a license code for a given expiration date"""
        date_int = year * 10000 + month * 100 + day
        encoded = date_int ^ LICENSE_SECRET_KEY
        code = LicenseManager._to_base36(encoded)
        checksum = sum(ord(c) for c in code) % 26
        code = code + chr(65 + checksum)
        return f"CFO-{code[:4]}-{code[4:]}"  # Change prefix for your app

    @staticmethod
    def validate_code(code):
        """Validate code and return (year, month, day) or None if invalid"""
        try:
            code = code.upper().replace("CFO-", "").replace("-", "")
            if len(code) < 5:
                return None

            checksum_char = code[-1]
            code_body = code[:-1]

            expected_checksum = sum(ord(c) for c in code_body) % 26
            if chr(65 + expected_checksum) != checksum_char:
                return None

            encoded = LicenseManager._from_base36(code_body)
            date_int = encoded ^ LICENSE_SECRET_KEY

            year = date_int // 10000
            month = (date_int % 10000) // 100
            day = date_int % 100

            # Validation ranges
            if year < 2024 or year > 2099:
                return None
            if month < 1 or month > 12:
                return None
            if day < 1 or day > 31:
                return None

            return (year, month, day)
        except:
            return None

    @staticmethod
    def _to_base36(num):
        chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        result = ""
        while num > 0:
            result = chars[num % 36] + result
            num //= 36
        return result or "0"

    @staticmethod
    def _from_base36(s):
        return int(s, 36)

    @staticmethod
    def get_expiration_date():
        """Get current expiration date from license file or default"""
        license_path = LicenseManager.get_license_path()
        if os.path.exists(license_path):
            try:
                with open(license_path, 'r') as f:
                    parts = f.read().strip().split('-')
                    if len(parts) == 3:
                        return datetime(int(parts[0]), int(parts[1]), int(parts[2]))
            except:
                pass
        parts = DEFAULT_EXPIRATION.split('-')
        return datetime(int(parts[0]), int(parts[1]), int(parts[2]))

    @staticmethod
    def save_expiration_date(year, month, day):
        """Save expiration date to license file"""
        license_path = LicenseManager.get_license_path()
        with open(license_path, 'w') as f:
            f.write(f"{year}-{month:02d}-{day:02d}")

    @staticmethod
    def is_first_run():
        """Check if this is first run (no license file)"""
        return not os.path.exists(LicenseManager.get_license_path())

    @staticmethod
    def is_expired():
        """Check if license has expired"""
        return datetime.now() > LicenseManager.get_expiration_date()

    @staticmethod
    def days_remaining():
        """Get days until expiration"""
        delta = LicenseManager.get_expiration_date() - datetime.now()
        return max(0, delta.days)

    @staticmethod
    def extend_with_code(code):
        """Try to extend with a code. Returns True if successful."""
        result = LicenseManager.validate_code(code)
        if result:
            year, month, day = result
            LicenseManager.save_expiration_date(year, month, day)
            return True
        return False
```

### Step 3: Add License Check to App Initialization

In your main app class `__init__`:

```python
def __init__(self, root):
    self.root = root

    # Check license BEFORE creating UI
    if not self._check_license():
        self.root.destroy()
        return

    # ... rest of your initialization
    self._create_ui()
```

### Step 4: Add License Check Methods

```python
def _check_license(self):
    """Check license on startup. Returns True if valid."""

    # First run - show activation
    if LicenseManager.is_first_run():
        return self._show_license_dialog(is_first_run=True)

    # Warn if expiring soon
    days_left = LicenseManager.days_remaining()
    if 0 < days_left <= 7:
        exp_date = LicenseManager.get_expiration_date()
        messagebox.showwarning(
            "License Expiring Soon",
            f"Expires on {exp_date.strftime('%B %d, %Y')}.\n"
            f"{days_left} day(s) remaining."
        )
        return True

    # Check if expired
    if not LicenseManager.is_expired():
        return True

    # Expired - show dialog
    return self._show_license_dialog(is_first_run=False)

def _show_license_dialog(self, is_first_run=False):
    """Show activation/extension dialog. Returns True if successful."""
    dialog = tk.Toplevel(self.root)
    dialog.title("Product Activation" if is_first_run else "License Expired")
    dialog.geometry("450x280")
    dialog.resizable(False, False)
    dialog.transient(self.root)
    dialog.grab_set()

    # Center on screen
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() - 450) // 2
    y = (dialog.winfo_screenheight() - 280) // 2
    dialog.geometry(f"450x280+{x}+{y}")

    result = {'continue': False}

    main_frame = ttk.Frame(dialog, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)

    # Title
    title = "Welcome! Please activate." if is_first_run else "License Expired"
    ttk.Label(main_frame, text=title, font=('Segoe UI', 14, 'bold')).pack(pady=(0, 15))

    # Code entry
    code_frame = ttk.Frame(main_frame)
    code_frame.pack(pady=(0, 10))
    ttk.Label(code_frame, text="License Code:").pack(side=tk.LEFT)
    code_var = tk.StringVar()
    code_entry = ttk.Entry(code_frame, textvariable=code_var, width=18, font=('Consolas', 11))
    code_entry.pack(side=tk.LEFT, padx=(10, 0))

    # Status
    status_label = ttk.Label(main_frame, text="", foreground='red')
    status_label.pack(pady=(5, 15))

    def try_activate():
        code = code_var.get().strip()
        if not code:
            status_label.config(text="Please enter a code")
            return
        if LicenseManager.extend_with_code(code):
            exp = LicenseManager.get_expiration_date()
            messagebox.showinfo("Success", f"Valid until {exp.strftime('%B %d, %Y')}")
            result['continue'] = True
            dialog.destroy()
        else:
            status_label.config(text="Invalid code")

    def cancel():
        result['continue'] = False
        dialog.destroy()

    # Buttons
    btn_frame = ttk.Frame(main_frame)
    btn_frame.pack()
    ttk.Button(btn_frame, text="Activate", command=try_activate).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="Exit", command=cancel).pack(side=tk.LEFT, padx=5)

    dialog.protocol("WM_DELETE_WINDOW", cancel)
    code_entry.focus_set()

    self.root.wait_window(dialog)
    return result['continue']
```

---

## Code Generator Script

Create a separate file `generate_license_code.py` (KEEP PRIVATE):

```python
"""License Code Generator - KEEP THIS FILE PRIVATE"""

LICENSE_SECRET_KEY = 0x4F43  # Must match your app!

def _to_base36(num):
    chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    result = ""
    while num > 0:
        result = chars[num % 36] + result
        num //= 36
    return result or "0"

def generate_code(year, month, day):
    date_int = year * 10000 + month * 100 + day
    encoded = date_int ^ LICENSE_SECRET_KEY
    code = _to_base36(encoded)
    checksum = sum(ord(c) for c in code) % 26
    code = code + chr(65 + checksum)
    return f"CFO-{code[:4]}-{code[4:]}"  # Change prefix

if __name__ == "__main__":
    from datetime import datetime, timedelta

    print("LICENSE CODE GENERATOR")
    print("=" * 40)

    today = datetime.now()

    # Common codes
    codes = [
        ("+1 month", today + timedelta(days=30)),
        ("+3 months", today + timedelta(days=90)),
        ("+6 months", today + timedelta(days=180)),
        ("Permanent", datetime(2099, 12, 31)),
    ]

    for label, date in codes:
        code = generate_code(date.year, date.month, date.day)
        print(f"{label:12} ({date.strftime('%Y-%m-%d')}): {code}")
```

---

## Customization Checklist

When using in a new project, change:

1. `LICENSE_SECRET_KEY` - Use a different hex value
2. `DEFAULT_EXPIRATION` - Set your default expiration date
3. `LICENSE_FILE` - Change the filename
4. License folder name in `get_license_path()` - Use your app name
5. Code prefix (e.g., "CFO-" → "MYAPP-") in `generate_code()` and `validate_code()`
6. Dialog text and styling to match your app

---

## Security Notes

- **Not cryptographically secure** - This deters casual copying, not determined hackers
- **Keep secret key private** - Don't share `LICENSE_SECRET_KEY` or code generator
- **Per-machine licensing** - License file in AppData means each PC needs activation
- **EXE can be shared** - But won't run without a valid code

---

## Quick Reference

| Code Type | Expiration | Use For |
|-----------|------------|---------|
| Time-limited | 1-6 months | Beta testers |
| Permanent | 2099-12-31 | Production/paying customers |

Generate codes:
```bash
python generate_license_code.py
```

Or in Python:
```python
code = generate_code(2026, 6, 30)  # June 30, 2026
print(code)
```
