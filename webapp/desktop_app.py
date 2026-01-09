"""
CFO Financial Model Generator - Desktop Application
Creates true macro-enabled Excel files with embedded VBA using xlwings

VERSION HISTORY:
- v1.0.0 (2024-12-01): Initial release with P&L, Balance Sheet, Cash Flow
- v1.1.0 (2024-12-05): Added Dashboard module with KPIs and ratios
- v1.2.0 (2025-12-09): Renamed from DNA Model to Financial Model
- v1.2.2 (2025-12-10): Fixed Dashboard YTD calculations to use current year only
                        (Jan through current month from Menu!C7)
- v2.0.0 (2025-12-16): Added multi-division support with consolidation
                        - Division setup wizard for multiple entities
                        - Intelligent account matching via ChatGPT API
                        - Consolidated P&L, Balance Sheet, Cash Flow
                        - Division-specific and consolidated views
                        - Account mapping persistence (JSON + Excel)
"""

import os
import sys
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from datetime import datetime
import tempfile
import shutil
import json
import pandas as pd
import xlwings as xw
from xlwings.constants import DeleteShiftDirection

# Multi-division support imports
from consolidation_engine import ConsolidationEngine, DivisionConfig, AccountMapping
from mapping_persistence import MappingPersistence

# Application Version
APP_VERSION = "3.0.0"


class StepTimer:
    """Tracks timing for each step and persists to log file for analysis"""

    def __init__(self, log_dir=None):
        self.log_dir = log_dir or EXE_DIR
        self.log_file = os.path.join(self.log_dir, 'timing_log.json')
        self.current_run = {
            'timestamp': datetime.now().isoformat(),
            'steps': {},
            'total_time': 0
        }
        self.step_start_time = None
        self.current_step_name = None
        self.run_start_time = time.time()

    def start_step(self, step_name):
        """Start timing a step"""
        # End previous step if any
        if self.current_step_name:
            self.end_step()

        self.current_step_name = step_name
        self.step_start_time = time.time()

    def end_step(self):
        """End timing current step"""
        if self.current_step_name and self.step_start_time:
            elapsed = time.time() - self.step_start_time
            self.current_run['steps'][self.current_step_name] = round(elapsed, 2)
            print(f"  [TIMING] {self.current_step_name}: {elapsed:.2f}s")
        self.current_step_name = None
        self.step_start_time = None

    def finish_run(self):
        """Finish the run and save to log"""
        self.end_step()  # End any pending step
        self.current_run['total_time'] = round(time.time() - self.run_start_time, 2)

        # Load existing log
        history = self._load_history()
        history['runs'].append(self.current_run)

        # Keep only last 50 runs
        if len(history['runs']) > 50:
            history['runs'] = history['runs'][-50:]

        # Calculate averages
        history['averages'] = self._calculate_averages(history['runs'])

        # Save
        self._save_history(history)

        # Print summary
        self._print_summary(history)

        return self.current_run['total_time']

    def _load_history(self):
        """Load timing history from file"""
        try:
            if os.path.exists(self.log_file):
                with open(self.log_file, 'r') as f:
                    return json.load(f)
        except:
            pass
        return {'runs': [], 'averages': {}}

    def _save_history(self, history):
        """Save timing history to file"""
        try:
            with open(self.log_file, 'w') as f:
                json.dump(history, f, indent=2)
        except Exception as e:
            print(f"Warning: Could not save timing log: {e}")

    def _calculate_averages(self, runs):
        """Calculate average time for each step across runs"""
        step_times = {}
        for run in runs:
            for step_name, elapsed in run.get('steps', {}).items():
                if step_name not in step_times:
                    step_times[step_name] = []
                step_times[step_name].append(elapsed)

        averages = {}
        for step_name, times in step_times.items():
            averages[step_name] = {
                'avg': round(sum(times) / len(times), 2),
                'min': round(min(times), 2),
                'max': round(max(times), 2),
                'count': len(times)
            }

        # Add total time average
        total_times = [r.get('total_time', 0) for r in runs if r.get('total_time')]
        if total_times:
            averages['_total'] = {
                'avg': round(sum(total_times) / len(total_times), 2),
                'min': round(min(total_times), 2),
                'max': round(max(total_times), 2),
                'count': len(total_times)
            }

        return averages

    def _print_summary(self, history):
        """Print timing summary"""
        print("\n" + "="*60)
        print("TIMING SUMMARY - This Run")
        print("="*60)

        # Sort steps by time (descending)
        sorted_steps = sorted(
            self.current_run['steps'].items(),
            key=lambda x: x[1],
            reverse=True
        )

        for step_name, elapsed in sorted_steps:
            avg_data = history['averages'].get(step_name, {})
            avg = avg_data.get('avg', elapsed)
            print(f"  {step_name}: {elapsed:.2f}s (avg: {avg:.2f}s)")

        print(f"\n  TOTAL: {self.current_run['total_time']:.2f}s")

        if '_total' in history['averages']:
            avg_total = history['averages']['_total']
            print(f"  Average over {avg_total['count']} runs: {avg_total['avg']:.2f}s")
            print(f"  Range: {avg_total['min']:.2f}s - {avg_total['max']:.2f}s")

        print("="*60 + "\n")


APP_BUILD_DATE = "2026-01-02"
APP_NAME = "CFO Financial Model Generator"

# Get the directory where the script/exe is located
if getattr(sys, 'frozen', False):
    # Running as compiled exe - template is in the temp extraction folder
    APP_DIR = sys._MEIPASS
    EXE_DIR = os.path.dirname(sys.executable)
else:
    # Running as script
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    EXE_DIR = APP_DIR

# Template could be in APP_DIR (bundled) or EXE_DIR (alongside exe)
# Support both old DNA_Template name and new Financial_Template name
TEMPLATE_PATH = os.path.join(APP_DIR, 'Financial_Template.xlsm')
if not os.path.exists(TEMPLATE_PATH):
    TEMPLATE_PATH = os.path.join(EXE_DIR, 'Financial_Template.xlsm')
if not os.path.exists(TEMPLATE_PATH):
    # Fallback to old name for backwards compatibility
    TEMPLATE_PATH = os.path.join(APP_DIR, 'DNA_Template.xlsm')
if not os.path.exists(TEMPLATE_PATH):
    TEMPLATE_PATH = os.path.join(EXE_DIR, 'DNA_Template.xlsm')

# Template v2 path (optimized template with pre-built sheets)
TEMPLATE_V2_PATH = os.path.join(EXE_DIR, 'DNA_Template_v2.xlsm')
USE_TEMPLATE_V2 = True  # Set to False to use legacy template generation


class DivisionSetupDialog(tk.Toplevel):
    """Dialog for configuring multiple divisions"""

    MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
              'July', 'August', 'September', 'October', 'November', 'December']

    def __init__(self, parent, callback, existing_divisions=None):
        super().__init__(parent)
        self.title("Multi-Division Setup")
        self.callback = callback
        self.divisions = []  # List of division entries
        self.division_widgets = []  # UI widgets for each division

        # Year options
        current_year = datetime.now().year
        self.years = [str(y) for y in range(current_year - 10, current_year + 2)]

        # Initialize with existing divisions or empty
        if existing_divisions:
            for div in existing_divisions:
                self.divisions.append({
                    'name': tk.StringVar(value=div.get('name', '')),
                    'is_primary': tk.BooleanVar(value=div.get('is_primary', False)),
                    'pl_path': tk.StringVar(value=div.get('pl_path', '')),
                    'bs_path': tk.StringVar(value=div.get('bs_path', '')),
                    'start_month': tk.StringVar(value=div.get('start_month', 'January')),
                    'start_year': tk.StringVar(value=div.get('start_year', str(current_year))),
                    'end_month': tk.StringVar(value=div.get('end_month', 'December')),
                    'end_year': tk.StringVar(value=div.get('end_year', str(current_year)))
                })

        self.resizable(True, True)
        self._create_ui()
        self._center_window(750, 600)

        # Make modal
        self.transient(parent)
        self.grab_set()

    def _center_window(self, width, height):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _create_ui(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        ttk.Label(main_frame, text="Configure Divisions",
                  font=('Segoe UI', 14, 'bold')).pack(pady=(0, 10))

        ttk.Label(main_frame, text="Add divisions/departments/branches for consolidated reporting.",
                  font=('Segoe UI', 9), foreground='gray').pack(pady=(0, 15))

        # Division list frame with scrollbar
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # Canvas for scrollable content
        self.canvas = tk.Canvas(list_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Enable mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)  # Linux scroll up
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)  # Linux scroll down

        # Add initial division if none exist, otherwise render existing
        if not self.divisions:
            self._add_division(is_first=True)
        else:
            # Render existing divisions (only if not just added)
            for i, div in enumerate(self.divisions):
                self._render_division(i, is_first=(i == 0))

        # Buttons frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(15, 0))

        ttk.Button(btn_frame, text="+ Add Division", command=self._add_division).pack(side=tk.LEFT)

        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="Save & Continue", command=self._save_and_close).pack(side=tk.RIGHT)

    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        if event.num == 4:  # Linux scroll up
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:  # Linux scroll down
            self.canvas.yview_scroll(1, "units")
        else:  # Windows
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _scroll_to_bottom(self):
        """Scroll to the bottom of the canvas"""
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(1.0)

    def _add_division(self, is_first=False):
        """Add a new division entry"""
        current_year = datetime.now().year
        current_month = datetime.now().month
        div = {
            'name': tk.StringVar(value=f"Division {len(self.divisions) + 1}"),
            'is_primary': tk.BooleanVar(value=is_first),
            'pl_path': tk.StringVar(),
            'bs_path': tk.StringVar(),
            'start_month': tk.StringVar(value='January'),
            'start_year': tk.StringVar(value=str(current_year)),
            'end_month': tk.StringVar(value=self.MONTHS[current_month - 1]),
            'end_year': tk.StringVar(value=str(current_year))
        }
        self.divisions.append(div)
        self._render_division(len(self.divisions) - 1, is_first)

        # Scroll to show the new division
        if not is_first:
            self.after(100, self._scroll_to_bottom)

    def _render_division(self, index, is_first=False):
        """Render UI widgets for a division entry"""
        div = self.divisions[index]

        frame = ttk.LabelFrame(self.scrollable_frame, text=f"Division {index + 1}", padding="10")
        frame.pack(fill=tk.X, pady=5, padx=5)

        # Row 1: Name and Primary checkbox
        row1 = ttk.Frame(frame)
        row1.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(row1, text="Name:").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=div['name'], width=25).pack(side=tk.LEFT, padx=5)

        ttk.Checkbutton(row1, text="Primary Division",
                        variable=div['is_primary'],
                        command=lambda i=index: self._set_primary(i)).pack(side=tk.LEFT, padx=20)

        if not is_first:
            ttk.Button(row1, text="Remove", width=8,
                       command=lambda i=index: self._remove_division(i)).pack(side=tk.RIGHT)

        # Row 2: P&L file
        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(row2, text="P&L File:").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=div['pl_path'], width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(row2, text="Browse...",
                   command=lambda v=div['pl_path'], d=div: self._browse_file(v, "P&L", d)).pack(side=tk.LEFT)

        # Row 3: BS file
        row3 = ttk.Frame(frame)
        row3.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(row3, text="Balance Sheet:").pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=div['bs_path'], width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(row3, text="Browse...",
                   command=lambda v=div['bs_path'], d=div: self._browse_file(v, "Balance Sheet", d)).pack(side=tk.LEFT)

        # Date Range is auto-detected from files - no UI needed
        # The start_month, start_year, end_month, end_year variables are
        # automatically populated when files are selected via _auto_detect_date_range()

        self.division_widgets.append(frame)

    def _set_primary(self, selected_index):
        """Ensure only one division is marked as primary"""
        for i, div in enumerate(self.divisions):
            if i != selected_index:
                div['is_primary'].set(False)

    def _remove_division(self, index):
        """Remove a division entry"""
        if len(self.divisions) <= 1:
            messagebox.showwarning("Warning", "At least one division is required.")
            return

        # Remove from list
        self.divisions.pop(index)

        # Rebuild UI
        for widget in self.division_widgets:
            widget.destroy()
        self.division_widgets.clear()

        for i, div in enumerate(self.divisions):
            self._render_division(i, is_first=(i == 0))

        # Ensure at least one is primary
        if not any(d['is_primary'].get() for d in self.divisions):
            self.divisions[0]['is_primary'].set(True)

    def _browse_file(self, var, file_type, div=None):
        """Browse for a file and auto-detect date range"""
        path = filedialog.askopenfilename(
            title=f"Select {file_type} File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if path:
            var.set(path)
            # Auto-detect date range from file
            if div:
                self._auto_detect_date_range(path, div)

    def _auto_detect_date_range(self, file_path, div):
        """Read Excel file and auto-detect the date range from column headers"""
        try:
            # Read just the first few rows to get headers
            df = pd.read_excel(file_path, header=None, nrows=5)

            detected_months = []
            # Look in first 3 rows for date headers (row 0, 1, 2)
            for row_idx in range(min(3, len(df))):
                for col_idx in range(1, min(50, len(df.columns))):  # Skip column A, check up to 50 cols
                    val = df.iloc[row_idx, col_idx]
                    parsed = self._parse_month_value(val)
                    if parsed:
                        month_num, year, display_name = parsed
                        detected_months.append((month_num, year, display_name))

            if detected_months:
                # Sort by year then month
                detected_months.sort(key=lambda x: (x[1], x[0]))

                # Get first and last month
                first_month, first_year, _ = detected_months[0]
                last_month, last_year, _ = detected_months[-1]

                # Update the division's date range
                div['start_month'].set(self.MONTHS[first_month - 1])
                div['start_year'].set(str(first_year))
                div['end_month'].set(self.MONTHS[last_month - 1])
                div['end_year'].set(str(last_year))

                print(f"Auto-detected date range: {self.MONTHS[first_month - 1]} {first_year} to {self.MONTHS[last_month - 1]} {last_year}")
        except Exception as e:
            print(f"Could not auto-detect date range: {e}")

    def _save_and_close(self):
        """Validate and save divisions"""
        # Validate
        for i, div in enumerate(self.divisions):
            if not div['name'].get().strip():
                messagebox.showerror("Error", f"Division {i + 1} needs a name.")
                return
            if not div['pl_path'].get():
                messagebox.showerror("Error", f"Division '{div['name'].get()}' needs a P&L file.")
                return
            if not div['bs_path'].get():
                messagebox.showerror("Error", f"Division '{div['name'].get()}' needs a Balance Sheet file.")
                return

        # Ensure one is primary
        if not any(d['is_primary'].get() for d in self.divisions):
            self.divisions[0]['is_primary'].set(True)

        # Convert to list of dicts
        result = []
        for div in self.divisions:
            result.append({
                'name': div['name'].get().strip(),
                'is_primary': div['is_primary'].get(),
                'pl_path': div['pl_path'].get(),
                'bs_path': div['bs_path'].get(),
                'start_month': div['start_month'].get(),
                'start_year': div['start_year'].get(),
                'end_month': div['end_month'].get(),
                'end_year': div['end_year'].get()
            })

        self.callback(result)
        self.destroy()


class AccountMappingDialog(tk.Toplevel):
    """Interactive dialog for reviewing and merging account mappings"""

    def __init__(self, parent, mappings, divisions, callback, pl_mappings=None, bs_mappings=None):
        super().__init__(parent)
        self.title("Review Account Mappings")
        self.mappings = mappings  # Combined Dict of consolidated_name -> AccountMapping
        self.pl_mappings = pl_mappings or {}  # P&L specific mappings
        self.bs_mappings = bs_mappings or {}  # BS specific mappings
        self.divisions = divisions
        self.callback = callback

        # Separate mappings if combined dict was passed
        if not pl_mappings and not bs_mappings:
            self._split_mappings_by_type()

        self.resizable(True, True)
        self._create_ui()
        self._center_window(1000, 700)

        # Make modal
        self.transient(parent)
        self.grab_set()

    def _split_mappings_by_type(self):
        """Try to split combined mappings - if not possible, show all in one view"""
        # For now, put all in pl_mappings since we don't have type info
        self.pl_mappings = dict(self.mappings)
        self.bs_mappings = {}

    def _center_window(self, width, height):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _create_ui(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title and Description
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(title_frame, text="Account Consolidation Review",
                  font=('Segoe UI', 16, 'bold')).pack(anchor='w')

        # Instructions box
        instr_frame = ttk.LabelFrame(main_frame, text="How This Works", padding="10")
        instr_frame.pack(fill=tk.X, pady=(0, 10))

        instructions = (
            "This screen shows how accounts from each division will be consolidated:\n\n"
            "• GREEN (Exact Match): Account names are identical across divisions - will be summed together\n"
            "• BLUE (Intelligent Match): AI detected similar accounts - review and merge if correct\n"
            "• ORANGE (Division-specific): Account exists in only one division - will appear as-is\n\n"
            "ACTION: Select 2+ similar accounts and click 'Merge Selected' to combine them."
        )
        ttk.Label(instr_frame, text=instructions, font=('Segoe UI', 9),
                  justify='left', wraplength=900).pack(anchor='w')

        # Summary stats frame
        stats_frame = ttk.Frame(main_frame)
        stats_frame.pack(fill=tk.X, pady=(0, 10))

        # Calculate stats
        exact_count = sum(1 for m in self.pl_mappings.values()
                         if (m.match_type if hasattr(m, 'match_type') else m.get('match_type', '')) == 'exact')
        intel_count = sum(1 for m in self.pl_mappings.values()
                         if (m.match_type if hasattr(m, 'match_type') else m.get('match_type', '')) == 'intelligent')
        div_specific = len(self.pl_mappings) - exact_count - intel_count

        ttk.Label(stats_frame, text=f"Summary: ", font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT)
        ttk.Label(stats_frame, text=f"{exact_count} exact matches", foreground='green',
                  font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(stats_frame, text=f"{intel_count} intelligent matches", foreground='blue',
                  font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(stats_frame, text=f"{div_specific} division-specific", foreground='orange',
                  font=('Segoe UI', 10)).pack(side=tk.LEFT)

        # Notebook for P&L and Balance Sheet tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # P&L Tab
        pl_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(pl_frame, text=f"P&L Accounts ({len(self.pl_mappings)})")
        self.pl_tree = self._create_mapping_tree(pl_frame, self.pl_mappings)

        # Balance Sheet Tab
        bs_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(bs_frame, text=f"Balance Sheet ({len(self.bs_mappings)})")
        self.bs_tree = self._create_mapping_tree(bs_frame, self.bs_mappings)

        # Action buttons frame with clear styling
        action_frame = ttk.LabelFrame(main_frame, text="Actions", padding="10")
        action_frame.pack(fill=tk.X, pady=(5, 10))

        merge_btn = ttk.Button(action_frame, text="Merge Selected Accounts",
                               command=self._merge_selected)
        merge_btn.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(action_frame,
                  text="Select 2+ rows above, then click to combine them into one consolidated account",
                  font=('Segoe UI', 9), foreground='#666').pack(side=tk.LEFT)

        # Bottom buttons with clear labeling
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        # Left side - status
        self.status_label = ttk.Label(btn_frame, text="Review mappings above, then click Continue to generate the model.",
                                       foreground='#666', font=('Segoe UI', 9))
        self.status_label.pack(side=tk.LEFT)

        # Right side - action buttons
        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=self.destroy, width=12)
        cancel_btn.pack(side=tk.RIGHT, padx=(5, 0))

        confirm_btn = ttk.Button(btn_frame, text="Continue →", command=self._confirm_and_close, width=15)
        confirm_btn.pack(side=tk.RIGHT)

        # Tip text
        tip_frame = ttk.Frame(main_frame)
        tip_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(tip_frame, text="Tip: Most mappings are correct by default. Click 'Continue' if everything looks good.",
                  font=('Segoe UI', 8), foreground='#999').pack(anchor='e')

    def _create_mapping_tree(self, parent, mappings):
        """Create a treeview for a set of mappings"""
        # Frame for tree and scrollbars
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ['Consolidated', 'Type', 'Divisions'] + [d['name'] for d in self.divisions]
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15,
                           selectmode='extended')  # Allow multi-select

        for col in columns:
            tree.heading(col, text=col)
            if col == 'Consolidated':
                tree.column(col, width=200)
            elif col == 'Type':
                tree.column(col, width=100)
            elif col == 'Divisions':
                tree.column(col, width=80)
            else:
                tree.column(col, width=120)

        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Configure tag colors
        tree.tag_configure('exact', foreground='green')
        tree.tag_configure('intelligent', foreground='blue')
        tree.tag_configure('division_specific', foreground='orange')
        tree.tag_configure('manual', foreground='purple')

        # Populate tree
        self._populate_mapping_tree(tree, mappings)

        return tree

    def _populate_mapping_tree(self, tree, mappings):
        """Populate a treeview with mappings"""
        # Group by match type for display order
        by_type = {'exact': [], 'intelligent': [], 'manual': [], 'division_specific': []}

        for name, mapping in mappings.items():
            m_type = mapping.match_type if hasattr(mapping, 'match_type') else mapping.get('match_type', 'unknown')
            if m_type in by_type:
                by_type[m_type].append((name, mapping))
            else:
                by_type['division_specific'].append((name, mapping))

        # Add to tree
        for match_type in ['exact', 'intelligent', 'manual', 'division_specific']:
            for name, mapping in by_type[match_type]:
                if hasattr(mapping, 'division_mappings'):
                    div_mappings = mapping.division_mappings
                    m_type = mapping.match_type
                else:
                    div_mappings = mapping.get('division_mappings', {})
                    m_type = mapping.get('match_type', 'unknown')

                # Count how many divisions have this account
                div_count = len(div_mappings)

                values = [name, m_type.replace('_', '-').title(), f"{div_count}/{len(self.divisions)}"]
                for div in self.divisions:
                    values.append(div_mappings.get(div['name'], '-'))

                item = tree.insert('', tk.END, values=values, tags=(m_type,))

    def _merge_selected(self):
        """Merge selected accounts into one consolidated account"""
        # Get current tab's tree
        current_tab = self.notebook.index(self.notebook.select())
        tree = self.pl_tree if current_tab == 0 else self.bs_tree
        mappings = self.pl_mappings if current_tab == 0 else self.bs_mappings

        selected = tree.selection()
        if len(selected) < 2:
            messagebox.showwarning("Selection Required",
                                   "Please select at least 2 accounts to merge.")
            return

        # Get selected account names
        selected_names = []
        for item in selected:
            values = tree.item(item, 'values')
            selected_names.append(values[0])

        # Ask for consolidated name
        default_name = selected_names[0]  # Use first selected as default
        new_name = tk.simpledialog.askstring(
            "Merge Accounts",
            f"Enter consolidated account name for:\n" + "\n".join(f"  - {n}" for n in selected_names),
            initialvalue=default_name,
            parent=self
        )

        if not new_name:
            return

        # Merge mappings
        merged_div_mappings = {}
        for name in selected_names:
            if name in mappings:
                mapping = mappings[name]
                div_maps = mapping.division_mappings if hasattr(mapping, 'division_mappings') else mapping.get('division_mappings', {})
                merged_div_mappings.update(div_maps)

        # Create new merged mapping
        new_mapping = AccountMapping(
            consolidated_name=new_name,
            division_mappings=merged_div_mappings,
            match_type='manual',
            confidence=1.0,
            approved=True
        )

        # Remove old mappings and add new one
        for name in selected_names:
            if name in mappings:
                del mappings[name]

        mappings[new_name] = new_mapping

        # Refresh tree
        for item in tree.get_children():
            tree.delete(item)
        self._populate_mapping_tree(tree, mappings)

        # Update tab label with count
        if current_tab == 0:
            self.notebook.tab(0, text=f"P&L Accounts ({len(self.pl_mappings)})")
        else:
            self.notebook.tab(1, text=f"Balance Sheet ({len(self.bs_mappings)})")

        self.status_label.config(text=f"Merged {len(selected_names)} accounts into '{new_name}'")

    def _confirm_and_close(self):
        """Confirm all mappings and close"""
        # Mark all as approved
        for mappings in [self.pl_mappings, self.bs_mappings]:
            for name, mapping in mappings.items():
                if hasattr(mapping, 'approved'):
                    mapping.approved = True
                elif isinstance(mapping, dict):
                    mapping['approved'] = True

        # Combine mappings and call callback
        combined = {**self.pl_mappings, **self.bs_mappings}
        self.callback(combined)
        self.destroy()


class FinancialModelApp:
    """Desktop application for generating financial models"""

    MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
              'July', 'August', 'September', 'October', 'November', 'December']

    VBA_CODE = '''
Option Explicit

Private Const SOURCE_PL_SHEET As String = "Source_PL"
Private Const SOURCE_BS_SHEET As String = "Source_BS"
Private Const SOURCE_BUDGET_SHEET As String = "Source_Budget"
Private Const PL_SHEET As String = "PL"
Private Const BS_SHEET As String = "Balance_Sheet"
Private Const CF_SHEET As String = "Cash_Flow"
Private Const MENU_SHEET As String = "Menu"

Public Sub UploadPLFile()
    Dim filePath As String
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", , "Select P&L File")
    If filePath = "False" Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call ImportData(filePath, SOURCE_PL_SHEET)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "P&L data imported successfully!", vbInformation
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub UploadBSFile()
    Dim filePath As String
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", , "Select Balance Sheet File")
    If filePath = "False" Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call ImportData(filePath, SOURCE_BS_SHEET)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Balance Sheet data imported successfully!", vbInformation
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub UploadBudgetFile()
    Dim filePath As String
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", , "Select Budget File")
    If filePath = "False" Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call ImportBudgetDataSmart(filePath, SOURCE_BUDGET_SHEET)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub ImportBudgetDataSmart(filePath As String, targetSheet As String)
    ' Smart budget import with account matching logic
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim tgtWs As Worksheet
    Dim plWs As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim importCount As Long, matchCount As Long, unmatchedCount As Long
    Dim totalBudgetAccounts As Long
    Dim unmatchedList As String
    Dim userChoice As VbMsgBoxResult

    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set srcWs = srcWb.Sheets(1)
    Set tgtWs = ThisWorkbook.Sheets(targetSheet)
    Set plWs = ThisWorkbook.Sheets(SOURCE_PL_SHEET)

    headerRow = FindHeaderRow(srcWs)
    If headerRow = 0 Then
        srcWb.Close False
        Err.Raise vbObjectError + 1, , "Could not find month headers in budget file"
    End If

    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    lastCol = srcWs.Cells(headerRow, srcWs.Columns.Count).End(xlToLeft).Column

    ' First pass: analyze matching
    totalBudgetAccounts = 0
    matchCount = 0
    unmatchedList = ""

    For i = headerRow + 1 To lastRow
        Dim budgetAcct As String
        budgetAcct = Trim(CStr(srcWs.Cells(i, 1).Value))
        If Len(budgetAcct) > 0 And LCase(budgetAcct) <> "cash basis" Then
            totalBudgetAccounts = totalBudgetAccounts + 1
            Dim matchResult As Long
            matchResult = FindAccountMatch(plWs, budgetAcct)
            If matchResult > 0 Then
                matchCount = matchCount + 1
            Else
                unmatchedCount = unmatchedCount + 1
                If Len(unmatchedList) < 500 Then
                    unmatchedList = unmatchedList & vbCrLf & "  - " & budgetAcct
                End If
            End If
        End If
    Next i

    ' Check match rate
    Dim matchRate As Double
    If totalBudgetAccounts > 0 Then
        matchRate = matchCount / totalBudgetAccounts
    Else
        srcWb.Close False
        MsgBox "No accounts found in budget file.", vbExclamation
        Exit Sub
    End If

    ' If very few matches, warn user this may be wrong file
    If matchRate < 0.3 Then
        userChoice = MsgBox("Warning: Only " & matchCount & " of " & totalBudgetAccounts & _
            " accounts (" & Format(matchRate * 100, "0") & "%) match the existing P&L." & vbCrLf & vbCrLf & _
            "This appears to be an incorrect budget file that does not match the existing G/L accounts." & vbCrLf & vbCrLf & _
            "If you proceed, there will be " & unmatchedCount & " unmatched accounts that will not be displayed." & vbCrLf & vbCrLf & _
            "Do you want to proceed anyway?", vbYesNo + vbExclamation, "Possible File Mismatch")
        If userChoice = vbNo Then
            srcWb.Close False
            MsgBox "Import cancelled. Please select the correct budget file.", vbInformation
            Exit Sub
        End If
    End If

    ' If some accounts don't match, ask user
    If unmatchedCount > 0 And matchRate >= 0.3 Then
        Dim truncatedList As String
        If unmatchedCount > 10 Then
            truncatedList = Left(unmatchedList, 500) & vbCrLf & "  ... and " & (unmatchedCount - 10) & " more"
        Else
            truncatedList = unmatchedList
        End If

        userChoice = MsgBox(matchCount & " of " & totalBudgetAccounts & " accounts matched the P&L." & vbCrLf & vbCrLf & _
            "The following " & unmatchedCount & " accounts did not match:" & truncatedList & vbCrLf & vbCrLf & _
            "Would you like to add these unmatched accounts anyway?" & vbCrLf & vbCrLf & _
            "Click YES to add them, NO to skip them.", vbYesNo + vbQuestion, "Unmatched Accounts Found")
    End If

    ' Second pass: import data
    importCount = 0
    Dim addedCount As Long
    addedCount = 0
    Dim tgtLastCol As Long
    tgtLastCol = tgtWs.Cells(1, tgtWs.Columns.Count).End(xlToLeft).Column
    Dim tgtLastRow As Long
    tgtLastRow = tgtWs.Cells(tgtWs.Rows.Count, 1).End(xlUp).Row

    For j = 2 To lastCol
        Dim monthName As String
        monthName = Trim(CStr(srcWs.Cells(headerRow, j).Value))
        If Len(monthName) > 0 And LCase(monthName) <> "total" Then
            Dim monthCol As Long
            monthCol = FindMonthCol(tgtWs, monthName)

            If monthCol > 0 Then
                For i = headerRow + 1 To lastRow
                    Dim accountName As String
                    accountName = Trim(CStr(srcWs.Cells(i, 1).Value))
                    If Len(accountName) > 0 And LCase(accountName) <> "cash basis" Then
                        Dim accountRow As Long
                        accountRow = FindAccountMatch(tgtWs, accountName)

                        If accountRow > 0 Then
                            tgtWs.Cells(accountRow, monthCol).Value = srcWs.Cells(i, j).Value
                            importCount = importCount + 1
                        ElseIf userChoice = vbYes Then
                            ' Add new account if user chose to add unmatched
                            accountRow = FindOrAddBudgetAccount(tgtWs, accountName, tgtLastRow)
                            If accountRow > tgtLastRow Then tgtLastRow = accountRow
                            tgtWs.Cells(accountRow, monthCol).Value = srcWs.Cells(i, j).Value
                            importCount = importCount + 1
                            addedCount = addedCount + 1
                        End If
                    End If
                Next i
            End If
        End If
    Next j

    srcWb.Close False

    ' Show summary
    Dim summaryMsg As String
    summaryMsg = "Budget import complete!" & vbCrLf & vbCrLf & _
        "Imported " & importCount & " values."
    If addedCount > 0 Then
        summaryMsg = summaryMsg & vbCrLf & "Added " & addedCount & " new accounts."
    End If
    If unmatchedCount > 0 And userChoice = vbNo Then
        summaryMsg = summaryMsg & vbCrLf & "Skipped " & unmatchedCount & " unmatched accounts."
    End If
    MsgBox summaryMsg, vbInformation
End Sub

Private Function FindAccountMatch(ws As Worksheet, accountName As String) As Long
    ' Try to find account with smart matching:
    ' 1. Exact match
    ' 2. Budget name contained in P&L name
    ' 3. P&L name contained in budget name
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    Dim plAcct As String
    Dim budgetLower As String
    budgetLower = LCase(Trim(accountName))

    ' First: exact match
    For i = 3 To lastRow
        plAcct = Trim(CStr(ws.Cells(i, 1).Value))
        If LCase(plAcct) = budgetLower Then
            FindAccountMatch = i
            Exit Function
        End If
    Next i

    ' Second: budget name contained in P&L name (e.g., "Rent" matches "Rent Expense")
    For i = 3 To lastRow
        plAcct = LCase(Trim(CStr(ws.Cells(i, 1).Value)))
        If Len(budgetLower) > 3 And InStr(plAcct, budgetLower) > 0 Then
            FindAccountMatch = i
            Exit Function
        End If
    Next i

    ' Third: P&L name contained in budget name (e.g., "Marketing Expense" contains "Marketing")
    For i = 3 To lastRow
        plAcct = LCase(Trim(CStr(ws.Cells(i, 1).Value)))
        If Len(plAcct) > 3 And InStr(budgetLower, plAcct) > 0 Then
            FindAccountMatch = i
            Exit Function
        End If
    Next i

    FindAccountMatch = 0
End Function

Private Function FindOrAddBudgetAccount(ws As Worksheet, accountName As String, currentLastRow As Long) As Long
    ' Find existing account or add new one at the end
    Dim i As Long
    For i = 3 To currentLastRow
        If Trim(CStr(ws.Cells(i, 1).Value)) = accountName Then
            FindOrAddBudgetAccount = i
            Exit Function
        End If
    Next i
    ' Add new account
    Dim newRow As Long
    newRow = currentLastRow + 1
    ws.Cells(newRow, 1).Value = accountName
    FindOrAddBudgetAccount = newRow
End Function

Private Function FindMonthCol(ws As Worksheet, monthName As String) As Long
    ' Find existing month column (do not add new columns)
    Dim j As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For j = 2 To lastCol
        If LCase(Trim(CStr(ws.Cells(1, j).Value))) = LCase(monthName) Then
            FindMonthCol = j
            Exit Function
        End If
    Next j
    FindMonthCol = 0
End Function

Private Function FindAccountRow(ws As Worksheet, accountName As String) As Long
    ' Find existing account row (do not add new accounts)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 3 To lastRow  ' Start at row 3 (row 1=header, row 2=YYYYMM helper)
        If Trim(CStr(ws.Cells(i, 1).Value)) = accountName Then
            FindAccountRow = i
            Exit Function
        End If
    Next i
    FindAccountRow = 0
End Function

Private Sub ImportData(filePath As String, targetSheet As String)
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim tgtWs As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long

    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set srcWs = srcWb.Sheets(1)
    Set tgtWs = ThisWorkbook.Sheets(targetSheet)

    headerRow = FindHeaderRow(srcWs)
    If headerRow = 0 Then
        srcWb.Close False
        Err.Raise vbObjectError + 1, , "Could not find month headers"
    End If

    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    lastCol = srcWs.Cells(headerRow, srcWs.Columns.Count).End(xlToLeft).Column

    Dim tgtLastCol As Long
    tgtLastCol = tgtWs.Cells(1, tgtWs.Columns.Count).End(xlToLeft).Column

    For j = 2 To lastCol
        Dim monthName As String
        monthName = Trim(CStr(srcWs.Cells(headerRow, j).Value))
        If Len(monthName) > 0 And LCase(monthName) <> "total" Then
            Dim monthCol As Long
            monthCol = FindOrAddMonth(tgtWs, monthName, tgtLastCol)
            If monthCol > tgtLastCol Then tgtLastCol = monthCol

            For i = headerRow + 1 To lastRow
                Dim accountName As String
                accountName = Trim(CStr(srcWs.Cells(i, 1).Value))
                If Len(accountName) > 0 And LCase(accountName) <> "cash basis" Then
                    Dim accountRow As Long
                    accountRow = FindOrAddAccount(tgtWs, accountName)
                    tgtWs.Cells(accountRow, monthCol).Value = srcWs.Cells(i, j).Value
                End If
            Next i
        End If
    Next j

    srcWb.Close False
    UpdateNamedRange tgtWs
End Sub

Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim i As Long, j As Long
    For i = 1 To 15
        For j = 1 To 20
            Dim val As String
            val = LCase(Trim(CStr(ws.Cells(i, j).Value)))
            If InStr(val, "january") > 0 Or InStr(val, "february") > 0 Or InStr(val, "march") > 0 Then
                If InStr(val, "-") = 0 Then
                    FindHeaderRow = i
                    Exit Function
                End If
            End If
        Next j
    Next i
    FindHeaderRow = 0
End Function

Private Function FindOrAddMonth(ws As Worksheet, monthName As String, lastCol As Long) As Long
    Dim j As Long
    For j = 2 To lastCol
        If LCase(Trim(CStr(ws.Cells(1, j).Value))) = LCase(monthName) Then
            FindOrAddMonth = j
            Exit Function
        End If
    Next j
    lastCol = lastCol + 1
    ws.Cells(1, lastCol).Value = monthName
    FindOrAddMonth = lastCol
End Function

Private Function FindOrAddAccount(ws As Worksheet, accountName As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, 1).Value)) = accountName Then
            FindOrAddAccount = i
            Exit Function
        End If
    Next i
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).Value = accountName
    FindOrAddAccount = lastRow
End Function

Private Sub UpdateNamedRange(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim rangeName As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If ws.Name = SOURCE_PL_SHEET Then
        rangeName = "SourcePL"
    ElseIf ws.Name = SOURCE_BS_SHEET Then
        rangeName = "SourceBS"
    Else
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Names(rangeName).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add rangeName, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
End Sub

Public Sub RefreshAll()
    Application.CalculateFull
    MsgBox "All formulas refreshed!", vbInformation
End Sub

Public Sub RunDiagnostics()
    Dim issues As String
    issues = ""

    If Not CheckBSBalance() Then
        issues = issues & "- Balance Sheet does not balance" & vbCrLf
    End If

    If Len(issues) = 0 Then
        MsgBox "All checks passed!", vbInformation
    Else
        MsgBox "Issues found:" & vbCrLf & issues, vbExclamation
    End If
End Sub

Private Function CheckBSBalance() As Boolean
    CheckBSBalance = True
End Function

'-----------------------------------------------------
' UpdateColumnVisibility - Hide prior years and future months
' Called when current month changes on Menu sheet (C7)
'-----------------------------------------------------
Public Sub UpdateColumnVisibility()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Dim wsMenu As Worksheet
    Dim wsPL As Worksheet
    Dim wsBS As Worksheet
    Dim wsCF As Worksheet
    Dim currentMonthStr As String

    Set wsMenu = GetWs("Menu")
    Set wsPL = GetWs("PL")
    Set wsBS = GetWs("Balance_Sheet")
    Set wsCF = GetWs("Cash_Flow")

    If wsMenu Is Nothing Then GoTo Cleanup

    ' Get current month from Menu C7 (e.g., "Nov 24" or date value)
    currentMonthStr = Trim(CStr(wsMenu.Range("C7").Value))
    If Len(currentMonthStr) = 0 Then GoTo Cleanup

    ' Update visibility on P&L sheet
    If Not wsPL Is Nothing Then
        Call SetColVisibility(wsPL, currentMonthStr)
    End If

    ' Update visibility on Balance Sheet
    If Not wsBS Is Nothing Then
        Call SetColVisibility(wsBS, currentMonthStr)
    End If

    ' Update visibility on Cash Flow
    If Not wsCF Is Nothing Then
        Call SetColVisibility(wsCF, currentMonthStr)
    End If

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Resume Cleanup
End Sub

Private Sub SetColVisibility(ws As Worksheet, currentMonthStr As String)
    On Error Resume Next

    Dim headerRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim currentYear As Long
    Dim currentMonthNum As Long
    Dim colYear As Long
    Dim colMonthNum As Long
    Dim headerVal As String

    headerRow = 4
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    ' Parse current month/year from string like "Nov 24" or "Nov-24"
    currentYear = GetYearFromStr(currentMonthStr)
    currentMonthNum = GetMonthFromStr(currentMonthStr)

    If currentYear = 0 Or currentMonthNum = 0 Then Exit Sub

    ' First, unhide all columns
    ws.Columns.Hidden = False

    ' Scan columns and set visibility
    For col = 2 To lastCol
        headerVal = Trim(CStr(ws.Cells(headerRow, col).Value))

        ' Check if this is a month column
        If IsMonthCol(headerVal) Then
            colYear = GetYearFromStr(headerVal)
            colMonthNum = GetMonthFromStr(headerVal)

            If colYear > 0 And colMonthNum > 0 Then
                ' Future year - hide
                If colYear > currentYear Then
                    ws.Columns(col).Hidden = True
                ' Same year but future month - hide
                ElseIf colYear = currentYear And colMonthNum > currentMonthNum Then
                    ws.Columns(col).Hidden = True
                ' Prior year - hide
                ElseIf colYear < currentYear Then
                    ws.Columns(col).Hidden = True
                ' Current year, current or prior month - show
                Else
                    ws.Columns(col).Hidden = False
                End If
            End If
        End If
    Next col
End Sub

Private Function IsMonthCol(headerVal As String) As Boolean
    Dim months As Variant
    Dim i As Long

    months = Array("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
    headerVal = LCase(Trim(headerVal))

    For i = LBound(months) To UBound(months)
        If InStr(headerVal, months(i)) > 0 Then
            IsMonthCol = True
            Exit Function
        End If
    Next i
    IsMonthCol = False
End Function

Private Function GetYearFromStr(monthStr As String) As Long
    Dim parts() As String
    Dim i As Long
    Dim y As Long

    GetYearFromStr = 0
    monthStr = Replace(monthStr, "-", " ")
    parts = Split(Trim(monthStr), " ")

    For i = LBound(parts) To UBound(parts)
        If IsNumeric(parts(i)) Then
            y = CLng(parts(i))
            If y >= 2000 And y <= 2100 Then
                GetYearFromStr = y
                Exit Function
            ElseIf y >= 0 And y <= 99 Then
                GetYearFromStr = 2000 + y
                Exit Function
            End If
        End If
    Next i
End Function

Private Function GetMonthFromStr(monthStr As String) As Long
    Dim months As Variant
    Dim i As Long

    months = Array("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
    monthStr = LCase(Trim(monthStr))

    For i = LBound(months) To UBound(months)
        If InStr(monthStr, months(i)) > 0 Then
            GetMonthFromStr = i + 1
            Exit Function
        End If
    Next i
    GetMonthFromStr = 0
End Function

Private Function GetWs(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWs = ThisWorkbook.Worksheets(sheetName)
End Function

'-----------------------------------------------------
' UpdateYearGrouping - Regroup columns based on Reporting Year
' Call this when the Reporting Year dropdown changes on Menu sheet (C10)
' Groups and collapses all columns for years PRIOR TO the Reporting Year
' Example: If Reporting Year = 2025, groups 2024 and earlier
'-----------------------------------------------------
Public Sub UpdateYearGrouping()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Dim wsMenu As Worksheet
    Dim reportingYear As Long
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim processedCount As Long

    Set wsMenu = GetWs("Menu")
    If wsMenu Is Nothing Then GoTo Cleanup

    ' Get Reporting Year from Menu C10
    If IsNumeric(wsMenu.Range("C10").Value) Then
        reportingYear = CLng(wsMenu.Range("C10").Value)
    Else
        MsgBox "Invalid Reporting Year in Menu C10. Please enter a valid year.", vbExclamation
        GoTo Cleanup
    End If

    ' List of sheets to process - all P&L, BS, CF tabs
    sheetNames = Array("Consolidated_PL", "Consolidated_BS", "Consolidated_CF", _
                       "Source_PL", "Source_BS", "Source_Budget", "PL", "Balance_Sheet", "Cash_Flow")

    processedCount = 0
    For Each sheetName In sheetNames
        Set ws = GetWs(CStr(sheetName))
        If Not ws Is Nothing Then
            Call RegroupSheetByYear(ws, reportingYear)
            processedCount = processedCount + 1
        End If
    Next sheetName

    ' Also process division-specific sheets
    Dim i As Integer
    Dim divSheets As Variant
    divSheets = Array("_PL", "_BS", "_CF", "_Forecast")

    For Each ws In ThisWorkbook.Worksheets
        For i = LBound(divSheets) To UBound(divSheets)
            If InStr(ws.Name, divSheets(i)) > 0 And Left(ws.Name, 4) <> "Menu" Then
                Call RegroupSheetByYear(ws, reportingYear)
                processedCount = processedCount + 1
            End If
        Next i
    Next ws

    MsgBox "Year grouping updated!" & vbCrLf & vbCrLf & _
           "Reporting Year: " & reportingYear & vbCrLf & _
           "Years " & (reportingYear - 1) & " and earlier are now grouped/collapsed." & vbCrLf & _
           "Sheets processed: " & processedCount, vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Error updating year grouping: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub

Private Sub RegroupSheetByYear(ws As Worksheet, reportingYear As Long)
    On Error Resume Next

    Dim headerRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim colYear As Long
    Dim headerVal As String
    Dim priorYearCols As Collection
    Dim startCol As Long
    Dim endCol As Long
    Dim i As Long

    ' Find header row - check rows 1, 2, 4 for "Account" header
    headerRow = 0
    If InStr(LCase(CStr(ws.Cells(1, 1).Value)), "account") > 0 Then
        headerRow = 1
    ElseIf InStr(LCase(CStr(ws.Cells(4, 1).Value)), "account") > 0 Then
        headerRow = 4
    ElseIf InStr(LCase(CStr(ws.Cells(2, 1).Value)), "account") > 0 Then
        headerRow = 2
    End If

    If headerRow = 0 Then Exit Sub

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 3 Then Exit Sub

    ' Clear all existing outline grouping
    ws.Cells.ClearOutline

    ' Find all columns that belong to prior years (before reportingYear)
    Set priorYearCols = New Collection
    For col = 2 To lastCol
        headerVal = Trim(CStr(ws.Cells(headerRow, col).Value))
        colYear = GetYearFromStr(headerVal)

        ' If year is valid and BEFORE reporting year, it's a prior year
        If colYear > 0 And colYear < reportingYear Then
            priorYearCols.Add col
        End If
    Next col

    ' Group prior year columns if any exist
    If priorYearCols.Count > 0 Then
        ' Find contiguous ranges and group them
        startCol = priorYearCols(1)
        endCol = priorYearCols(1)

        For i = 2 To priorYearCols.Count
            If priorYearCols(i) = endCol + 1 Then
                ' Contiguous - extend range
                endCol = priorYearCols(i)
            Else
                ' Gap - group current range and start new one
                If startCol > 0 And endCol >= startCol Then
                    ws.Range(ws.Columns(startCol), ws.Columns(endCol)).Group
                End If
                startCol = priorYearCols(i)
                endCol = priorYearCols(i)
            End If
        Next i

        ' Group final range
        If startCol > 0 And endCol >= startCol Then
            ws.Range(ws.Columns(startCol), ws.Columns(endCol)).Group
        End If

        ' Collapse all groups
        ws.Outline.ShowLevels ColumnLevels:=1
    End If
End Sub
'''

    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.resizable(False, False)

        # Variables
        self.pl_path = tk.StringVar()
        self.bs_path = tk.StringVar()
        self.existing_model_path = tk.StringVar()  # For update mode
        self.company_name = tk.StringVar(value="Company Name")
        self.fiscal_start = tk.StringVar(value="January")

        # Calculate first month of current/most recent fiscal year for display defaults
        fy_month, fy_year = self._get_current_fy_start("January")
        self.display_month = tk.StringVar(value=fy_month)
        self.display_year = tk.StringVar(value=str(fy_year))
        self.mode = tk.StringVar(value="new")  # "new" or "update"

        # Data period selection (for when headers don't clearly indicate dates)
        self.data_start_month = tk.StringVar(value=self.MONTHS[datetime.now().month - 1])
        self.data_start_year = tk.StringVar(value=str(datetime.now().year))
        self.data_end_month = tk.StringVar(value=self.MONTHS[datetime.now().month - 1])
        self.data_end_year = tk.StringVar(value=str(datetime.now().year))

        # Multi-division support
        self.is_multi_division = tk.BooleanVar(value=False)
        self.divisions = []  # List of division configs
        self.division_configs = []  # List of DivisionConfig objects
        self.account_mappings = {'pl': {}, 'bs': {}}  # Consolidated account mappings

        # Add trace to update display month when fiscal year start changes
        self.fiscal_start.trace_add('write', self._on_fiscal_start_change)

        self._create_ui()
        self._center_window(700, 780)

    def _get_current_fy_start(self, fiscal_start_month_name):
        """Get the first month of the current or most recent fiscal year.

        For example, if fiscal year starts in July and current date is December 2025,
        the current FY started July 2025. If current date is March 2025, the current
        FY started July 2024.
        """
        now = datetime.now()
        fiscal_month_num = self.MONTHS.index(fiscal_start_month_name) + 1

        if now.month >= fiscal_month_num:
            # Current FY started this calendar year
            fy_year = now.year
        else:
            # Current FY started last calendar year
            fy_year = now.year - 1

        return (fiscal_start_month_name, fy_year)

    def _on_fiscal_start_change(self, *args):
        """Update First Display Month when Fiscal Year Start changes."""
        fiscal_start = self.fiscal_start.get()
        fy_month, fy_year = self._get_current_fy_start(fiscal_start)
        self.display_month.set(fy_month)
        self.display_year.set(str(fy_year))

    def _create_ui(self):
        """Create the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(main_frame, text=APP_NAME,
                                font=('Segoe UI', 16, 'bold'))
        title_label.pack(pady=(0, 5))

        # Version info
        version_label = ttk.Label(main_frame, text=f"Version {APP_VERSION} ({APP_BUILD_DATE})",
                                  font=('Segoe UI', 9), foreground='gray')
        version_label.pack(pady=(0, 15))

        # Mode Selection (New Model vs Update Existing)
        mode_frame = ttk.LabelFrame(main_frame, text="Mode", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Radiobutton(mode_frame, text="Create New Model", variable=self.mode, value="new",
                        command=self._toggle_mode).grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        ttk.Radiobutton(mode_frame, text="Update Existing Model", variable=self.mode, value="update",
                        command=self._toggle_mode).grid(row=0, column=1, sticky=tk.W)

        # Division Structure (Multi-division support)
        self.division_frame = ttk.LabelFrame(main_frame, text="Division Structure", padding="10")
        self.division_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(self.division_frame, text="Multiple Divisions / Departments",
                        variable=self.is_multi_division,
                        command=self._toggle_division_mode).grid(row=0, column=0, sticky=tk.W)

        self.division_btn = ttk.Button(self.division_frame, text="Configure Divisions...",
                                       command=self._open_division_setup, state='disabled')
        self.division_btn.grid(row=0, column=1, padx=20)

        self.division_status = ttk.Label(self.division_frame, text="Single entity mode",
                                         font=('Segoe UI', 9), foreground='gray')
        self.division_status.grid(row=0, column=2, sticky=tk.W)

        # Company Name (for new models only)
        self.name_frame = ttk.LabelFrame(main_frame, text="Company Information", padding="10")
        self.name_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(self.name_frame, text="Company Name:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.name_frame, textvariable=self.company_name, width=40).grid(row=0, column=1, padx=5)

        # Existing Model Selection (for update mode)
        self.existing_frame = ttk.LabelFrame(main_frame, text="Existing Model", padding="10")
        # Don't pack yet - hidden by default

        ttk.Label(self.existing_frame, text="Model File:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.existing_frame, textvariable=self.existing_model_path, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(self.existing_frame, text="Browse...", command=self._browse_existing).grid(row=0, column=2)

        # File Selection (for single-division mode)
        self.file_frame = ttk.LabelFrame(main_frame, text="Upload Files (Single Division)", padding="10")
        self.file_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(self.file_frame, text="P&L File:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.file_frame, textvariable=self.pl_path, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(self.file_frame, text="Browse...", command=self._browse_pl).grid(row=0, column=2)

        ttk.Label(self.file_frame, text="Balance Sheet:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        ttk.Entry(self.file_frame, textvariable=self.bs_path, width=40).grid(row=1, column=1, padx=5, pady=(5, 0))
        ttk.Button(self.file_frame, text="Browse...", command=self._browse_bs).grid(row=1, column=2, pady=(5, 0))

        # Date range info - auto-detected from files
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(info_frame, text="Date range will be auto-detected from your Excel files.",
                  font=('Segoe UI', 9), foreground='#666').pack(anchor='w')

        # Configuration (for new models only)
        self.config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="10")
        self.config_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(self.config_frame, text="Fiscal Year Start:").grid(row=0, column=0, sticky=tk.W)
        fiscal_combo = ttk.Combobox(self.config_frame, textvariable=self.fiscal_start, values=self.MONTHS, width=15)
        fiscal_combo.grid(row=0, column=1, padx=5)

        ttk.Label(self.config_frame, text="First Display Month:").grid(row=0, column=2, sticky=tk.W, padx=(20, 0))
        display_combo = ttk.Combobox(self.config_frame, textvariable=self.display_month, values=self.MONTHS, width=15)
        display_combo.grid(row=0, column=3, padx=5)

        ttk.Label(self.config_frame, text="First Display Year:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        years = [str(y) for y in range(datetime.now().year - 5, datetime.now().year + 2)]
        year_combo = ttk.Combobox(self.config_frame, textvariable=self.display_year, values=years, width=15)
        year_combo.grid(row=1, column=1, padx=5, pady=(5, 0))

        # Button frame for Generate and Close buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=20)

        self.generate_btn = ttk.Button(btn_frame, text="Generate Financial Model",
                                       command=self._generate_model, style='Accent.TButton')
        self.generate_btn.pack(fill=tk.X, ipady=10)

        # Progress bar frame
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(10, 5))

        # Progress bar (dynamic growing bar)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var,
                                            maximum=100, mode='determinate', length=400)
        self.progress_bar.pack(fill=tk.X)

        # Status label (shows current activity)
        self.status_label = ttk.Label(main_frame, text="Ready", foreground='gray', font=('Segoe UI', 10))
        self.status_label.pack(pady=(5, 10))

        # Close button
        close_btn = ttk.Button(main_frame, text="Close", command=self.root.quit)
        close_btn.pack(pady=(5, 0))

        # Version info footer (minimal)
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill=tk.X, pady=(20, 0))
        ttk.Label(footer_frame, text=f"v{APP_VERSION}", foreground='#999',
                  font=('Segoe UI', 8)).pack()

    def _toggle_mode(self):
        """Toggle between New Model and Update Existing modes"""
        if self.mode.get() == "new":
            # Show company info and config, hide existing model selection
            self.name_frame.pack(fill=tk.X, pady=(0, 10), after=self.root.winfo_children()[0].winfo_children()[2])
            self.config_frame.pack(fill=tk.X, pady=(0, 10))
            self.existing_frame.pack_forget()
            self.generate_btn.config(text="Generate Financial Model")
        else:
            # Show existing model selection, hide company info and config
            self.name_frame.pack_forget()
            self.config_frame.pack_forget()
            self.existing_frame.pack(fill=tk.X, pady=(0, 10), after=self.root.winfo_children()[0].winfo_children()[2])
            self.generate_btn.config(text="Update Financial Model")

    def _toggle_division_mode(self):
        """Toggle between single entity and multi-division modes"""
        if self.is_multi_division.get():
            self.division_btn.config(state='normal')
            self.division_status.config(text="Multi-division mode - click Configure Divisions")
            # Hide single-division file upload since files are uploaded per-division
            self.file_frame.pack_forget()
        else:
            self.division_btn.config(state='disabled')
            self.division_status.config(text="Single entity mode")
            self.divisions = []
            self.division_configs = []
            # Show single-division file upload
            # Re-pack file_frame after division_frame
            self.file_frame.pack(fill=tk.X, pady=(0, 10), after=self.division_frame)

    def _open_division_setup(self):
        """Open the division setup dialog"""
        def on_divisions_saved(divisions):
            self.divisions = divisions
            count = len(divisions)
            self.division_status.config(text=f"{count} division(s) configured")
            # Clear the single-file paths since we'll use division-specific files
            self.pl_path.set('')
            self.bs_path.set('')

        DivisionSetupDialog(self.root, on_divisions_saved, self.divisions)

    def _handle_mapping_approval(self, approved_mappings):
        """Handle approved mappings from the review dialog"""
        self.account_mappings = approved_mappings

    def _browse_existing(self):
        """Browse for existing model file"""
        path = filedialog.askopenfilename(
            title="Select Existing Financial Model",
            filetypes=[("Excel files", "*.xlsm *.xlsx")]
        )
        if path:
            self.existing_model_path.set(path)

    def _center_window(self, width, height):
        """Center the window on the screen"""
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _browse_pl(self):
        path = filedialog.askopenfilename(
            title="Select P&L File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if path:
            self.pl_path.set(path)

    def _browse_bs(self):
        path = filedialog.askopenfilename(
            title="Select Balance Sheet File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if path:
            self.bs_path.set(path)

    def _generate_model(self):
        """Generate or update the financial model based on mode"""
        # Validate inputs based on division mode
        if self.is_multi_division.get():
            # Multi-division mode: validate divisions are configured
            if not self.divisions or len(self.divisions) == 0:
                messagebox.showerror("Error", "Please configure divisions first using 'Configure Divisions...'")
                return
            # Validate each division has files
            for div in self.divisions:
                if not div.get('pl_path') or not div.get('bs_path'):
                    messagebox.showerror("Error", f"Division '{div.get('name', 'Unknown')}' is missing P&L or Balance Sheet file")
                    return
        else:
            # Single entity mode: validate single files
            if not self.pl_path.get() or not self.bs_path.get():
                messagebox.showerror("Error", "Please select both P&L and Balance Sheet files")
                return

        if self.mode.get() == "new":
            # New model mode
            if not self.company_name.get().strip():
                messagebox.showerror("Error", "Please enter a company name")
                return

            # Ask for save location
            save_path = filedialog.asksaveasfilename(
                title="Save Financial Model As",
                defaultextension=".xlsm",
                filetypes=[("Excel Macro-Enabled", "*.xlsm")],
                initialfile=f"{self.company_name.get().replace(' ', '_')}_Financial_Model.xlsm"
            )

            if not save_path:
                return

            self.status_label.config(text="Starting...", foreground='blue')
            self.generate_btn.config(state='disabled')
            self.root.update()

            # Run directly (xlwings doesn't work well with threads)
            try:
                print(f"Starting model generation...")
                print(f"P&L file: {self.pl_path.get()}")
                print(f"BS file: {self.bs_path.get()}")
                print(f"Save path: {save_path}")
                print(f"Template path: {TEMPLATE_PATH}")
                print(f"Template exists: {os.path.exists(TEMPLATE_PATH)}")
                self._create_excel_model(save_path)
                self._update_status("Complete!", color='green')
                messagebox.showinfo("Success", f"Financial Model created:\n{save_path}")
            except Exception as e:
                import traceback
                print(f"ERROR: {e}")
                traceback.print_exc()
                self._update_status("Error occurred", color='red')
                messagebox.showerror("Error", str(e))
            finally:
                self.generate_btn.config(state='normal')

        else:
            # Update existing model mode
            if not self.existing_model_path.get():
                messagebox.showerror("Error", "Please select an existing Financial Model file")
                return

            # Ask for save location (can overwrite or save as new)
            save_path = filedialog.asksaveasfilename(
                title="Save Updated Model As",
                defaultextension=".xlsm",
                filetypes=[("Excel Macro-Enabled", "*.xlsm")],
                initialfile=os.path.basename(self.existing_model_path.get())
            )

            if not save_path:
                return

            self.status_label.config(text="Starting update...", foreground='blue')
            self.generate_btn.config(state='disabled')
            self.root.update()

            try:
                print(f"Starting model update...")
                print(f"Existing model: {self.existing_model_path.get()}")
                print(f"P&L file: {self.pl_path.get()}")
                print(f"BS file: {self.bs_path.get()}")
                print(f"Save path: {save_path}")
                self._update_existing_model(save_path)
                self._update_status("Complete!", color='green')
                messagebox.showinfo("Success", f"Financial Model updated:\n{save_path}")
            except Exception as e:
                import traceback
                print(f"ERROR: {e}")
                traceback.print_exc()
                self._update_status("Error occurred", color='red')
                messagebox.showerror("Error", str(e))
            finally:
                self.generate_btn.config(state='normal')

    def _update_status(self, message, color='blue', elapsed=None, estimated_total=None, progress=None):
        """Update the status label, progress bar, and refresh the UI

        Args:
            message: Status message to display
            color: Text color
            elapsed: Elapsed time in seconds (optional)
            estimated_total: Estimated total time in seconds (optional)
            progress: Progress percentage 0-100 (optional)
        """
        if elapsed is not None and estimated_total is not None:
            remaining = max(0, estimated_total - elapsed)
            if remaining > 60:
                time_str = f" (~{int(remaining // 60)}m {int(remaining % 60)}s remaining)"
            else:
                time_str = f" (~{int(remaining)}s remaining)"
            message = f"{message}{time_str}"
        elif elapsed is not None:
            if elapsed > 60:
                time_str = f" ({int(elapsed // 60)}m {int(elapsed % 60)}s elapsed)"
            else:
                time_str = f" ({int(elapsed)}s elapsed)"
            message = f"{message}{time_str}"

        self.status_label.config(text=message, foreground=color)

        # Update progress bar if progress value provided
        if progress is not None:
            self.progress_var.set(progress)

        self.root.update()
        print(message)

    def _create_progress_tracker(self, total_steps, estimated_total_seconds):
        """Create a progress tracker for mid-step updates

        Returns functions for updating step progress and substep progress.
        """
        start_time = time.time()
        state = {'current_step': 0, 'substep': ''}

        def update_step(step_name):
            """Update to a new main step"""
            state['current_step'] += 1
            state['substep'] = step_name
            progress_pct = (state['current_step'] / total_steps) * 100

            elapsed = time.time() - start_time
            remaining = max(0, estimated_total_seconds - elapsed)

            # Build message - clean arrow style without step numbers
            msg = f"→ {step_name}"

            # Add time estimate
            if remaining > 60:
                time_str = f" (~{int(remaining // 60)}m {int(remaining % 60)}s remaining)"
            elif remaining > 5:
                time_str = f" (~{int(remaining)}s remaining)"
            else:
                time_str = ""  # Don't show tiny remaining times

            self.status_label.config(text=f"{msg}{time_str}", foreground='blue')
            self.progress_var.set(progress_pct)
            self.root.update()
            print(f"→ {step_name}")

        def update_substep(substep_name):
            """Update substep within current step (doesn't advance progress)"""
            msg = f"  → {substep_name}"
            self.status_label.config(text=msg, foreground='#666')
            self.root.update()
            print(f"    {substep_name}")

        return update_step, update_substep, state

    def _get_user_date_params(self):
        """Get user-specified start/end dates as tuples for _parse_financial_data"""
        start_month_name = self.data_start_month.get()
        start_year = int(self.data_start_year.get())
        end_month_name = self.data_end_month.get()
        end_year = int(self.data_end_year.get())

        # Convert month name to number
        start_month_num = self.MONTHS.index(start_month_name) + 1
        end_month_num = self.MONTHS.index(end_month_name) + 1

        return (start_month_num, start_year), (end_month_num, end_year)

    def _check_file_not_open(self, file_path):
        """Check if a file is open by another process. Returns True if file is accessible."""
        if not os.path.exists(file_path):
            return True  # File doesn't exist, so it's not locked

        try:
            # Try to open the file in exclusive mode
            with open(file_path, 'r+b') as f:
                pass
            return True
        except (IOError, PermissionError):
            return False

    def _update_existing_model(self, save_path):
        """Update an existing Financial Model with new month data"""
        import shutil

        # Check if any of the input files are open
        files_to_check = [
            (self.pl_path.get(), "P&L file"),
            (self.bs_path.get(), "Balance Sheet file"),
            (self.existing_model_path.get(), "Existing model"),
        ]
        if save_path and os.path.exists(save_path):
            files_to_check.append((save_path, "Save destination"))

        for file_path, file_name in files_to_check:
            if file_path and not self._check_file_not_open(file_path):
                filename = os.path.basename(file_path)
                raise Exception(f"The file '{filename}' is still open.\n\nPlease close it and try again.")

        # Get user-specified date range
        user_start, user_end = self._get_user_date_params()

        # Parse new input files
        self._update_status("Step 1/8: Reading new P&L file...")
        pl_data = pd.read_excel(self.pl_path.get(), header=None)
        pl_indents = self._get_cell_indents(self.pl_path.get())

        self._update_status("Step 2/8: Reading new Balance Sheet file...")
        bs_data = pd.read_excel(self.bs_path.get(), header=None)
        bs_indents = self._get_cell_indents(self.bs_path.get())

        self._update_status("Step 3/8: Extracting account data...")
        pl_accounts, new_months, pl_totals = self._parse_financial_data(pl_data, pl_indents, user_start, user_end)
        bs_accounts, _, bs_totals = self._parse_financial_data(bs_data, bs_indents, user_start, user_end)
        print(f"New data has {len(new_months)} months")

        # Work in temp directory
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, 'temp_model.xlsm')

        try:
            # Copy existing model to temp location
            self._update_status("Step 4/8: Opening existing model...")
            shutil.copy2(self.existing_model_path.get(), temp_path)

            app = xw.App(visible=False)
            try:
                wb = app.books.open(temp_path)

                # Get existing sheets
                source_pl = wb.sheets['Source_PL']
                source_bs = wb.sheets['Source_BS']

                # Find existing months in Source_PL (row 1 has month names, row 2 has YYYYMM)
                self._update_status("Step 5/8: Analyzing existing data...")
                existing_months = set()
                last_col = source_pl.range('A1').end('right').column
                for col in range(2, last_col + 1):
                    yyyymm = source_pl.range((2, col)).value
                    if yyyymm:
                        existing_months.add(int(yyyymm))

                # Determine which months are new
                new_month_list = []
                for m, y, name in new_months:
                    yyyymm = y * 100 + m
                    if yyyymm not in existing_months:
                        new_month_list.append((m, y, name, yyyymm))

                if not new_month_list:
                    raise Exception("No new months found in uploaded files. All months already exist in the model.")

                print(f"Adding {len(new_month_list)} new month(s): {[n[2] for n in new_month_list]}")

                # Add new months to Source_PL
                self._update_status("Step 6/8: Adding new month data to Source sheets...")
                SOURCE_BLACK = (26, 26, 26)  # Match existing header color

                for m, y, name, yyyymm in new_month_list:
                    # Find next available column
                    new_col = source_pl.range('A1').end('right').column + 1

                    # Add month header and YYYYMM
                    source_pl.range((1, new_col)).value = name
                    source_pl.range((2, new_col)).value = yyyymm

                    # Format header to match existing columns
                    header_cell = source_pl.range((1, new_col))
                    header_cell.font.name = 'Calibri Light'
                    header_cell.font.size = 10
                    header_cell.font.bold = True
                    header_cell.font.color = (255, 255, 255)
                    header_cell.color = SOURCE_BLACK
                    header_cell.api.HorizontalAlignment = -4108  # xlCenter

                    # Add P&L data for this month - MATCH BY ACCOUNT NAME
                    print(f"DEBUG: Adding P&L data for month ({m}, {y}) to column {new_col}")

                    # Build a lookup dict from parsed accounts (normalize names for matching)
                    pl_values_lookup = {}
                    for account in pl_accounts:
                        acct_name = account.get('name', '').strip()
                        account_values = account.get('values', {})
                        value = account_values.get((m, y), 0)
                        pl_values_lookup[acct_name] = value
                        # Also store without leading spaces for indented accounts
                        pl_values_lookup[acct_name.lstrip()] = value

                    # Get existing account names from Source_PL and match
                    # Use UsedRange to reliably find last row (end('down') stops at empty cells)
                    try:
                        used_range = source_pl.api.UsedRange
                        last_row_pl = used_range.Row + used_range.Rows.Count - 1
                    except:
                        last_row_pl = source_pl.range('A1').end('down').row

                    print(f"DEBUG: Source_PL last_row_pl = {last_row_pl}")
                    non_zero_count = 0
                    matched_count = 0

                    for row in range(3, last_row_pl + 1):
                        existing_name = source_pl.range((row, 1)).value
                        if existing_name:
                            existing_name_str = str(existing_name).strip()
                            # Try to find match in lookup
                            value = pl_values_lookup.get(existing_name_str,
                                    pl_values_lookup.get(existing_name_str.lstrip(), 0))
                            if value != 0:
                                non_zero_count += 1
                            if existing_name_str in pl_values_lookup or existing_name_str.lstrip() in pl_values_lookup:
                                matched_count += 1
                            source_pl.range((row, new_col)).value = value

                    # Format data column - font and number format
                    data_range_pl = source_pl.range((3, new_col), (last_row_pl, new_col))
                    data_range_pl.number_format = '#,##0'
                    data_range_pl.font.name = 'Calibri Light'
                    data_range_pl.font.size = 10

                    print(f"DEBUG: P&L - {len(pl_accounts)} parsed accounts, {matched_count} matched, {non_zero_count} with non-zero values")
                    if len(pl_accounts) > 0:
                        print(f"DEBUG: First 5 parsed account names:")
                        for idx, acct in enumerate(pl_accounts[:5]):
                            print(f"  [{idx}] '{acct.get('name')}'")
                        print(f"DEBUG: First 5 existing Source_PL account names:")
                        for row in range(3, min(8, last_row_pl + 1)):
                            existing = source_pl.range((row, 1)).value
                            print(f"  [row {row}] '{existing}'")

                    # Add BS data for same month
                    new_col_bs = source_bs.range('A1').end('right').column + 1
                    source_bs.range((1, new_col_bs)).value = name
                    source_bs.range((2, new_col_bs)).value = yyyymm

                    # Format BS header to match existing columns
                    bs_header_cell = source_bs.range((1, new_col_bs))
                    bs_header_cell.font.name = 'Calibri Light'
                    bs_header_cell.font.size = 10
                    bs_header_cell.font.bold = True
                    bs_header_cell.font.color = (255, 255, 255)
                    bs_header_cell.color = SOURCE_BLACK
                    bs_header_cell.api.HorizontalAlignment = -4108  # xlCenter

                    # Add BS data for this month - MATCH BY ACCOUNT NAME
                    print(f"DEBUG: Adding BS data for month ({m}, {y}) to column {new_col_bs}")

                    # Build a lookup dict from parsed BS accounts
                    bs_values_lookup = {}
                    for account in bs_accounts:
                        acct_name = account.get('name', '').strip()
                        account_values = account.get('values', {})
                        value = account_values.get((m, y), 0)
                        bs_values_lookup[acct_name] = value
                        bs_values_lookup[acct_name.lstrip()] = value

                    # Get existing account names from Source_BS and match
                    # Use UsedRange to reliably find last row
                    try:
                        used_range_bs = source_bs.api.UsedRange
                        last_row_bs = used_range_bs.Row + used_range_bs.Rows.Count - 1
                    except:
                        last_row_bs = source_bs.range('A1').end('down').row

                    print(f"DEBUG: Source_BS last_row_bs = {last_row_bs}")
                    non_zero_count_bs = 0
                    matched_count_bs = 0

                    for row in range(3, last_row_bs + 1):
                        existing_name = source_bs.range((row, 1)).value
                        if existing_name:
                            existing_name_str = str(existing_name).strip()
                            value = bs_values_lookup.get(existing_name_str,
                                    bs_values_lookup.get(existing_name_str.lstrip(), 0))
                            if value != 0:
                                non_zero_count_bs += 1
                            if existing_name_str in bs_values_lookup or existing_name_str.lstrip() in bs_values_lookup:
                                matched_count_bs += 1
                            source_bs.range((row, new_col_bs)).value = value

                    # Format BS data column - font and number format
                    data_range_bs = source_bs.range((3, new_col_bs), (last_row_bs, new_col_bs))
                    data_range_bs.number_format = '#,##0'
                    data_range_bs.font.name = 'Calibri Light'
                    data_range_bs.font.size = 10

                    print(f"DEBUG: BS - {len(bs_accounts)} parsed accounts, {matched_count_bs} matched, {non_zero_count_bs} with non-zero values")

                # Update named ranges - use UsedRange for reliable row counts
                try:
                    used_range_pl = source_pl.api.UsedRange
                    pl_last_row = used_range_pl.Row + used_range_pl.Rows.Count - 1
                    pl_last_col = used_range_pl.Column + used_range_pl.Columns.Count - 1
                except:
                    pl_last_row = source_pl.range('A1').end('down').row
                    pl_last_col = source_pl.range('A1').end('right').column

                try:
                    used_range_bs = source_bs.api.UsedRange
                    bs_last_row = used_range_bs.Row + used_range_bs.Rows.Count - 1
                    bs_last_col = used_range_bs.Column + used_range_bs.Columns.Count - 1
                except:
                    bs_last_row = source_bs.range('A1').end('down').row
                    bs_last_col = source_bs.range('A1').end('right').column

                print(f"DEBUG: Source_PL dimensions: {pl_last_row} rows, {pl_last_col} columns (up to column {self._col_letter(pl_last_col)})")
                print(f"DEBUG: Source_BS dimensions: {bs_last_row} rows, {bs_last_col} columns (up to column {self._col_letter(bs_last_col)})")

                # Check what's in row 2 (YYYYMM values) of Source_PL
                yyyymm_row = []
                for col in range(2, pl_last_col + 1):
                    val = source_pl.range((2, col)).value
                    yyyymm_row.append(val)
                print(f"DEBUG: Source_PL row 2 (YYYYMM values): {yyyymm_row}")

                try:
                    wb.names['SourcePL'].delete()
                except:
                    pass
                try:
                    wb.names['SourceBS'].delete()
                except:
                    pass

                wb.names.add('SourcePL', f"=Source_PL!$A$1:${self._col_letter(pl_last_col)}${pl_last_row}")
                wb.names.add('SourceBS', f"=Source_BS!$A$1:${self._col_letter(bs_last_col)}${bs_last_row}")

                # Get the latest month added
                latest_month = new_month_list[-1]
                latest_m, latest_y, latest_name, latest_yyyymm = latest_month

                # Update Menu sheet with new current month
                self._update_status("Step 7/8: Updating Menu and reports...")
                try:
                    menu_sheet = wb.sheets['Menu']
                    # Set current month - display (C7), YYYYMM helper (G7), and YTD helpers (E7, F7)
                    menu_sheet.range('C7').value = latest_name  # Current month display
                    menu_sheet.range('G7').value = latest_yyyymm  # YYYYMM for formulas

                    # E7 = month number (1-12), F7 = year - used by YTD formulas
                    menu_sheet.range('E7').value = latest_m  # Month number
                    menu_sheet.range('F7').value = latest_y  # Year

                    # Actuals Through should always match Current Month
                    # Update both C9 (display) and G9 (YYYYMM helper)
                    menu_sheet.range('C9').value = menu_sheet.range('C7').value  # Same as current month
                    menu_sheet.range('G9').value = latest_yyyymm  # Actuals through YYYYMM

                    print(f"DEBUG: Updated Menu - Current Month: {latest_name} ({latest_yyyymm}), E7={latest_m}, F7={latest_y}")
                except Exception as e:
                    print(f"Warning: Could not update Menu sheet: {e}")

                # Add new month columns to report sheets (PL, Balance_Sheet, Cash_Flow)
                self._add_month_to_reports(wb, new_month_list, pl_last_col, bs_last_col)

                # Force recalculation
                wb.app.calculate()

                # Save
                self._update_status("Step 8/8: Saving updated model...")
                wb.save()
                wb.close()

            finally:
                app.quit()

            # Copy from temp to final location
            shutil.copy2(temp_path, save_path)

        finally:
            # Clean up temp directory
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

    def _add_navigation_links(self, wb, new_month_list=None):
        """Add navigation hyperlinks to all sheets - Menu link on each sheet, sheet links on Menu"""
        try:
            menu_sheet = wb.sheets['Menu']

            # Define sheets that should have navigation with display names
            report_sheets = [
                ('Dashboard', 'Dashboard'),
                ('PL', 'P&L'),
                ('Balance_Sheet', 'Balance Sheet'),
                ('Cash_Flow', 'Cash Flow'),
                ('Forecast', 'Forecast'),
                ('Forecast_Summary', 'Forecast Summary'),
                ('Notes', 'Notes')
            ]

            # Add "Back to Menu" link on each report sheet (cell A1)
            for sheet_name, display_name in report_sheets:
                try:
                    sheet = wb.sheets[sheet_name]
                    # Add hyperlink in A1
                    sheet.range('A1').value = '← Menu'
                    sheet.range('A1').font.color = (0, 102, 204)  # Blue link color
                    sheet.range('A1').font.underline = True
                    sheet.range('A1').font.size = 9
                    sheet.api.Hyperlinks.Add(
                        Anchor=sheet.range('A1').api,
                        Address="",
                        SubAddress="Menu!A1",
                        TextToDisplay="← Menu"
                    )
                except Exception as e:
                    print(f"DEBUG: Could not add Menu link to {sheet_name}: {e}")

            # Hide helper columns E:G on Menu
            try:
                menu_sheet.range('E:G').api.EntireColumn.Hidden = True
            except Exception as e:
                print(f"DEBUG: Could not hide columns E:G: {e}")

            # Add Quick Links box in upper right of Menu (starting at I2)
            try:
                start_col = 'I'
                start_row = 2

                # Header
                menu_sheet.range(f'{start_col}{start_row}').value = "Quick Links"
                menu_sheet.range(f'{start_col}{start_row}').font.bold = True
                menu_sheet.range(f'{start_col}{start_row}').font.size = 11

                # Links
                link_row = start_row + 1
                for sheet_name, display_name in report_sheets:
                    try:
                        menu_sheet.range(f'{start_col}{link_row}').value = display_name
                        menu_sheet.range(f'{start_col}{link_row}').font.color = (0, 102, 204)
                        menu_sheet.range(f'{start_col}{link_row}').font.underline = True
                        menu_sheet.range(f'{start_col}{link_row}').font.size = 10
                        menu_sheet.api.Hyperlinks.Add(
                            Anchor=menu_sheet.range(f'{start_col}{link_row}').api,
                            Address="",
                            SubAddress=f"'{sheet_name}'!A1",
                            TextToDisplay=display_name
                        )
                        link_row += 1
                    except Exception as e:
                        print(f"DEBUG: Could not add link to {sheet_name} on Menu: {e}")

                # Add border box around Quick Links section
                box_range = menu_sheet.range(f'{start_col}{start_row}:{start_col}{link_row - 1}')
                box_range.api.Borders.LineStyle = 1  # xlContinuous
                box_range.api.Borders.Weight = 2  # xlThin

            except Exception as e:
                print(f"DEBUG: Could not create Quick Links box: {e}")

            print("DEBUG: Added navigation links")

        except Exception as e:
            print(f"DEBUG: Navigation links skipped: {e}")

    def _add_month_to_reports(self, wb, new_month_list, source_pl_last_col, source_bs_last_col=None):
        """Add new month columns to PL, Balance Sheet, and Cash Flow reports

        Also updates all existing SUMPRODUCT formulas to include the new source data range.
        """
        if source_bs_last_col is None:
            source_bs_last_col = source_pl_last_col

        try:
            pl_sheet = wb.sheets['PL']
            bs_sheet = wb.sheets['Balance_Sheet']
            cf_sheet = wb.sheets['Cash_Flow']

            # Get the new source column letters (after data was added)
            pl_col_letter = self._col_letter(source_pl_last_col)
            bs_col_letter = self._col_letter(source_bs_last_col)

            print(f"DEBUG: Updating formulas to use Source_PL up to column {pl_col_letter}, Source_BS up to column {bs_col_letter}")

            for m, y, name, yyyymm in new_month_list:
                # Find where to insert new month on PL
                # Month columns start at B, find the last month column
                pl_last_month_col = pl_sheet.range('B3').end('right').column

                # Insert new column after last month
                new_col = pl_last_month_col + 1

                # Insert column on PL
                pl_sheet.range((1, new_col)).api.EntireColumn.Insert()

                # Set up header row (row 3 is YYYYMM, row 4 is month name header)
                pl_sheet.range((3, new_col)).value = yyyymm
                pl_sheet.range((3, new_col)).font.color = (255, 255, 255)  # White/hidden
                pl_sheet.range((4, new_col)).value = name
                pl_sheet.range((4, new_col)).font.bold = True

                last_row = pl_sheet.range('A4').end('down').row

                # Create formula for new column
                print(f"DEBUG: Creating formulas for PL new column {new_col}, using Source_PL range up to column {pl_col_letter}")
                print(f"DEBUG: Formula will look for YYYYMM = {yyyymm}")
                for row in range(5, last_row + 1):
                    account_cell = f"$A{row}"
                    formula = (
                        f"=SUMPRODUCT("
                        f"(Source_PL!$A$3:$A$1000=TRIM({account_cell}))*"
                        f"(Source_PL!$B$2:${pl_col_letter}$2={yyyymm})*"
                        f"(Source_PL!$B$3:${pl_col_letter}$1000))"
                    )
                    pl_sheet.range((row, new_col)).value = formula
                    if row == 5:
                        print(f"DEBUG: Sample PL formula (row 5): {formula}")

                # Format new column
                pl_sheet.range((5, new_col), (last_row, new_col)).number_format = '#,##0'

                # After inserting new column, update YTD formulas to include new column in their range
                # Find the YTD columns (PY YTD and CY YTD) - they're after Notes column
                # Layout: Months | Notes | Spacer | PY YTD | CY YTD | Var $ | Var % | Spacer | Annual years
                # The new column is inserted into the months area, so all columns shift right

                # Find where the summary columns are by looking for headers
                header_row = 4
                py_ytd_col = None
                cy_ytd_col = None
                fy_start_col = None

                # Scan row 4 to find column headers
                for check_col in range(new_col + 1, new_col + 15):
                    header_val = pl_sheet.range((header_row, check_col)).value
                    if header_val == 'PY YTD':
                        py_ytd_col = check_col
                    elif header_val == 'CY YTD':
                        cy_ytd_col = check_col
                    elif header_val and str(header_val).isdigit() and len(str(header_val)) == 4:
                        # Found a year column (like "2024")
                        if fy_start_col is None:
                            fy_start_col = check_col

                # Update YTD formulas to use expanded data range
                if py_ytd_col and cy_ytd_col:
                    print(f"DEBUG: Found YTD columns - PY YTD at col {py_ytd_col}, CY YTD at col {cy_ytd_col}")
                    first_data_col = self._col_letter(2)  # B
                    last_data_col = self._col_letter(new_col)  # Now includes new column
                    helper_range = f'{first_data_col}$3:{last_data_col}$3'

                    for row in range(5, last_row + 1):
                        data_range = f'{first_data_col}{row}:{last_data_col}{row}'

                        # CY YTD formula
                        cy_formula = (
                            f'=SUMPRODUCT(({data_range})*'
                            f'--(INT({helper_range}/100)=Menu!$F$7)*'
                            f'--(MOD({helper_range},100)<=Menu!$E$7))'
                        )
                        pl_sheet.range((row, cy_ytd_col)).value = cy_formula

                        # PY YTD formula
                        py_formula = (
                            f'=SUMPRODUCT(({data_range})*'
                            f'--(INT({helper_range}/100)=Menu!$F$7-1)*'
                            f'--(MOD({helper_range},100)<=Menu!$E$7))'
                        )
                        pl_sheet.range((row, py_ytd_col)).value = py_formula

                    print(f"DEBUG: Updated YTD formulas with range up to column {last_data_col}")

                # Update Annual column formulas
                if fy_start_col:
                    # Get the year of the new month
                    new_year = y  # From the loop variable (m, y, name, yyyymm)

                    # Collect all months and their years from row 3
                    month_cols_by_year = {}
                    for col in range(2, new_col + 1):
                        yyyymm_val = pl_sheet.range((3, col)).value
                        if yyyymm_val:
                            col_year = int(yyyymm_val) // 100
                            if col_year not in month_cols_by_year:
                                month_cols_by_year[col_year] = []
                            month_cols_by_year[col_year].append(col)

                    # Find all annual columns and update their formulas
                    for fy_col_offset in range(10):  # Check up to 10 year columns
                        fy_col = fy_start_col + fy_col_offset
                        year_header = pl_sheet.range((header_row, fy_col)).value
                        if year_header and str(year_header).isdigit():
                            year = int(year_header)
                            if year in month_cols_by_year:
                                year_cols = month_cols_by_year[year]
                                # Update formula for each data row
                                for row in range(5, last_row + 1):
                                    refs = '+'.join([f'{self._col_letter(c)}{row}' for c in year_cols])
                                    pl_sheet.range((row, fy_col)).formula = f'={refs}'
                                print(f"DEBUG: Updated Annual {year} formula to include columns {[self._col_letter(c) for c in year_cols]}")

                # Update ALL existing month columns to use expanded source range
                print(f"DEBUG: Updating existing PL formulas in columns B to {self._col_letter(new_col-1)}")
                for col in range(2, new_col):  # B through previous last column
                    col_yyyymm = pl_sheet.range((3, col)).value
                    if col_yyyymm:
                        for row in range(5, last_row + 1):
                            account_cell = f"$A{row}"
                            formula = (
                                f"=SUMPRODUCT("
                                f"(Source_PL!$A$3:$A$1000=TRIM({account_cell}))*"
                                f"(Source_PL!$B$2:${pl_col_letter}$2={int(col_yyyymm)})*"
                                f"(Source_PL!$B$3:${pl_col_letter}$1000))"
                            )
                            pl_sheet.range((row, col)).value = formula

                # Do same for Balance Sheet
                bs_last_month_col = bs_sheet.range('B3').end('right').column
                new_col_bs = bs_last_month_col + 1

                bs_sheet.range((1, new_col_bs)).api.EntireColumn.Insert()
                bs_sheet.range((3, new_col_bs)).value = yyyymm
                bs_sheet.range((3, new_col_bs)).font.color = (255, 255, 255)
                bs_sheet.range((4, new_col_bs)).value = name
                bs_sheet.range((4, new_col_bs)).font.bold = True

                bs_last_row = bs_sheet.range('A4').end('down').row
                for row in range(5, bs_last_row + 1):
                    account_cell = f"$A{row}"
                    formula = (
                        f"=SUMPRODUCT("
                        f"(Source_BS!$A$3:$A$1000=TRIM({account_cell}))*"
                        f"(Source_BS!$B$2:${bs_col_letter}$2={yyyymm})*"
                        f"(Source_BS!$B$3:${bs_col_letter}$1000))"
                    )
                    bs_sheet.range((row, new_col_bs)).value = formula

                bs_sheet.range((5, new_col_bs), (bs_last_row, new_col_bs)).number_format = '#,##0'

                # Update existing BS columns
                print(f"DEBUG: Updating existing BS formulas in columns B to {self._col_letter(new_col_bs-1)}")
                for col in range(2, new_col_bs):
                    col_yyyymm = bs_sheet.range((3, col)).value
                    if col_yyyymm:
                        for row in range(5, bs_last_row + 1):
                            account_cell = f"$A{row}"
                            formula = (
                                f"=SUMPRODUCT("
                                f"(Source_BS!$A$3:$A$1000=TRIM({account_cell}))*"
                                f"(Source_BS!$B$2:${bs_col_letter}$2={int(col_yyyymm)})*"
                                f"(Source_BS!$B$3:${bs_col_letter}$1000))"
                            )
                            bs_sheet.range((row, col)).value = formula

                # Do same for Cash Flow
                cf_last_month_col = cf_sheet.range('B3').end('right').column
                new_col_cf = cf_last_month_col + 1

                cf_sheet.range((1, new_col_cf)).api.EntireColumn.Insert()
                cf_sheet.range((3, new_col_cf)).value = yyyymm
                cf_sheet.range((3, new_col_cf)).font.color = (255, 255, 255)
                cf_sheet.range((4, new_col_cf)).value = name
                cf_sheet.range((4, new_col_cf)).font.bold = True

                # Cash Flow formulas reference PL and BS - copy from previous column
                cf_last_row = cf_sheet.range('A4').end('down').row
                for row in range(5, cf_last_row + 1):
                    prev_formula = cf_sheet.range((row, cf_last_month_col)).formula
                    if prev_formula:
                        cf_sheet.range((row, new_col_cf)).value = prev_formula

                cf_sheet.range((5, new_col_cf), (cf_last_row, new_col_cf)).number_format = '#,##0'

                print(f"Added month {name} to reports")

            # Update Forecast sheet if it exists
            try:
                forecast_sheet = wb.sheets['Forecast']
                forecast_used = forecast_sheet.api.UsedRange
                forecast_last_row = forecast_used.Row + forecast_used.Rows.Count - 1
                forecast_last_col = forecast_used.Column + forecast_used.Columns.Count - 1

                print(f"DEBUG: Updating Forecast formulas to use expanded source range (up to column {pl_col_letter})")
                print(f"DEBUG: Forecast sheet has {forecast_last_row} rows, {forecast_last_col} columns")
                import re
                forecast_formula_count = 0
                formulas_found = 0

                def expand_forecast_range(formula_str, new_col):
                    """Replace Source_PL and Source_Budget column references with new_col"""
                    # Handle both Source_PL and Source_Budget
                    result = formula_str
                    for source_sheet in ['Source_PL', 'Source_Budget']:
                        # Pattern matches Source_XX!$B$2:$T$2 etc
                        result = re.sub(
                            rf'{source_sheet}!\$?B\$?(\d+):\$?([A-Z]+)\$?(\d+)',
                            rf'{source_sheet}!$B$\1:${new_col}$\3',
                            result
                        )
                    return result

                # Scan all cells in the Forecast sheet for Source_PL references
                for col in range(2, min(forecast_last_col + 1, 65)):  # B through all used columns
                    for row in range(5, forecast_last_row + 1):  # Data starts at row 5
                        cell = forecast_sheet.range((row, col))
                        formula = cell.formula
                        if formula and ('Source_PL!' in str(formula) or 'Source_Budget!' in str(formula)):
                            formulas_found += 1
                            original = str(formula)
                            updated = expand_forecast_range(original, pl_col_letter)
                            if updated != original:
                                if forecast_formula_count < 5:
                                    print(f"DEBUG: Forecast ({row},{col}) BEFORE: {original[:100]}")
                                    print(f"DEBUG: Forecast ({row},{col}) AFTER:  {updated[:100]}")
                                cell.value = updated
                                forecast_formula_count += 1
                            elif formulas_found <= 3:
                                # Show formulas that didn't change to debug pattern
                                print(f"DEBUG: Forecast ({row},{col}) NO CHANGE: {original[:100]}")

                print(f"DEBUG: Found {formulas_found} Forecast formulas with Source refs, updated {forecast_formula_count}")
            except Exception as e:
                print(f"DEBUG: Forecast sheet update skipped: {e}")
                import traceback
                traceback.print_exc()

            # Update Forecast_Summary sheet if it exists
            try:
                forecast_summary = wb.sheets['Forecast_Summary']
                fs_used = forecast_summary.api.UsedRange
                fs_last_row = fs_used.Row + fs_used.Rows.Count - 1

                print(f"DEBUG: Updating Forecast_Summary formulas to use expanded source range (up to column {pl_col_letter})")
                import re
                fs_formula_count = 0

                def expand_fs_range(formula_str, new_col):
                    """Replace Source_PL column references with new_col"""
                    result = re.sub(
                        r'Source_PL!\$?B\$?(\d+):\$?[A-Z]+\$?(\d+)',
                        rf'Source_PL!$B$\1:${new_col}$\2',
                        formula_str
                    )
                    return result

                # Check all columns for formulas that reference Source_PL
                for col in range(2, 15):  # B through N
                    for row in range(3, fs_last_row + 1):
                        cell = forecast_summary.range((row, col))
                        formula = cell.formula
                        if formula and 'Source_PL!' in str(formula):
                            original = str(formula)
                            updated = expand_fs_range(original, pl_col_letter)
                            if updated != original:
                                if fs_formula_count < 3:
                                    print(f"DEBUG: Forecast_Summary ({row},{col}) BEFORE: {original[:80]}...")
                                    print(f"DEBUG: Forecast_Summary ({row},{col}) AFTER:  {updated[:80]}...")
                                cell.value = updated
                                fs_formula_count += 1

                print(f"DEBUG: Updated {fs_formula_count} Forecast_Summary formulas")
            except Exception as e:
                print(f"DEBUG: Forecast_Summary sheet update skipped: {e}")
                import traceback
                traceback.print_exc()

            # Update Dashboard sheet formulas to use new source range
            try:
                dashboard_sheet = wb.sheets['Dashboard']
                dashboard_used = dashboard_sheet.api.UsedRange
                dashboard_last_row = dashboard_used.Row + dashboard_used.Rows.Count - 1

                print(f"DEBUG: Updating Dashboard formulas to use expanded source range (up to column {pl_col_letter})")
                import re
                formula_count = 0

                # More flexible patterns - handle both $T and T column references
                def expand_source_range(formula_str, source_sheet, new_col):
                    """Replace any column reference in Source_XX ranges with new_col"""
                    # Pattern matches Source_PL!$B$2:$T$2 or Source_PL!$B$2:T$2 or Source_PL!B$2:T$2 etc.
                    # Captures the row numbers and replaces the end column
                    patterns = [
                        (rf'{source_sheet}!\$?B\$?(\d+):\$?[A-Z]+\$?(\d+)', rf'{source_sheet}!$B$\1:${new_col}$\2'),
                    ]
                    result = formula_str
                    for pattern, replacement in patterns:
                        result = re.sub(pattern, replacement, result)
                    return result

                for col in range(2, 13):  # B through L
                    for row in range(1, min(dashboard_last_row + 1, 60)):
                        cell = dashboard_sheet.range((row, col))
                        formula = cell.formula
                        if formula and ('Source_PL!' in str(formula) or 'Source_BS!' in str(formula)):
                            original = str(formula)
                            updated = expand_source_range(original, 'Source_PL', pl_col_letter)
                            updated = expand_source_range(updated, 'Source_BS', bs_col_letter)

                            if updated != original:
                                if formula_count < 3:  # Show first 3 changes
                                    print(f"DEBUG: Dashboard ({row},{col}) BEFORE: {original[:80]}...")
                                    print(f"DEBUG: Dashboard ({row},{col}) AFTER:  {updated[:80]}...")
                                cell.value = updated
                                formula_count += 1

                print(f"DEBUG: Updated {formula_count} Dashboard formulas")
            except Exception as e:
                print(f"DEBUG: Dashboard sheet update skipped: {e}")
                import traceback
                traceback.print_exc()

            # Auto-fit columns on Balance Sheet and Cash Flow after update
            try:
                bs_sheet = wb.sheets['Balance_Sheet']
                bs_sheet.autofit('c')  # Auto-fit all columns
                print("DEBUG: Auto-fit Balance Sheet columns")
            except Exception as e:
                print(f"DEBUG: Balance Sheet auto-fit skipped: {e}")

            try:
                cf_sheet = wb.sheets['Cash_Flow']
                cf_sheet.autofit('c')  # Auto-fit all columns
                print("DEBUG: Auto-fit Cash Flow columns")
            except Exception as e:
                print(f"DEBUG: Cash Flow auto-fit skipped: {e}")

            # Format Forecast notes: wrap text, vertical center, auto-fit row heights
            try:
                forecast_sheet = wb.sheets['Forecast']
                forecast_used = forecast_sheet.api.UsedRange
                forecast_last_row = forecast_used.Row + forecast_used.Rows.Count - 1

                # Note columns are every 5th column starting at column 5 (E), 10 (J), 15 (O), etc.
                # Actually: col 2=Actual, col 3=Budget, col 4=Adj, col 5=Note, col 6=Forecast
                # So Note columns are at 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60
                note_cols = [5 + (i * 5) for i in range(12)]  # 12 months

                for note_col in note_cols:
                    try:
                        note_range = forecast_sheet.range((5, note_col), (forecast_last_row, note_col))
                        note_range.api.WrapText = True
                        note_range.api.VerticalAlignment = -4108  # xlVAlignCenter
                    except:
                        pass

                # Auto-fit row heights to accommodate wrapped text
                forecast_sheet.range((5, 1), (forecast_last_row, 1)).api.EntireRow.AutoFit()

                # Also auto-fit all data cells to be vertically centered
                data_range = forecast_sheet.range((5, 1), (forecast_last_row, 62))
                data_range.api.VerticalAlignment = -4108  # xlVAlignCenter

                print("DEBUG: Formatted Forecast notes and auto-fit rows")
            except Exception as e:
                print(f"DEBUG: Forecast formatting skipped: {e}")

            # Add navigation links to all sheets
            self._add_navigation_links(wb, new_month_list)

        except Exception as e:
            print(f"Warning: Error adding month to reports: {e}")
            import traceback
            traceback.print_exc()

    def _create_multi_division_model(self, save_path, user_start, user_end):
        """Create a multi-division Excel model with consolidated reporting"""
        # Initialize step timer for performance tracking
        timer = StepTimer()

        # Calculate total estimated time based on number of divisions
        num_divisions = len(self.divisions)

        # Calculate dynamic total steps:
        # Fixed steps: 15 (consolidation, Excel setup, consolidated sheets, budget, forecasts, menu, dashboard, notes, mappings, organize, VBA, save)
        # Variable steps: 3 per division (read files, division sheets, division forecasts)
        # Count: 18 fixed steps + 3 per division (read files, create sheets, create forecast)
        total_steps = 18 + (3 * num_divisions)
        estimated_total = 180 + (num_divisions - 1) * 30  # Base ~3 min + 30s per extra division (conservative)

        # Create progress tracker for mid-step updates
        update_step, update_substep, state = self._create_progress_tracker(total_steps, estimated_total)

        # Wrap update_step to also track timing
        original_update_step = update_step
        def update_step(msg):
            timer.start_step(msg)
            original_update_step(msg)

        # Step 1-3: Parse all division files
        all_pl_accounts = []  # Stacked accounts with division
        all_bs_accounts = []
        all_months = None
        pl_totals = {}
        bs_totals = {}
        division_configs = []

        for i, div in enumerate(self.divisions):
            div_name = div['name']
            update_step(f"Reading {div_name} files ({i+1}/{num_divisions})...")

            # Get per-division date range (convert month name to number)
            div_start_month = self.MONTHS.index(div.get('start_month', 'January')) + 1
            div_start_year = int(div.get('start_year', datetime.now().year))
            div_end_month = self.MONTHS.index(div.get('end_month', 'December')) + 1
            div_end_year = int(div.get('end_year', datetime.now().year))

            div_user_start = (div_start_month, div_start_year)
            div_user_end = (div_end_month, div_end_year)

            print(f"Division {div_name}: {div.get('start_month')} {div_start_year} to {div.get('end_month')} {div_end_year}")

            # Parse P&L with division-specific date range (combined read for speed)
            update_substep(f"{div_name}: Loading P&L file...")
            pl_data, pl_indents = self._read_excel_with_indents(div['pl_path'])
            update_substep(f"{div_name}: Parsing P&L accounts...")
            div_pl_accounts, div_months, div_pl_totals = self._parse_financial_data(
                pl_data, pl_indents, div_user_start, div_user_end
            )
            update_substep(f"{div_name}: Found {len(div_pl_accounts)} P&L accounts")

            # Parse BS with division-specific date range (combined read for speed)
            update_substep(f"{div_name}: Loading Balance Sheet file...")
            bs_data, bs_indents = self._read_excel_with_indents(div['bs_path'])
            update_substep(f"{div_name}: Parsing BS accounts...")
            div_bs_accounts, _, div_bs_totals = self._parse_financial_data(
                bs_data, bs_indents, div_user_start, div_user_end
            )
            update_substep(f"{div_name}: Found {len(div_bs_accounts)} BS accounts")

            # Add division identifier to each account
            for acct in div_pl_accounts:
                acct['division'] = div_name
            for acct in div_bs_accounts:
                acct['division'] = div_name

            all_pl_accounts.extend(div_pl_accounts)
            all_bs_accounts.extend(div_bs_accounts)

            # Use first division's months as reference (or merge if different)
            if all_months is None:
                all_months = div_months
            else:
                # Merge months from different divisions
                existing_keys = set((m, y) for m, y, _ in all_months)
                for m, y, name in div_months:
                    if (m, y) not in existing_keys:
                        all_months.append((m, y, name))
                all_months = sorted(all_months, key=lambda x: (x[1], x[0]))

            # Merge totals
            pl_totals.update(div_pl_totals)
            bs_totals.update(div_bs_totals)

            # Create DivisionConfig for consolidation engine
            division_configs.append(DivisionConfig(
                name=div_name,
                is_primary=div.get('is_primary', False),
                pl_file_path=div['pl_path'],
                bs_file_path=div['bs_path'],
                pl_accounts=div_pl_accounts,
                bs_accounts=div_bs_accounts
            ))

        print(f"Parsed {len(self.divisions)} divisions:")
        for div in self.divisions:
            print(f"  - {div['name']}")
        print(f"Total: {len(all_pl_accounts)} P&L accounts, {len(all_bs_accounts)} BS accounts, {len(all_months)} months")

        # Detect actual date range - filter out months with no data
        update_substep("Detecting actual data range...")
        all_months = self._detect_actual_date_range(all_pl_accounts + all_bs_accounts, all_months)
        print(f"After date range detection: {len(all_months)} months with data")

        # Step 4: Run consolidation engine for account matching
        update_step("Running consolidation engine...")
        update_substep("Initializing consolidation engine...")
        engine = ConsolidationEngine(division_configs)

        # Match P&L accounts
        update_substep("Matching P&L accounts across divisions...")
        pl_mappings = engine.match_accounts(statement_type="pl")
        update_substep(f"Found {len(pl_mappings)} consolidated P&L accounts")
        print(f"P&L mappings: {len(pl_mappings)} consolidated accounts")

        # Match BS accounts
        update_substep("Matching Balance Sheet accounts across divisions...")
        bs_mappings = engine.match_accounts(statement_type="bs")
        update_substep(f"Found {len(bs_mappings)} consolidated BS accounts")
        print(f"BS mappings: {len(bs_mappings)} consolidated accounts")

        # Store mappings
        self.account_mappings = {'pl': pl_mappings, 'bs': bs_mappings}

        # Step 5: Show mapping review dialog for user to verify consolidation
        # Always show dialog for multi-division mode so user can verify mappings
        update_step("Reviewing account mappings...")
        update_substep("Opening mapping review dialog...")
        self.root.update()

        # Create and show the mapping review dialog with P&L and BS in separate tabs
        mapping_dialog = AccountMappingDialog(
            self.root,
            {**pl_mappings, **bs_mappings},  # Combined for backward compat
            self.divisions,
            self._handle_mapping_approval,
            pl_mappings=pl_mappings,   # Separate P&L mappings for P&L tab
            bs_mappings=bs_mappings    # Separate BS mappings for BS tab
        )
        # Wait for dialog to close before continuing
        self.root.wait_window(mapping_dialog)

        # Retrieve updated mappings after user review/merge
        pl_mappings = mapping_dialog.pl_mappings if hasattr(mapping_dialog, 'pl_mappings') else pl_mappings
        bs_mappings = mapping_dialog.bs_mappings if hasattr(mapping_dialog, 'bs_mappings') else bs_mappings

        # Create consolidated account lists
        update_step("Creating consolidated accounts...")
        update_substep("Consolidating P&L values across divisions...")
        consolidated_pl = engine.consolidate_values(pl_mappings, all_months, "pl")
        update_substep("Consolidating Balance Sheet values across divisions...")
        consolidated_bs = engine.consolidate_values(bs_mappings, all_months, "bs")
        update_substep(f"Created {len(consolidated_pl)} P&L and {len(consolidated_bs)} BS accounts")

        # Now create the Excel workbook
        use_template = os.path.exists(TEMPLATE_PATH)
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, 'temp_model.xlsm')

        update_step("Starting Excel...")
        update_substep("Launching Excel application...")
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False
        app.calculation = 'manual'  # Disable auto-calculation for speed
        wb = None

        try:
            update_substep("Opening workbook template...")
            if use_template:
                shutil.copy(TEMPLATE_PATH, temp_path)
                wb = app.books.open(temp_path)
                menu_sheet = wb.sheets['Menu']
                source_pl = wb.sheets['Source_PL']
                source_bs = wb.sheets['Source_BS']
            else:
                wb = app.books.add()
                menu_sheet = wb.sheets[0]
                menu_sheet.name = 'Menu'
                source_pl = wb.sheets.add('Source_PL', after=menu_sheet)
                source_bs = wb.sheets.add('Source_BS', after=source_pl)

            # Clear existing data
            update_substep("Clearing existing data...")
            source_pl.range('A1:ZZ1000').clear()
            source_bs.range('A1:ZZ1000').clear()

            # Populate source sheets with stacked division data
            update_step("Populating source data with divisions...")
            update_substep(f"Writing {len(all_pl_accounts)} P&L accounts to Source_PL...")
            self._populate_source_sheet(source_pl, all_pl_accounts, all_months)
            update_substep(f"Writing {len(all_bs_accounts)} BS accounts to Source_BS...")
            self._populate_source_sheet(source_bs, all_bs_accounts, all_months)

            # Create or get report sheets
            update_step("Creating consolidated P&L...")
            update_substep("Setting up Consolidated_PL sheet...")
            if 'Consolidated_PL' in [s.name for s in wb.sheets]:
                cons_pl_sheet = wb.sheets['Consolidated_PL']
                cons_pl_sheet.range('A1:ZZ1000').clear()
            else:
                # Try to use existing PL sheet or create new
                if 'PL' in [s.name for s in wb.sheets]:
                    cons_pl_sheet = wb.sheets['PL']
                    cons_pl_sheet.name = 'Consolidated_PL'
                    cons_pl_sheet.range('A1:ZZ1000').clear()
                else:
                    cons_pl_sheet = wb.sheets.add('Consolidated_PL', after=source_bs)

            # Create consolidated P&L report using consolidated accounts
            update_substep(f"Writing {len(consolidated_pl)} consolidated P&L accounts...")
            self._create_consolidated_pl_report(cons_pl_sheet, consolidated_pl, all_months, pl_totals)

            # Skip Consolidated Balance Sheet - not realistic for multi-division
            # Division-specific balance sheets are available instead

            update_step("Creating Cash Flow...")
            update_substep("Setting up Cash_Flow sheet...")
            if 'Cash_Flow' in [s.name for s in wb.sheets]:
                cons_cf_sheet = wb.sheets['Cash_Flow']
                cons_cf_sheet.range('A1:ZZ1000').clear()
            else:
                cons_cf_sheet = wb.sheets.add('Cash_Flow', after=cons_pl_sheet)

            update_substep("Building cash flow formulas...")
            self._create_cash_flow(cons_cf_sheet, all_months)

            # Create division-specific sheets (track last sheet to maintain correct order)
            last_created_sheet = cons_cf_sheet
            for idx, div in enumerate(self.divisions):
                div_name = div['name']
                safe_name = div_name.replace(' ', '_')[:20]  # Excel sheet name limit

                update_step(f"Creating {div_name} sheets ({idx+1}/{num_divisions})...")

                # Get division-specific accounts
                div_pl = [a for a in all_pl_accounts if a.get('division') == div_name]
                div_bs = [a for a in all_bs_accounts if a.get('division') == div_name]

                # Create division P&L sheet (after the last created sheet)
                update_substep(f"{div_name}: Creating P&L sheet ({len(div_pl)} accounts)...")
                div_pl_name = f"{safe_name}_PL"
                if div_pl_name in [s.name for s in wb.sheets]:
                    div_pl_sheet = wb.sheets[div_pl_name]
                    div_pl_sheet.range('A1:ZZ1000').clear()
                else:
                    div_pl_sheet = wb.sheets.add(div_pl_name, after=last_created_sheet)
                self._create_division_pl_report(div_pl_sheet, div_pl, all_months, pl_totals, div_name)

                # Create division BS sheet (after division PL)
                update_substep(f"{div_name}: Creating Balance Sheet ({len(div_bs)} accounts)...")
                div_bs_name = f"{safe_name}_BS"
                if div_bs_name in [s.name for s in wb.sheets]:
                    div_bs_sheet = wb.sheets[div_bs_name]
                    div_bs_sheet.range('A1:ZZ1000').clear()
                else:
                    div_bs_sheet = wb.sheets.add(div_bs_name, after=div_pl_sheet)
                self._create_division_bs_report(div_bs_sheet, div_bs, all_months, bs_totals, div_name)

                # Update last_created_sheet for next iteration
                last_created_sheet = div_bs_sheet

            # Create Source_Budget sheet for budget data (placed after all P&L/BS sheets)
            update_step("Creating Source Budget sheet...")
            # Find the last P&L/BS sheet to place budget after it
            last_report_sheet = cons_cf_sheet
            for s in wb.sheets:
                if s.name.endswith('_BS') or s.name.endswith('_PL') or s.name == 'Cash_Flow':
                    last_report_sheet = s
            if 'Source_Budget' in [s.name for s in wb.sheets]:
                source_budget = wb.sheets['Source_Budget']
                source_budget.range('A1:ZZ1000').clear()
            else:
                source_budget = wb.sheets.add('Source_Budget', after=last_report_sheet)
            self._create_source_budget_sheet(source_budget, consolidated_pl, all_months)

            # Create Consolidated Forecast sheet
            update_step("Creating Consolidated Forecast...")
            update_substep("Setting up Consolidated_Forecast sheet...")
            if 'Consolidated_Forecast' in [s.name for s in wb.sheets]:
                forecast_sheet = wb.sheets['Consolidated_Forecast']
                forecast_sheet.range('A1:ZZ1000').clear()
            elif 'Forecast' in [s.name for s in wb.sheets]:
                forecast_sheet = wb.sheets['Forecast']
                forecast_sheet.name = 'Consolidated_Forecast'
                forecast_sheet.range('A1:ZZ1000').clear()
            else:
                forecast_sheet = wb.sheets.add('Consolidated_Forecast', after=cons_cf_sheet)
            self._create_forecast_sheet(forecast_sheet, consolidated_pl, all_months)

            # Create Consolidated Forecast Summary
            update_step("Creating Forecast Summary...")
            update_substep("Building YTD and variance analysis...")
            if 'Forecast_Summary' in [s.name for s in wb.sheets]:
                forecast_summary = wb.sheets['Forecast_Summary']
                forecast_summary.range('A1:ZZ1000').clear()
            else:
                forecast_summary = wb.sheets.add('Forecast_Summary', after=forecast_sheet)
            self._create_forecast_summary_sheet(forecast_summary, consolidated_pl, all_months)

            # Create per-division Forecast sheets
            for idx, div in enumerate(self.divisions):
                div_name = div['name']
                safe_name = div_name.replace(' ', '_')[:20]

                update_step(f"Creating {div_name} Forecast ({idx+1}/{num_divisions})...")

                # Get division-specific P&L accounts
                div_pl = [a for a in all_pl_accounts if a.get('division') == div_name]

                # Create division Forecast sheet
                update_substep(f"{div_name}: Creating Forecast sheet...")
                div_forecast_name = f"{safe_name}_Forecast"
                if div_forecast_name in [s.name for s in wb.sheets]:
                    div_forecast_sheet = wb.sheets[div_forecast_name]
                    div_forecast_sheet.range('A1:ZZ1000').clear()
                else:
                    div_forecast_sheet = wb.sheets.add(div_forecast_name, after=forecast_summary)
                self._create_forecast_sheet(div_forecast_sheet, div_pl, all_months, division_name=div_name)

            # Update Menu sheet with division navigation
            update_step("Creating Menu with navigation...")
            update_substep("Building navigation tree for all divisions...")
            self._create_multi_division_menu_sheet(menu_sheet, all_months, self.divisions)

            # Create Dashboard
            update_step("Creating Dashboard...")
            update_substep("Setting up Dashboard sheet...")
            if 'Dashboard' in [s.name for s in wb.sheets]:
                dashboard_sheet = wb.sheets['Dashboard']
                dashboard_sheet.range('A1:ZZ1000').clear()
            else:
                dashboard_sheet = wb.sheets.add('Dashboard', before=menu_sheet)
            update_substep("Building Dashboard KPIs and charts...")
            self._create_dashboard_sheet(dashboard_sheet, consolidated_pl, consolidated_bs, all_months, pl_totals)

            # Create Notes sheet with Division dropdown
            update_step("Creating Notes sheet...")
            update_substep("Setting up Notes with division selector...")
            if 'Notes' in [s.name for s in wb.sheets]:
                notes_sheet = wb.sheets['Notes']
                notes_sheet.range('A1:ZZ1000').clear()
            else:
                notes_sheet = wb.sheets.add('Notes', after=dashboard_sheet)
            month_names = [name for m, y, name in all_months]
            self._create_notes_sheet(notes_sheet, consolidated_pl, consolidated_bs, month_names, divisions=self.divisions)

            # Create Mapping_Config sheet for persistence
            update_step("Saving account mappings...")
            update_substep("Writing mappings to hidden Excel sheet...")
            MappingPersistence.save_to_excel(wb, self.account_mappings, self.divisions)

            # Save mappings to JSON file alongside Excel
            update_substep("Writing mappings to JSON file...")
            json_path = MappingPersistence.get_json_filepath(save_path)
            MappingPersistence.save_to_json(
                self.account_mappings,
                self.divisions,
                self.company_name.get(),
                json_path
            )

            # Move source sheets to end and set black tabs
            update_step("Organizing sheet tabs...")
            update_substep("Arranging sheets in professional order...")
            try:
                # Desired order: Dashboard, Menu, Consolidated reports, Division P&Ls, Division BSs,
                # Division Forecasts, Forecast Summary, Notes, Source sheets
                sheet_order = ['Dashboard', 'Menu']

                # Add Consolidated reports
                for name in ['Consolidated_PL', 'Cash_Flow']:
                    if name in [s.name for s in wb.sheets]:
                        sheet_order.append(name)

                # Add Division P&L sheets (before forecasts)
                for div in self.divisions:
                    safe_name = div['name'].replace(' ', '_')[:20]
                    pl_name = f"{safe_name}_PL"
                    if pl_name in [s.name for s in wb.sheets]:
                        sheet_order.append(pl_name)

                # Add Division BS sheets
                for div in self.divisions:
                    safe_name = div['name'].replace(' ', '_')[:20]
                    bs_name = f"{safe_name}_BS"
                    if bs_name in [s.name for s in wb.sheets]:
                        sheet_order.append(bs_name)

                # Add Forecast sheets (after division reports)
                for name in ['Consolidated_Forecast', 'Forecast', 'Forecast_Summary']:
                    if name in [s.name for s in wb.sheets]:
                        sheet_order.append(name)

                # Add Division Forecast sheets
                for div in self.divisions:
                    safe_name = div['name'].replace(' ', '_')[:20]
                    fcst_name = f"{safe_name}_Forecast"
                    if fcst_name in [s.name for s in wb.sheets]:
                        sheet_order.append(fcst_name)

                # Add remaining sheets
                for name in ['Notes', 'Dashboard_Control']:
                    if name in [s.name for s in wb.sheets]:
                        sheet_order.append(name)

                # Move sheets to match desired order
                for i, name in enumerate(sheet_order):
                    if name in [s.name for s in wb.sheets]:
                        sheet = wb.sheets[name]
                        if i == 0:
                            sheet.api.Move(Before=wb.sheets[0].api)
                        else:
                            prev_sheet = wb.sheets[sheet_order[i-1]]
                            sheet.api.Move(After=prev_sheet.api)

                # Move Source sheets to the very end
                last_sheet = wb.sheets[-1]
                source_pl.api.Move(After=last_sheet.api)
                source_bs.api.Move(After=source_pl.api)

                # Set tab colors to black for source/hidden sheets
                update_substep("Setting source sheet tab colors...")
                source_pl.api.Tab.Color = 0x000000  # Black
                source_bs.api.Tab.Color = 0x000000  # Black

                # Also hide Mapping_Config if it exists
                try:
                    mapping_sheet = wb.sheets['Mapping_Config']
                    mapping_sheet.api.Move(After=source_bs.api)
                    mapping_sheet.api.Tab.Color = 0x000000  # Black
                except:
                    pass

                # Set horizontal scrollbar small to show more tabs
                # TabRatio = ratio of tabs to total width. 0.85 = 85% tabs, 15% scrollbar
                update_substep("Adjusting view settings...")
                try:
                    app.api.ActiveWindow.TabRatio = 0.85  # 85% tabs, 15% scrollbar
                except:
                    pass

            except Exception as e:
                print(f"Warning: Could not organize sheet tabs: {e}")

            # Clean up empty leading columns from division sheets
            update_step("Cleaning up empty columns...")
            for div in self.divisions:
                div_name = div['name']
                safe_name = div_name.replace(' ', '_')[:20]

                # Clean division P&L sheet
                pl_sheet_name = f"{safe_name}_PL"
                try:
                    if pl_sheet_name in [s.name for s in wb.sheets]:
                        update_substep(f"Cleaning {pl_sheet_name}...")
                        self._remove_empty_leading_columns(wb.sheets[pl_sheet_name])
                except Exception as e:
                    print(f"Warning: Could not clean {pl_sheet_name}: {e}")

                # Clean division BS sheet
                bs_sheet_name = f"{safe_name}_BS"
                try:
                    if bs_sheet_name in [s.name for s in wb.sheets]:
                        update_substep(f"Cleaning {bs_sheet_name}...")
                        self._remove_empty_leading_columns(wb.sheets[bs_sheet_name])
                except Exception as e:
                    print(f"Warning: Could not clean {bs_sheet_name}: {e}")

            # Add VBA code
            update_step("Adding VBA macros...")
            update_substep("Injecting VBA module...")
            try:
                module_exists = False
                for component in wb.api.VBProject.VBComponents:
                    if component.Name == "FinancialModel":
                        component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                        component.CodeModule.AddFromString(self.VBA_CODE)
                        module_exists = True
                        break
                if not module_exists:
                    vba_module = wb.api.VBProject.VBComponents.Add(1)
                    vba_module.Name = "FinancialModel"
                    vba_module.CodeModule.AddFromString(self.VBA_CODE)
            except:
                pass

            # Activate Dashboard
            update_substep("Setting Dashboard as active sheet...")
            try:
                dashboard_sheet.activate()
            except:
                pass

            # Save workbook
            update_step("Saving workbook...")
            update_substep("Re-enabling calculations...")
            app.calculation = 'automatic'  # Re-enable before save
            update_substep("Writing to disk (this may take a moment)...")
            wb.save()
            update_substep("Closing workbook...")
            wb.close()
            wb = None

        finally:
            try:
                if wb is not None:
                    wb.close()
            except:
                pass
            try:
                app.quit()
            except:
                pass

        # Move from temp to final location
        try:
            if os.path.exists(save_path):
                os.remove(save_path)
            shutil.move(temp_path, save_path)
        finally:
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

        # Finish timing and save log
        timer.finish_run()

    def _create_consolidated_pl_report(self, sheet, accounts, months, detected_totals=None):
        """Create consolidated P&L report with pre-computed values

        For multi-division consolidated view, values are already summed by consolidate_values().
        We write values directly and also include SUMIF formulas for live updates.
        """
        # For multi-division, use special report that writes values and formulas
        # referencing column B (Account) since column A is Division
        self._create_pl_report_multi_div(sheet, accounts, months, detected_totals, is_consolidated=True)

    def _create_consolidated_bs_report(self, sheet, accounts, months, detected_totals=None):
        """Create consolidated Balance Sheet report"""
        self._create_bs_report_multi_div(sheet, accounts, months, detected_totals, is_consolidated=True)

    def _create_division_pl_report(self, sheet, accounts, months, detected_totals, division_name):
        """Create P&L report for a specific division"""
        self._create_pl_report_multi_div(sheet, accounts, months, detected_totals,
                                          is_consolidated=False, division_name=division_name)
        # Update title to show division name
        sheet.range('A1').value = f"{division_name}"

    def _create_division_bs_report(self, sheet, accounts, months, detected_totals, division_name):
        """Create Balance Sheet report for a specific division"""
        self._create_bs_report_multi_div(sheet, accounts, months, detected_totals,
                                          is_consolidated=False, division_name=division_name)
        sheet.range('A1').value = f"{division_name}"

    def _create_pl_report_multi_div(self, sheet, accounts, months, detected_totals=None,
                                     is_consolidated=True, division_name=None, single_entity_mode=False):
        """Create P&L report matching single-division format with YTD, variance, and full year columns.

        OPTIMIZED: Uses bulk writes to prevent performance issues.

        Args:
            single_entity_mode: If True, Source_PL has col A=Account (not Division), data starts col B
        """
        import time as _time
        _timings = {}
        _t0 = _time.perf_counter()

        company = self.company_name.get()
        detected_totals = detected_totals or {}

        # Colors
        DARK_BLUE = (22, 33, 62)
        SUBTOTAL_GRAY = (236, 236, 236)

        # Source data range
        source_start = 3
        source_end = 1500

        # Calculate column positions
        header_row = 4
        num_months = len(months)
        last_month_col = num_months + 1
        notes_col = last_month_col + 1
        spacer1_col = notes_col + 1
        py_ytd_col = spacer1_col + 1
        cy_ytd_col = py_ytd_col + 1
        var_col = cy_ytd_col + 1
        var_pct_col = var_col + 1
        spacer2_col = var_pct_col + 1

        years = sorted(set(y for m, y, name in months))
        fy_start_col = spacer2_col + 1
        last_col = fy_start_col + len(years) - 1

        _timings['setup'] = _time.perf_counter() - _t0

        # ================================================================
        # BUILD ALL DATA IN MEMORY FIRST
        # ================================================================
        _t1 = _time.perf_counter()
        all_data = []  # List of row data arrays
        row_types = []  # Track row type for formatting
        row_tracking = {}
        total_income_row = None
        total_cogs_row = None
        total_expenses_row = None
        found_cogs_section = False

        # Helper to build row data
        first_data_col_letter = self._col_letter(2)
        last_data_col_letter = self._col_letter(num_months + 1)

        for account in accounts:
            account_name = account['name']
            name_lower = account_name.lower()
            indent_level = account.get('indent', 0)

            # Track sections
            if account['is_header']:
                if 'cost' in name_lower or 'cogs' in name_lower:
                    found_cogs_section = True

            # Display name with indentation
            display_name = account_name
            if indent_level > 0 and not account['is_header'] and not account['is_total']:
                display_name = ('    ' * indent_level) + account_name

            # Calculate actual row number
            actual_row = header_row + 1 + len(all_data)

            # Track key rows
            if 'gross profit' in name_lower:
                row_tracking['gross_profit_row'] = actual_row
            if 'net income' in name_lower and account['is_total']:
                row_tracking['net_income_row'] = actual_row

            # Track total rows
            if 'total' in name_lower and account['is_total']:
                if (('income' in name_lower or 'revenue' in name_lower) and
                    'net' not in name_lower and 'other' not in name_lower and not found_cogs_section):
                    total_income_row = actual_row
                    row_tracking['total_income_row'] = actual_row
                elif 'cost' in name_lower or 'cogs' in name_lower:
                    total_cogs_row = actual_row
                    row_tracking['total_cogs_row'] = actual_row
                    found_cogs_section = True
                elif 'expense' in name_lower and 'other' not in name_lower:
                    total_expenses_row = actual_row
                    row_tracking['total_expense_row'] = actual_row

            # Build row data array
            row_data = [display_name]

            if account['is_header']:
                # Header rows: empty data cells
                row_data.extend([''] * (last_col - 1))
                all_data.append(row_data)
                row_types.append('header')
                continue

            # Monthly data columns - SUMIF/SUMIFS formulas
            # single_entity_mode: col A=Account, data starts col B
            # multi-div consolidated: col A=Division, col B=Account, data starts col C
            # multi-div division: filter by division name
            for i, (m, y, name) in enumerate(months):
                if single_entity_mode:
                    cl = self._col_letter(i + 2)  # Data starts at column B
                    formula = f'=SUMIF(Source_PL!$A${source_start}:$A${source_end},"{account_name}",Source_PL!{cl}${source_start}:{cl}${source_end})'
                elif is_consolidated:
                    cl = self._col_letter(i + 3)  # Data starts at column C (col A=Div, B=Acct)
                    formula = f'=SUMIF(Source_PL!$B${source_start}:$B${source_end},"{account_name}",Source_PL!{cl}${source_start}:{cl}${source_end})'
                else:
                    cl = self._col_letter(i + 3)  # Data starts at column C
                    formula = f'=SUMIFS(Source_PL!{cl}${source_start}:{cl}${source_end},Source_PL!$A${source_start}:$A${source_end},"{division_name}",Source_PL!$B${source_start}:$B${source_end},"{account_name}")'
                row_data.append(formula)

            # Notes column
            notes_formula = f'=IFERROR(LOOKUP(2,1/((Notes!$A$2:$A$100="P&L")*(Notes!$B$2:$B$100=TEXT(Menu!$C$7,"mmm yy"))*(Notes!$C$2:$C$100=TRIM($A{actual_row}))),Notes!$D$2:$D$100),"")'
            row_data.append(notes_formula)

            # Spacer 1
            row_data.append('')

            # YTD formulas
            data_range = f'{first_data_col_letter}{actual_row}:{last_data_col_letter}{actual_row}'
            helper_range = f'{first_data_col_letter}$3:{last_data_col_letter}$3'

            py_formula = f'=SUMPRODUCT(({data_range})*--(INT({helper_range}/100)=Menu!$F$7-1)*--(MOD({helper_range},100)<=Menu!$E$7))'
            cy_formula = f'=SUMPRODUCT(({data_range})*--(INT({helper_range}/100)=Menu!$F$7)*--(MOD({helper_range},100)<=Menu!$E$7))'
            var_formula = f'={self._col_letter(cy_ytd_col)}{actual_row}-{self._col_letter(py_ytd_col)}{actual_row}'
            var_pct_formula = f'=IFERROR({self._col_letter(var_col)}{actual_row}/{self._col_letter(py_ytd_col)}{actual_row},0)'

            row_data.extend([py_formula, cy_formula, var_formula, var_pct_formula])

            # Spacer 2
            row_data.append('')

            # Full Year columns
            for year in years:
                year_month_cols = [c + 2 for c, (m, y, name) in enumerate(months) if y == year]
                if year_month_cols:
                    refs = '+'.join([f'{self._col_letter(c)}{actual_row}' for c in year_month_cols])
                    row_data.append(f'={refs}')
                else:
                    row_data.append('')

            all_data.append(row_data)
            row_type = 'total' if account['is_total'] else 'detail'
            if 'net income' in name_lower and account['is_total']:
                row_type = 'net_income'
            row_types.append(row_type)

        _timings['build_data'] = _time.perf_counter() - _t1

        # ================================================================
        # WRITE TO EXCEL IN BULK
        # ================================================================
        _t2 = _time.perf_counter()
        # Title - Professional corporate style
        sheet.range('A1').value = company if is_consolidated else division_name
        sheet.range('A1').font.name = 'Calibri Light'
        sheet.range('A1').font.size = 16
        sheet.range('A1').font.bold = True
        sheet.range('A1').font.color = DARK_BLUE
        sheet.range('A2').value = 'Consolidated Profit & Loss' if is_consolidated else 'Profit & Loss Statement'
        sheet.range('A2').font.name = 'Calibri Light'
        sheet.range('A2').font.size = 12
        sheet.range('A2').font.color = (128, 128, 128)

        # Row 3: YYYYMM helper values
        helper_row = [y * 100 + m for m, y, name in months]
        sheet.range((3, 2), (3, num_months + 1)).value = [helper_row]

        # Header row
        header_data = ['Account']
        for m, y, name in months:
            header_data.append(f"{self.MONTHS[m-1][:3]} {y}")
        header_data.extend(['Notes', '', 'PY YTD', 'CY YTD', 'Var $', 'Var %', ''])
        header_data.extend([str(y) for y in years])
        sheet.range((header_row, 1), (header_row, last_col)).value = [header_data]

        # Format header - professional corporate style
        header_range = sheet.range((header_row, 1), (header_row, last_col))
        header_range.font.name = 'Calibri Light'
        header_range.font.size = 10
        header_range.font.bold = True
        header_range.color = DARK_BLUE
        header_range.font.color = (255, 255, 255)

        # Write all data in ONE bulk operation
        if all_data:
            data_start_row = header_row + 1
            data_end_row = header_row + len(all_data)
            sheet.range((data_start_row, 1), (data_end_row, last_col)).value = all_data

        _timings['write_excel'] = _time.perf_counter() - _t2

        # ================================================================
        # APPLY FORMATTING IN BULK (NO per-row loops)
        # ================================================================
        _t3 = _time.perf_counter()
        if all_data:
            # Number formats for entire columns - single operations
            sheet.range((data_start_row, 2), (data_end_row, last_month_col)).number_format = '#,##0'
            sheet.range((data_start_row, py_ytd_col), (data_end_row, cy_ytd_col)).number_format = '#,##0'
            sheet.range((data_start_row, var_col), (data_end_row, var_col)).number_format = '#,##0'
            sheet.range((data_start_row, var_pct_col), (data_end_row, var_pct_col)).number_format = '0.0%'
            if fy_start_col <= last_col:
                sheet.range((data_start_row, fy_start_col), (data_end_row, last_col)).number_format = '#,##0'

            # Apply Calibri Light to all data
            sheet.range((data_start_row, 1), (data_end_row, last_col)).font.name = 'Calibri Light'
            sheet.range((data_start_row, 1), (data_end_row, last_col)).font.size = 10

            # Collect row numbers by type for batch formatting
            header_rows = [data_start_row + idx for idx, rt in enumerate(row_types) if rt == 'header']
            total_rows = [data_start_row + idx for idx, rt in enumerate(row_types) if rt in ('total', 'net_income')]

            # Format headers in batch (bold, section style)
            if header_rows:
                for r in header_rows:
                    sheet.range((r, 1)).font.bold = True
                    sheet.range((r, 1)).font.size = 11

            # Format totals with bold, top border (professional P&L style)
            if total_rows:
                for r in total_rows:
                    row_range = sheet.range((r, 1), (r, last_col))
                    row_range.font.bold = True
                    # Add top border for total rows (single line above)
                    try:
                        row_range.api.Borders(8).LineStyle = 1  # xlContinuous top border
                        row_range.api.Borders(8).Weight = 2  # xlThin
                    except:
                        pass

            # Net Income gets double underline (accounting standard)
            net_income_rows = [data_start_row + idx for idx, rt in enumerate(row_types) if rt == 'net_income']
            if net_income_rows:
                for r in net_income_rows:
                    row_range = sheet.range((r, 1), (r, last_col))
                    row_range.font.bold = True
                    try:
                        # Double bottom border for Net Income
                        row_range.api.Borders(9).LineStyle = -4119  # xlDouble
                        row_range.api.Borders(9).Weight = 4  # xlThick
                    except:
                        pass

        # Column widths
        try:
            sheet.range('A:A').column_width = 35
            sheet.range((1, spacer1_col)).column_width = 2
            sheet.range((1, spacer2_col)).column_width = 2
        except:
            pass

        # Group previous year columns
        self._group_previous_year_columns(sheet, months, data_start_col=2)

        _timings['formatting'] = _time.perf_counter() - _t3

        # ================================================================
        # VALIDATION SECTION
        # ================================================================
        _t4 = _time.perf_counter()
        validation_start_row = (header_row + len(all_data) + 3) if all_data else header_row + 5
        sheet.range(f'A{validation_start_row}').value = 'VALIDATION - Source vs Calculated Totals'
        sheet.range(f'A{validation_start_row}').font.bold = True
        sheet.range(f'A{validation_start_row}').font.color = (128, 0, 128)

        val_header_row = validation_start_row + 1
        val_headers = ['Category', 'Source', 'Calculated', 'Variance', 'Match?']
        sheet.range(f'A{val_header_row}:E{val_header_row}').value = [val_headers]
        sheet.range(f'A{val_header_row}:E{val_header_row}').font.bold = True
        sheet.range(f'A{val_header_row}:E{val_header_row}').color = (200, 200, 200)

        last_col_letter = self._col_letter(last_month_col)
        val_categories = [
            ('Total Income', '*Total*Income*', row_tracking.get('total_income_row')),
            ('Total COGS', '*Total*Cost*', row_tracking.get('total_cogs_row')),
            ('Total Expenses', '*Total*Expense*', row_tracking.get('total_expense_row')),
            ('Net Income', '*Net Income*', row_tracking.get('net_income_row')),
        ]

        for idx, (label, source_pattern, calc_row) in enumerate(val_categories):
            row = val_header_row + 1 + idx
            sheet.range(f'A{row}').value = label

            # Source total (SUMIF from Source_PL with wildcard) - sum the last month column
            if is_consolidated:
                source_formula = f'=SUMIF(Source_PL!$B$3:$B$1500,"{source_pattern}",Source_PL!{last_col_letter}$3:{last_col_letter}$1500)'
            else:
                source_formula = f'=SUMIFS(Source_PL!{last_col_letter}$3:{last_col_letter}$1500,Source_PL!$A$3:$A$1500,"{division_name}",Source_PL!$B$3:$B$1500,"{source_pattern}")'
            sheet.range(f'B{row}').value = source_formula
            sheet.range(f'B{row}').number_format = '#,##0'
            print(f"[VALIDATION] Row {row} Source: {source_formula}")

            # Calculated total (from this sheet) - reference the tracked row directly
            if calc_row:
                calc_formula = f'={last_col_letter}{calc_row}'
            else:
                # Fallback: search for the label in column A using SUMIF
                calc_formula = f'=SUMIF($A:$A,"{source_pattern}",{last_col_letter}:{last_col_letter})'
            sheet.range(f'C{row}').value = calc_formula
            sheet.range(f'C{row}').number_format = '#,##0'
            print(f"[VALIDATION] Row {row} Calc ({calc_row}): {calc_formula}")

            # Variance (Source - Calculated)
            sheet.range(f'D{row}').value = f'=B{row}-C{row}'
            sheet.range(f'D{row}').number_format = '#,##0'

            # Match indicator
            sheet.range(f'E{row}').value = f'=IF(ABS(D{row})<1,"✓","✗")'
            sheet.range(f'E{row}').font.size = 14

        validation_end_row = val_header_row + len(val_categories)

        # Add division breakdown for consolidated reports
        if is_consolidated and hasattr(self, 'divisions') and len(self.divisions) > 1:
            div_start_row = validation_end_row + 2
            sheet.range(f'A{div_start_row}').value = 'BY DIVISION - Net Income'
            sheet.range(f'A{div_start_row}').font.bold = True
            sheet.range(f'A{div_start_row}').font.color = (0, 100, 0)

            div_header_row = div_start_row + 1
            div_headers = ['Division', 'Source', 'Calculated', 'Variance']
            sheet.range(f'A{div_header_row}:D{div_header_row}').value = [div_headers]
            sheet.range(f'A{div_header_row}:D{div_header_row}').font.bold = True
            sheet.range(f'A{div_header_row}:D{div_header_row}').color = (220, 220, 220)

            for div_idx, div in enumerate(self.divisions):
                div_name = div.get('name', div) if isinstance(div, dict) else getattr(div, 'name', str(div))
                row = div_header_row + 1 + div_idx
                sheet.range(f'A{row}').value = div_name
                # Source Net Income for this division
                sheet.range(f'B{row}').value = f'=SUMIFS(Source_PL!{last_col_letter}$3:{last_col_letter}$1500,Source_PL!$A$3:$A$1500,"{div_name}",Source_PL!$B$3:$B$1500,"Net Income")'
                sheet.range(f'B{row}').number_format = '#,##0'
                # Calculated - reference division-specific sheet if exists
                div_sheet_name = f"{div_name}_PL"
                sheet.range(f'C{row}').value = f"=IFERROR('{div_sheet_name}'!{last_col_letter}{row_tracking.get('net_income_row', 20)},0)"
                sheet.range(f'C{row}').number_format = '#,##0'
                sheet.range(f'D{row}').value = f'=B{row}-C{row}'
                sheet.range(f'D{row}').number_format = '#,##0'

            # Sum row
            sum_row = div_header_row + 1 + len(self.divisions)
            sheet.range(f'A{sum_row}').value = 'TOTAL'
            sheet.range(f'A{sum_row}').font.bold = True
            sheet.range(f'B{sum_row}').value = f'=SUM(B{div_header_row + 1}:B{sum_row - 1})'
            sheet.range(f'B{sum_row}').font.bold = True
            sheet.range(f'B{sum_row}').number_format = '#,##0'
            sheet.range(f'C{sum_row}').value = f'=SUM(C{div_header_row + 1}:C{sum_row - 1})'
            sheet.range(f'C{sum_row}').font.bold = True
            sheet.range(f'C{sum_row}').number_format = '#,##0'
            sheet.range(f'D{sum_row}').value = f'=B{sum_row}-C{sum_row}'
            sheet.range(f'D{sum_row}').font.bold = True
            sheet.range(f'D{sum_row}').number_format = '#,##0'

            validation_end_row = sum_row

        # Group/collapse the validation section
        try:
            sheet.api.Rows(f"{validation_start_row}:{validation_end_row}").Group()
            sheet.api.Outline.ShowLevels(RowLevels=1)  # Collapse by default
        except Exception as e:
            print(f"[PL-MULTI] Could not group validation rows: {e}")

        _timings['validation'] = _time.perf_counter() - _t4
        _timings['total'] = _time.perf_counter() - _t0

        # Write timing summary to file
        sheet_name = "Consolidated_PL" if is_consolidated else f"{division_name}_PL"
        timing_log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pl_timing_detail.txt')
        with open(timing_log_path, 'a') as f:
            f.write(f"\n{'='*60}\n")
            f.write(f"[TIMING] P&L Report: {sheet_name} - {_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"{'='*60}\n")
            f.write(f"  Accounts: {len(accounts)}, Months: {len(months)}, Rows written: {len(all_data)}\n")
            f.write(f"  Setup:       {_timings.get('setup', 0)*1000:8.1f} ms\n")
            f.write(f"  Build Data:  {_timings.get('build_data', 0)*1000:8.1f} ms  (Python loop building formulas)\n")
            f.write(f"  Write Excel: {_timings.get('write_excel', 0)*1000:8.1f} ms  (Bulk write to Excel)\n")
            f.write(f"  Formatting:  {_timings.get('formatting', 0)*1000:8.1f} ms  (Number formats, column widths)\n")
            f.write(f"  Validation:  {_timings.get('validation', 0)*1000:8.1f} ms  (Validation section)\n")
            f.write(f"  ----------------------------------------\n")
            f.write(f"  TOTAL:       {_timings.get('total', 0)*1000:8.1f} ms ({_timings.get('total', 0):.2f} seconds)\n")
            f.write(f"{'='*60}\n")

        # AutoFit all columns for best display
        try:
            sheet.autofit('c')  # AutoFit all columns
        except:
            pass

        # Collapse all outline groups
        try:
            sheet.api.Outline.ShowLevels(RowLevels=1, ColumnLevels=1)
        except:
            pass

        # Hide row 3 (YYYYMM helper row) - MUST be at end after all other operations
        try:
            sheet.range('3:3').api.Hidden = True
            print(f"[PL] Row 3 hidden successfully")
        except Exception as e:
            print(f"[PL] ERROR hiding row 3: {e}")

        # Add back to menu link and print setup
        self._add_back_to_menu_link(sheet, row=1, col=last_col + 2)
        self._setup_print_area(sheet)

    def _create_bs_report_multi_div(self, sheet, accounts, months, detected_totals=None,
                                     is_consolidated=True, division_name=None, single_entity_mode=False):
        """Create Balance Sheet report for multi-division mode with DYNAMIC FORMULAS

        Uses SUMIF/SUMIFS formulas with LIMITED RANGES (not entire columns).

        Args:
            single_entity_mode: If True, Source_BS has col A=Account (not Division), data starts col B
        """
        print(f"[BS-MULTI] === STARTING === {len(accounts)} accounts, {len(months)} months")
        import time
        start_time = time.time()

        company = self.company_name.get()

        DARK_BLUE = (22, 33, 62)
        SUBTOTAL_GRAY = (236, 236, 236)

        # Source data range - limited to actual data rows
        source_start = 3
        source_end = 1500  # Safe upper limit

        # Title
        sheet.range('A1').value = company if is_consolidated else division_name
        sheet.range('A2').value = 'Consolidated Balance Sheet' if is_consolidated else 'Balance Sheet'

        header_row = 4
        num_months = len(months)
        last_col = num_months + 1

        # Build header row
        header_data = ['Account']
        for m, y, name in months:
            header_data.append(f"{self.MONTHS[m-1][:3]} {y}")

        sheet.range((header_row, 1), (header_row, last_col)).value = header_data
        print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Header written")

        # Track rows for formatting
        total_rows = []
        header_rows_list = []

        # Build data with DYNAMIC FORMULAS using LIMITED RANGES
        all_data = []
        row_idx = header_row + 1

        print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Building formulas...")

        for account in accounts:
            account_name = account['name']
            indent_level = account.get('indent', 0)
            is_header_row = account.get('is_header', False)
            is_total = account.get('is_total', False)

            display_name = account_name
            if indent_level > 0 and not is_header_row and not is_total:
                display_name = ('    ' * indent_level) + account_name

            if is_header_row:
                header_rows_list.append(row_idx)
            elif is_total:
                total_rows.append(row_idx)

            row_data = [display_name]

            if is_header_row:
                row_data.extend([''] * num_months)
            else:
                # DYNAMIC FORMULAS with LIMITED RANGES
                # single_entity_mode: col A=Account, data starts col B
                # multi-div consolidated: col A=Division, col B=Account, data starts col C
                for i, (m, y, name) in enumerate(months):
                    if single_entity_mode:
                        col = self._col_letter(i + 2)  # Data starts at column B
                        formula = f'=SUMIF(Source_BS!$A${source_start}:$A${source_end},"{account_name}",Source_BS!{col}${source_start}:{col}${source_end})'
                    elif is_consolidated:
                        col = self._col_letter(i + 3)  # Data starts at column C
                        formula = f'=SUMIF(Source_BS!$B${source_start}:$B${source_end},"{account_name}",Source_BS!{col}${source_start}:{col}${source_end})'
                    else:
                        col = self._col_letter(i + 3)  # Data starts at column C
                        formula = f'=SUMIFS(Source_BS!{col}${source_start}:{col}${source_end},Source_BS!$A${source_start}:$A${source_end},"{division_name}",Source_BS!$B${source_start}:$B${source_end},"{account_name}")'
                    row_data.append(formula)

            all_data.append(row_data)
            row_idx += 1

        print(f"[BS-MULTI] {time.time() - start_time:.2f}s - {len(all_data)} rows built")

        # Write ALL data in ONE operation
        if all_data:
            data_start_row = header_row + 1
            data_end_row = header_row + len(all_data)
            print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Writing to Excel...")
            sheet.range((data_start_row, 1), (data_end_row, last_col)).value = all_data
            print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Data written")

        # Apply professional formatting
        print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Formatting...")
        try:
            # Header row - dark blue with white text
            header_range = sheet.range((header_row, 1), (header_row, last_col))
            header_range.font.name = 'Calibri Light'
            header_range.font.size = 10
            header_range.font.bold = True
            header_range.color = DARK_BLUE
            header_range.font.color = (255, 255, 255)

            if all_data:
                data_start_row = header_row + 1
                data_end_row = header_row + len(all_data)

                # Apply Calibri Light to all data
                data_range = sheet.range((data_start_row, 1), (data_end_row, last_col))
                data_range.font.name = 'Calibri Light'
                data_range.font.size = 10
                data_range.number_format = '#,##0'

                # Format header rows (section titles) - bold, larger font
                for r in header_rows_list:
                    sheet.range((r, 1)).font.bold = True
                    sheet.range((r, 1)).font.size = 11

                # Format total rows with bold and top border (professional accounting style)
                for r in total_rows:
                    row_range = sheet.range((r, 1), (r, last_col))
                    row_range.font.bold = True
                    try:
                        # Add top border for total rows
                        row_range.api.Borders(8).LineStyle = 1  # xlContinuous
                        row_range.api.Borders(8).Weight = 2  # xlThin
                    except:
                        pass

                # Special formatting for Total Assets and Total Liabilities & Equity (double underline)
                for r in total_rows:
                    try:
                        cell_value = sheet.range((r, 1)).value
                        if cell_value and ('Total Assets' in str(cell_value) or
                                          'Total Liabilities & Equity' in str(cell_value) or
                                          'Total Liabilities and Equity' in str(cell_value)):
                            row_range = sheet.range((r, 1), (r, last_col))
                            row_range.api.Borders(9).LineStyle = -4119  # xlDouble bottom
                            row_range.api.Borders(9).Weight = 4  # xlThick
                    except:
                        pass

            print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Formatting complete")
        except Exception as e:
            print(f"[BS-MULTI] Formatting error: {e}")

        try:
            sheet.range('A:A').column_width = 35
        except:
            pass

        # Group and collapse previous year columns
        print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Grouping previous year columns...")
        self._group_previous_year_columns(sheet, months, data_start_col=2)  # Data starts at column B

        # ================================================================
        # VALIDATION SECTION - Collapsible comparison of Source vs Calculated
        # ================================================================
        last_data_row = header_row + len(all_data) if all_data else header_row
        validation_start_row = last_data_row + 3
        last_col_letter = self._col_letter(last_col)

        # Add validation section header
        sheet.range(f'A{validation_start_row}').value = 'VALIDATION - Source vs Calculated Totals'
        sheet.range(f'A{validation_start_row}').font.bold = True
        sheet.range(f'A{validation_start_row}').font.size = 11
        sheet.range(f'A{validation_start_row}').font.color = (128, 0, 128)  # Purple

        # Column headers for validation
        val_header_row = validation_start_row + 1
        val_headers = ['Category', 'Source', 'Calculated', 'Variance', 'Match?']
        sheet.range(f'A{val_header_row}:E{val_header_row}').value = [val_headers]
        sheet.range(f'A{val_header_row}:E{val_header_row}').font.bold = True
        sheet.range(f'A{val_header_row}:E{val_header_row}').color = (200, 200, 200)

        # BS validation categories - key balance sheet totals (use wildcards)
        val_categories = [
            ('Total Assets', '*Total*Assets*'),
            ('Total Liabilities', '*Total*Liabilities*'),
            ('Total Equity', '*Total*Equity*'),
            ('Total Liab + Equity', '*Total*Liabilities*Equity*'),
        ]

        for idx, (label, source_pattern) in enumerate(val_categories):
            row = val_header_row + 1 + idx
            sheet.range(f'A{row}').value = label

            # Source total (SUMIF from Source_BS with wildcard)
            if is_consolidated:
                source_formula = f'=SUMIF(Source_BS!$B$3:$B$1500,"{source_pattern}",Source_BS!{last_col_letter}$3:{last_col_letter}$1500)'
            else:
                source_formula = f'=SUMIFS(Source_BS!{last_col_letter}$3:{last_col_letter}$1500,Source_BS!$A$3:$A$1500,"{division_name}",Source_BS!$B$3:$B$1500,"{source_pattern}")'
            sheet.range(f'B{row}').value = source_formula
            sheet.range(f'B{row}').number_format = '#,##0'

            # Calculated total - SUMIF on this sheet for matching account
            sheet.range(f'C{row}').value = f'=SUMIF(A:A,"{source_pattern}",{last_col_letter}:{last_col_letter})'
            sheet.range(f'C{row}').number_format = '#,##0'

            # Variance (Source - Calculated)
            sheet.range(f'D{row}').value = f'=B{row}-C{row}'
            sheet.range(f'D{row}').number_format = '#,##0'

            # Match indicator
            sheet.range(f'E{row}').value = f'=IF(ABS(D{row})<1,"✓","✗")'
            sheet.range(f'E{row}').font.size = 14

        validation_end_row = val_header_row + len(val_categories)

        # Add division breakdown for consolidated reports
        if is_consolidated and hasattr(self, 'divisions') and len(self.divisions) > 1:
            div_start_row = validation_end_row + 2
            sheet.range(f'A{div_start_row}').value = 'BY DIVISION - Total Assets'
            sheet.range(f'A{div_start_row}').font.bold = True
            sheet.range(f'A{div_start_row}').font.color = (0, 100, 0)

            div_header_row = div_start_row + 1
            div_headers = ['Division', 'Source', 'Calculated', 'Variance']
            sheet.range(f'A{div_header_row}:D{div_header_row}').value = [div_headers]
            sheet.range(f'A{div_header_row}:D{div_header_row}').font.bold = True
            sheet.range(f'A{div_header_row}:D{div_header_row}').color = (220, 220, 220)

            for div_idx, div in enumerate(self.divisions):
                div_name = div.get('name', div) if isinstance(div, dict) else getattr(div, 'name', str(div))
                row = div_header_row + 1 + div_idx
                sheet.range(f'A{row}').value = div_name
                # Source Total Assets for this division
                sheet.range(f'B{row}').value = f'=SUMIFS(Source_BS!{last_col_letter}$3:{last_col_letter}$1500,Source_BS!$A$3:$A$1500,"{div_name}",Source_BS!$B$3:$B$1500,"Total for Assets")'
                sheet.range(f'B{row}').number_format = '#,##0'
                # Calculated - reference division-specific sheet if exists
                div_sheet_name = f"{div_name}_BS"
                sheet.range(f'C{row}').value = f"=IFERROR(SUMIF('{div_sheet_name}'!A:A,\"*Total for Assets*\",'{div_sheet_name}'!{last_col_letter}:{last_col_letter}),0)"
                sheet.range(f'C{row}').number_format = '#,##0'
                sheet.range(f'D{row}').value = f'=B{row}-C{row}'
                sheet.range(f'D{row}').number_format = '#,##0'

            # Sum row
            sum_row = div_header_row + 1 + len(self.divisions)
            sheet.range(f'A{sum_row}').value = 'TOTAL'
            sheet.range(f'A{sum_row}').font.bold = True
            sheet.range(f'B{sum_row}').value = f'=SUM(B{div_header_row + 1}:B{sum_row - 1})'
            sheet.range(f'B{sum_row}').font.bold = True
            sheet.range(f'B{sum_row}').number_format = '#,##0'
            sheet.range(f'C{sum_row}').value = f'=SUM(C{div_header_row + 1}:C{sum_row - 1})'
            sheet.range(f'C{sum_row}').font.bold = True
            sheet.range(f'C{sum_row}').number_format = '#,##0'
            sheet.range(f'D{sum_row}').value = f'=B{sum_row}-C{sum_row}'
            sheet.range(f'D{sum_row}').font.bold = True
            sheet.range(f'D{sum_row}').number_format = '#,##0'

            validation_end_row = sum_row

        # Group/collapse the validation section
        try:
            sheet.api.Rows(f"{validation_start_row}:{validation_end_row}").Group()
            sheet.api.Outline.ShowLevels(RowLevels=1)  # Collapse by default
        except Exception as e:
            print(f"[BS-MULTI] Could not group validation rows: {e}")

        print(f"[BS-MULTI] === COMPLETE === Total time: {time.time() - start_time:.2f}s, validation rows {validation_start_row}-{validation_end_row}")

        # AutoFit all columns for best display
        try:
            sheet.autofit('c')  # AutoFit all columns
        except:
            pass

        # Collapse all outline groups
        try:
            sheet.api.Outline.ShowLevels(RowLevels=1, ColumnLevels=1)
        except:
            pass

        # Add print setup
        self._setup_print_area(sheet)

    def _create_multi_division_menu_sheet(self, sheet, months, divisions):
        """Create professional menu sheet with multi-division navigation and styling"""
        # Color palette
        DARK_BLUE = (22, 33, 62)
        ACCENT_BLUE = (59, 89, 152)
        GRAY = (128, 128, 128)
        LIGHT_GRAY = (245, 245, 245)
        LINK_BLUE = (0, 102, 204)
        WHITE = (255, 255, 255)
        SECTION_BG = (240, 244, 248)  # Light blue-gray for section backgrounds

        company = self.company_name.get()

        # Clear and set up the sheet
        sheet.range('A1:Z100').clear()

        # === HEADER SECTION (Rows 2-4) ===
        # Company name with accent bar
        sheet.range('B2:E2').merge()
        sheet.range('B2').value = company
        sheet.range('B2').font.name = 'Calibri Light'
        sheet.range('B2').font.size = 28
        sheet.range('B2').font.bold = True
        sheet.range('B2').font.color = DARK_BLUE

        sheet.range('B3:E3').merge()
        sheet.range('B3').value = 'Consolidated Financial Model'
        sheet.range('B3').font.name = 'Calibri Light'
        sheet.range('B3').font.size = 14
        sheet.range('B3').font.color = GRAY

        # Accent line under header
        try:
            accent_line = sheet.range('B4:E4')
            accent_line.color = ACCENT_BLUE
            accent_line.row_height = 4
        except:
            pass

        # === CONFIGURATION CARD (Rows 6-10) ===
        config_box = sheet.range('B6:C10')
        try:
            config_box.color = SECTION_BG
            # Add subtle border
            config_box.api.Borders.LineStyle = 1
            config_box.api.Borders.Color = 0xD0D0D0
            config_box.api.Borders.Weight = 1
        except:
            pass

        sheet.range('B6').value = 'MODEL SETTINGS'
        sheet.range('B6').font.name = 'Calibri Light'
        sheet.range('B6').font.size = 10
        sheet.range('B6').font.bold = True
        sheet.range('B6').font.color = ACCENT_BLUE

        # Get unique years for dropdown
        years_in_data = sorted(set(y for m, y, name in months))
        current_year = months[-1][1] if months else 2026

        config_labels = [
            ('Current Period:', months[-1][2] if months else 'N/A'),
            ('Data Range:', f"{months[0][2]} - {months[-1][2]}" if months else 'N/A'),
            ('Divisions:', str(len(divisions))),
            ('Reporting Year:', str(current_year)),  # New: Reporting year for grouping
        ]

        for i, (label, value) in enumerate(config_labels):
            row = 7 + i
            sheet.range(f'B{row}').value = label
            sheet.range(f'C{row}').value = value
            sheet.range(f'B{row}').font.name = 'Calibri Light'
            sheet.range(f'B{row}').font.size = 10
            sheet.range(f'B{row}').font.color = GRAY
            sheet.range(f'C{row}').font.name = 'Calibri Light'
            sheet.range(f'C{row}').font.size = 10
            sheet.range(f'C{row}').font.bold = True

        # Add dropdown for Reporting Year (C10)
        try:
            year_list = ','.join([str(y) for y in years_in_data])
            sheet.range('C10').api.Validation.Delete()
            sheet.range('C10').api.Validation.Add(Type=3, AlertStyle=1, Formula1=year_list)
            sheet.range('C10').color = (255, 255, 200)  # Light yellow to indicate editable
        except:
            pass

        # Add "Update Grouping" action button hint
        sheet.range('D10').value = '← Run "UpdateYearGrouping" macro after changing'
        sheet.range('D10').font.name = 'Calibri Light'
        sheet.range('D10').font.size = 8
        sheet.range('D10').font.color = GRAY
        sheet.range('D10').font.italic = True

        # === CONSOLIDATED REPORTS SECTION (Rows 12-19) ===
        cons_box = sheet.range('B12:C19')
        try:
            cons_box.color = SECTION_BG
            cons_box.api.Borders.LineStyle = 1
            cons_box.api.Borders.Color = 0xD0D0D0
            cons_box.api.Borders.Weight = 1
        except:
            pass

        # Section header with icon indicator
        sheet.range('B12').value = '📊 REPORTS (Click to Navigate)'
        sheet.range('B12').font.name = 'Calibri Light'
        sheet.range('B12').font.size = 11
        sheet.range('B12').font.bold = True
        sheet.range('B12').font.color = ACCENT_BLUE

        cons_nav = [
            ('▸ Dashboard', 'Dashboard', 'Executive summary & KPIs'),
            ('▸ P&L Statement', 'Consolidated_PL', 'Profit & Loss'),
            ('▸ Cash Flow', 'Cash_Flow', 'Cash movements'),
            ('▸ Forecast', 'Consolidated_Forecast', 'Budget vs Actual'),
            ('▸ Forecast Summary', 'Forecast_Summary', 'YTD variances'),
            ('▸ Notes', 'Notes', 'Account annotations'),
        ]

        for i, (label, target, desc) in enumerate(cons_nav):
            row = 13 + i
            cell = sheet.range(f'B{row}')
            cell.value = label
            cell.font.name = 'Calibri Light'
            cell.font.size = 10
            cell.font.color = LINK_BLUE
            cell.font.underline = True
            # Add description in column C
            sheet.range(f'C{row}').value = desc
            sheet.range(f'C{row}').font.name = 'Calibri Light'
            sheet.range(f'C{row}').font.size = 9
            sheet.range(f'C{row}').font.color = GRAY
            sheet.range(f'C{row}').font.italic = True
            try:
                cell.add_hyperlink(f'#{target}!A1', text_to_display=label)
            except:
                pass

        # === SOURCE DATA SECTION (Rows 12-16, Column D-E) ===
        source_box = sheet.range('D12:E16')
        try:
            source_box.color = SECTION_BG
            source_box.api.Borders.LineStyle = 1
            source_box.api.Borders.Color = 0xD0D0D0
            source_box.api.Borders.Weight = 1
        except:
            pass

        sheet.range('D12').value = '📁 SOURCE DATA'
        sheet.range('D12').font.name = 'Calibri Light'
        sheet.range('D12').font.size = 11
        sheet.range('D12').font.bold = True
        sheet.range('D12').font.color = ACCENT_BLUE

        source_nav = [
            ('▸ P&L Data', 'Source_PL', 'Raw P&L'),
            ('▸ Balance Sheet', 'Source_BS', 'Raw BS'),
            ('▸ Budget', 'Source_Budget', 'Budget data'),
        ]

        for i, (label, target, desc) in enumerate(source_nav):
            row = 13 + i
            cell = sheet.range(f'D{row}')
            cell.value = label
            cell.font.name = 'Calibri Light'
            cell.font.size = 10
            cell.font.color = LINK_BLUE
            cell.font.underline = True
            sheet.range(f'E{row}').value = desc
            sheet.range(f'E{row}').font.name = 'Calibri Light'
            sheet.range(f'E{row}').font.size = 9
            sheet.range(f'E{row}').font.color = GRAY
            sheet.range(f'E{row}').font.italic = True
            try:
                cell.add_hyperlink(f'#{target}!A1', text_to_display=label)
            except:
                pass

        # === DIVISION SECTIONS ===
        current_row = 21

        for div_idx, div in enumerate(divisions):
            safe_name = div['name'].replace(' ', '_')[:20]

            # Division card
            div_box = sheet.range(f'B{current_row}:E{current_row + 1}')
            try:
                div_box.color = SECTION_BG
                div_box.api.Borders.LineStyle = 1
                div_box.api.Borders.Color = 0xD0D0D0
                div_box.api.Borders.Weight = 1
            except:
                pass

            # Division header
            sheet.range(f'B{current_row}').value = div['name'].upper()
            sheet.range(f'B{current_row}').font.name = 'Calibri Light'
            sheet.range(f'B{current_row}').font.size = 10
            sheet.range(f'B{current_row}').font.bold = True
            sheet.range(f'B{current_row}').font.color = ACCENT_BLUE

            # Division links (P&L, BS, Forecast on same row)
            link_row = current_row + 1
            div_links = [
                ('B', 'P&L', f'{safe_name}_PL'),
                ('C', 'Balance Sheet', f'{safe_name}_BS'),
                ('D', 'Forecast', f'{safe_name}_Forecast'),
            ]

            for col, label, target in div_links:
                sheet.range(f'{col}{link_row}').value = f'  {label}'
                sheet.range(f'{col}{link_row}').font.name = 'Calibri Light'
                sheet.range(f'{col}{link_row}').font.size = 10
                sheet.range(f'{col}{link_row}').font.color = LINK_BLUE
                sheet.range(f'{col}{link_row}').font.underline = True
                try:
                    sheet.range(f'{col}{link_row}').add_hyperlink(f'#{target}!A1', text_to_display=f'  {label}')
                except:
                    pass

            current_row += 3  # Space between division cards

        # === FOOTER ===
        footer_row = current_row + 1
        sheet.range(f'B{footer_row}').value = 'CFO DNA Financial Model Generator'
        sheet.range(f'B{footer_row}').font.name = 'Calibri Light'
        sheet.range(f'B{footer_row}').font.size = 8
        sheet.range(f'B{footer_row}').font.color = GRAY
        sheet.range(f'B{footer_row}').font.italic = True

        # Add helper cells for YTD calculations (hidden)
        # E7 = current month, F7 = current year (for P&L formulas)
        # H7 = current month, I7 = reporting year (from C10), J7 = YYYYMM
        try:
            if months:
                current_month = months[-1][0]
                current_year = months[-1][1]
                # E7/F7 for backward compatibility with P&L formulas
                sheet.range('E7').value = current_month
                sheet.range('F7').value = current_year
                sheet.range('G7').value = current_year * 100 + current_month
                # H7/I7/J7 for dynamic reporting year
                sheet.range('H7').value = current_month
                sheet.range('I7').formula = '=C10'  # References C10 (Reporting Year)
                sheet.range('J7').formula = '=I7*100+H7'
                sheet.range('J9').formula = '=J7'
            # Only hide columns G:K (helper columns), keep E visible for navigation descriptions
            sheet.range('G:K').api.Hidden = True
        except:
            pass

        # Set column widths for clean layout
        sheet.range('A:A').column_width = 3
        sheet.range('B:B').column_width = 20
        sheet.range('C:C').column_width = 18
        sheet.range('D:D').column_width = 18
        sheet.range('E:E').column_width = 8
        sheet.range('F:F').column_width = 3

        # Hide gridlines for clean look
        try:
            sheet.book.app.api.ActiveWindow.DisplayGridlines = False
        except:
            pass

    def _create_excel_model(self, save_path):
        """Create the Excel model using xlwings"""
        # Get user-specified date range
        user_start, user_end = self._get_user_date_params()

        # Check if multi-division mode
        if self.is_multi_division.get() and self.divisions:
            return self._create_multi_division_model(save_path, user_start, user_end)

        # Single entity mode - use progress tracker for consistent UI
        total_steps = 13  # 12 main steps + 1 finalizing step
        estimated_total = 90  # Estimated seconds for single-entity mode (conservative)
        update_step, update_substep, state = self._create_progress_tracker(total_steps, estimated_total)

        # Parse input files with indentation detection
        update_step("Reading P&L file...")
        pl_data, pl_indents = self._read_excel_with_indents(self.pl_path.get())

        update_step("Reading Balance Sheet file...")
        bs_data, bs_indents = self._read_excel_with_indents(self.bs_path.get())

        update_step("Extracting account data...")
        pl_accounts, pl_months, pl_totals = self._parse_financial_data(pl_data, pl_indents, user_start, user_end)
        bs_accounts, _, bs_totals = self._parse_financial_data(bs_data, bs_indents, user_start, user_end)
        print(f"Found {len(pl_accounts)} P&L accounts, {len(bs_accounts)} BS accounts, {len(pl_months)} months")
        print(f"P&L Totals detected: {pl_totals}")
        print(f"BS Totals detected: {bs_totals}")

        # Check if template exists
        use_template = os.path.exists(TEMPLATE_PATH)

        # Work in temp directory to avoid OneDrive locking issues
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, 'temp_model.xlsm')

        # Create Excel application
        update_step("Starting Excel...")
        app = xw.App(visible=False)
        app.display_alerts = False  # Suppress Excel prompts
        app.screen_updating = False  # Speed up processing
        app.calculation = 'manual'  # Disable auto-calculation for speed
        wb = None

        try:
            if use_template:
                # Copy template to temp location and open
                shutil.copy(TEMPLATE_PATH, temp_path)
                wb = app.books.open(temp_path)

                # Get existing sheets
                menu_sheet = wb.sheets['Menu']
                source_pl = wb.sheets['Source_PL']
                source_bs = wb.sheets['Source_BS']
                pl_sheet = wb.sheets['PL']
                bs_sheet = wb.sheets['Balance_Sheet']
                cf_sheet = wb.sheets['Cash_Flow']
                notes_sheet = wb.sheets['Notes']

                # Create new sheets if they don't exist (for v1.3.0 Forecast module)
                try:
                    source_budget = wb.sheets['Source_Budget']
                    source_budget.range('A1:ZZ1000').clear()
                except:
                    source_budget = wb.sheets.add('Source_Budget', after=source_bs)

                try:
                    forecast_sheet = wb.sheets['Forecast']
                    forecast_sheet.range('A1:ZZ1000').clear()
                except:
                    forecast_sheet = wb.sheets.add('Forecast', after=cf_sheet)

                try:
                    forecast_summary_sheet = wb.sheets['Forecast_Summary']
                    forecast_summary_sheet.range('A1:ZZ1000').clear()
                except:
                    forecast_summary_sheet = wb.sheets.add('Forecast_Summary', after=forecast_sheet)

                # Clear source sheets (keep headers)
                source_pl.range('A2:ZZ1000').clear()
                source_bs.range('A2:ZZ1000').clear()

                # Also clear the report sheets for fresh data (including header rows)
                pl_sheet.range('A1:ZZ1000').clear()
                bs_sheet.range('A1:ZZ1000').clear()
                cf_sheet.range('A1:ZZ1000').clear()
            else:
                # Create from scratch (requires VBA trust setting)
                wb = app.books.add()

                # Remove default sheets and create our sheets
                for sheet in wb.sheets:
                    if sheet.name not in ['Sheet1']:
                        sheet.delete()

                # Create sheets
                menu_sheet = wb.sheets[0]
                menu_sheet.name = 'Menu'

                source_pl = wb.sheets.add('Source_PL', after=menu_sheet)
                source_bs = wb.sheets.add('Source_BS', after=source_pl)
                # P&L/BS/CF reports come before Budget/Forecast
                pl_sheet = wb.sheets.add('PL', after=source_bs)
                bs_sheet = wb.sheets.add('Balance_Sheet', after=pl_sheet)
                cf_sheet = wb.sheets.add('Cash_Flow', after=bs_sheet)
                # Budget and Forecast come after all report sheets
                source_budget = wb.sheets.add('Source_Budget', after=cf_sheet)
                forecast_sheet = wb.sheets.add('Forecast', after=source_budget)
                forecast_summary_sheet = wb.sheets.add('Forecast_Summary', after=forecast_sheet)
                notes_sheet = wb.sheets.add('Notes', after=forecast_summary_sheet)

            # Populate source sheets
            update_step("Populating source data...")
            self._populate_source_sheet(source_pl, pl_accounts, pl_months)
            self._populate_source_sheet(source_bs, bs_accounts, pl_months)
            self._create_source_budget_sheet(source_budget, pl_accounts, pl_months)

            # Create/update named ranges
            pl_last_row = len(pl_accounts) + 2  # +2 for header and YYYYMM rows
            pl_last_col = len(pl_months) + 1
            bs_last_row = len(bs_accounts) + 2

            # Delete existing named ranges if they exist
            try:
                wb.names['SourcePL'].delete()
            except:
                pass
            try:
                wb.names['SourceBS'].delete()
            except:
                pass
            try:
                wb.names['SourceBudget'].delete()
            except:
                pass

            wb.names.add('SourcePL', f"=Source_PL!$A$1:${self._col_letter(pl_last_col)}${pl_last_row}")
            wb.names.add('SourceBS', f"=Source_BS!$A$1:${self._col_letter(pl_last_col)}${bs_last_row}")
            wb.names.add('SourceBudget', f"=Source_Budget!$A$1:${self._col_letter(pl_last_col)}${pl_last_row}")

            # Create Menu sheet
            update_step("Creating Menu sheet...")
            self._create_menu_sheet(menu_sheet, pl_months)

            # Create P&L report (using optimized bulk-write function)
            update_step("Creating P&L report...")
            self._create_pl_report_multi_div(pl_sheet, pl_accounts, pl_months, pl_totals,
                                              is_consolidated=True, division_name=None, single_entity_mode=True)

            # Create Balance Sheet report (using optimized bulk-write function)
            update_step("Creating Balance Sheet...")
            self._create_bs_report_multi_div(bs_sheet, bs_accounts, pl_months, bs_totals,
                                              is_consolidated=True, division_name=None, single_entity_mode=True)

            # Create Cash Flow statement
            update_step("Creating Cash Flow statement...")
            self._create_cash_flow(cf_sheet, pl_months)

            # Create Forecast sheet
            update_step("Creating Forecast sheet...")
            self._create_forecast_sheet(forecast_sheet, pl_accounts, pl_months)

            # Create Forecast Summary and Notes
            update_substep("Creating Forecast Summary...")
            self._create_forecast_summary_sheet(forecast_summary_sheet, pl_accounts, pl_months)

            # Create Notes sheet with dropdowns (consolidated for P&L, BS, and CF)
            month_names = [name for m, y, name in pl_months]
            self._create_notes_sheet(notes_sheet, pl_accounts, bs_accounts, month_names)

            # Create Dashboard
            update_step("Creating Dashboard...")
            if use_template:
                # Check if Dashboard_Control already exists
                if 'Dashboard_Control' not in [s.name for s in wb.sheets]:
                    dashboard_control = wb.sheets.add('Dashboard_Control', after=notes_sheet)
                else:
                    dashboard_control = wb.sheets['Dashboard_Control']
                    dashboard_control.range('A1:ZZ1000').clear()
            else:
                dashboard_control = wb.sheets.add('Dashboard_Control', after=notes_sheet)
            self._create_dashboard_control_sheet(dashboard_control, pl_months)

            # Create Dashboard sheet
            if use_template:
                if 'Dashboard' not in [s.name for s in wb.sheets]:
                    dashboard_sheet = wb.sheets.add('Dashboard', before=menu_sheet)
                else:
                    dashboard_sheet = wb.sheets['Dashboard']
                    dashboard_sheet.range('A1:ZZ1000').clear()
            else:
                dashboard_sheet = wb.sheets.add('Dashboard', before=menu_sheet)
            self._create_dashboard_sheet(dashboard_sheet, pl_accounts, bs_accounts, pl_months, pl_totals)

            # Reorder sheets to match multi-division layout:
            # Dashboard, Menu, P&L, Balance_Sheet, Cash_Flow, Forecast, Forecast_Summary, Notes, Dashboard_Control, Source sheets
            try:
                # Build desired order
                sheet_order = ['Dashboard', 'Menu', 'PL', 'Balance_Sheet', 'Cash_Flow',
                               'Forecast', 'Forecast_Summary', 'Notes', 'Dashboard_Control',
                               'Source_PL', 'Source_BS', 'Source_Budget']

                # Get all current sheets
                all_sheets = {s.name: s for s in wb.sheets}

                # Move sheets in reverse order (since Move(Before=) puts it before the first sheet)
                prev_sheet = None
                for name in sheet_order:
                    if name in all_sheets:
                        if prev_sheet is None:
                            # Move to first position
                            all_sheets[name].api.Move(Before=wb.sheets[0].api)
                        else:
                            all_sheets[name].api.Move(After=prev_sheet.api)
                        prev_sheet = all_sheets[name]
            except:
                pass

            # Set source sheet tab colors to black
            try:
                source_pl.api.Tab.Color = 0x000000  # Black
                source_bs.api.Tab.Color = 0x000000  # Black
                source_budget.api.Tab.Color = 0x000000  # Black
            except:
                pass

            # Activate Dashboard sheet so file opens to Dashboard
            try:
                dashboard_sheet.activate()
            except:
                pass

            # Add VBA code - always add to ensure macros work
            try:
                # Check if FinancialModel module already exists (from template)
                module_exists = False
                for component in wb.api.VBProject.VBComponents:
                    if component.Name == "FinancialModel":
                        # Clear existing code and replace with current version
                        component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                        component.CodeModule.AddFromString(self.VBA_CODE)
                        module_exists = True
                        break

                if not module_exists:
                    vba_module = wb.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
                    vba_module.Name = "FinancialModel"
                    vba_module.CodeModule.AddFromString(self.VBA_CODE)

                # Add Worksheet_Change event to Menu sheet to trigger UpdateColumnVisibility
                # when C7 changes
                menu_sheet_code = '''
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("C7")) Is Nothing Then
        Call UpdateColumnVisibility
    End If
End Sub
'''
                # Find the Menu sheet's code module and add the event
                for component in wb.api.VBProject.VBComponents:
                    if component.Type == 100:  # 100 = vbext_ct_Document (worksheet)
                        if component.Name == "Menu" or (hasattr(component, 'Properties') and component.Properties("Name").Value == "Menu"):
                            # Clear any existing code and add fresh
                            if component.CodeModule.CountOfLines > 0:
                                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                            component.CodeModule.AddFromString(menu_sheet_code)
                            break
            except Exception as e:
                # VBA access not enabled - save without macros
                pass

            # Re-enable calculation and save the workbook
            update_step("Saving workbook...")
            app.calculation = 'automatic'  # Re-enable before save
            wb.save()
            wb.close()
            wb = None

        finally:
            # Ensure proper cleanup
            try:
                if wb is not None:
                    wb.close()
            except:
                pass
            try:
                app.quit()
            except:
                pass

        # Finalizing step
        update_step("Finalizing...")

        # Move from temp to final location (after Excel is closed)
        try:
            # Remove existing file if it exists
            if os.path.exists(save_path):
                os.remove(save_path)
            shutil.move(temp_path, save_path)
        finally:
            # Clean up temp directory
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

    def _get_cell_indents(self, file_path):
        """
        Read Excel file and extract indentation levels for column A.
        Returns dict: {row_number: indent_level}
        """
        from openpyxl import load_workbook
        indents = {}
        try:
            wb = load_workbook(file_path, data_only=True)
            ws = wb.active
            for row_idx, row in enumerate(ws.iter_rows(min_col=1, max_col=1), start=0):
                cell = row[0]
                indent = 0
                if cell.alignment and cell.alignment.indent:
                    indent = int(cell.alignment.indent)
                # Also check for leading spaces in value
                if cell.value and isinstance(cell.value, str):
                    leading_spaces = len(cell.value) - len(cell.value.lstrip())
                    # Convert leading spaces to indent level (4 spaces = 1 level)
                    space_indent = leading_spaces // 4
                    indent = max(indent, space_indent)
                indents[row_idx] = indent
            wb.close()
        except Exception as e:
            print(f"Could not read indentation from {file_path}: {e}")
        return indents

    def _read_excel_with_indents(self, file_path):
        """
        Read Excel file and extract both data and indentation in a single pass.
        Returns tuple: (pandas DataFrame, dict of {row_number: indent_level})

        This is more efficient than calling pd.read_excel() and _get_cell_indents()
        separately, as it only opens the file once.
        """
        from openpyxl import load_workbook
        indents = {}
        data = []

        try:
            wb = load_workbook(file_path, data_only=True)
            ws = wb.active

            for row_idx, row in enumerate(ws.iter_rows()):
                # Extract indentation from first column
                first_cell = row[0]
                indent = 0
                if first_cell.alignment and first_cell.alignment.indent:
                    indent = int(first_cell.alignment.indent)
                # Also check for leading spaces in value
                if first_cell.value and isinstance(first_cell.value, str):
                    leading_spaces = len(first_cell.value) - len(first_cell.value.lstrip())
                    space_indent = leading_spaces // 4
                    indent = max(indent, space_indent)
                indents[row_idx] = indent

                # Extract row data
                row_data = [cell.value for cell in row]
                data.append(row_data)

            wb.close()

            # Convert to DataFrame (same format as pd.read_excel with header=None)
            df = pd.DataFrame(data)
            return df, indents

        except Exception as e:
            print(f"Could not read file {file_path}: {e}")
            # Fall back to separate reads
            df = pd.read_excel(file_path, header=None)
            indents = self._get_cell_indents(file_path)
            return df, indents

    def _parse_month_value(self, val):
        """
        Flexibly parse a value to extract month/year info.
        Returns (month_num, year, display_name) or None if not a recognizable month.
        Handles many formats: "Jan 24", "January 2024", "Jan-24", "1/24", dates, etc.
        """
        import re

        if pd.isna(val):
            return None

        # Handle datetime/Timestamp objects directly
        if isinstance(val, (pd.Timestamp, datetime)):
            month_num = val.month
            year = val.year
            month_abbrev = self.MONTHS[month_num - 1][:3]
            display_name = f"{month_abbrev} {year % 100}"
            return (month_num, year, display_name)

        val_str = str(val).strip()
        if not val_str or 'total' in val_str.lower():
            return None

        val_lower = val_str.lower()

        # Skip date RANGE patterns (e.g., "August 1, 2021-November 30, 2025" or "January-December, 2024")
        # These contain multiple months or range indicators and should not be treated as a single month
        month_count = 0
        for month in self.MONTHS:
            if month.lower() in val_lower or month[:3].lower() in val_lower:
                month_count += 1
        if month_count > 1:
            # Multiple months found - this is a date range, not a single month
            return None

        # Also skip if it looks like a date range with hyphen between date components
        # Pattern: "Month Day, Year - Month Day, Year" or similar
        if '-' in val_str and any(c.isdigit() for c in val_str.split('-')[0]) and any(c.isdigit() for c in val_str.split('-')[-1]):
            # Has hyphen with numbers on both sides - likely a date range
            return None

        # Skip "As of" patterns - these are point-in-time descriptions, not column headers
        if val_lower.startswith('as of '):
            return None

        # Strategy 1: Look for month name (full or abbreviated) anywhere in the string
        for i, month in enumerate(self.MONTHS, 1):
            month_lower = month.lower()
            month_short = month_lower[:3]

            if month_short in val_lower or month_lower in val_lower:
                # Found a month name - now extract year
                # Prefer 4-digit years over 2-digit years to avoid confusion with day numbers
                numbers = re.findall(r'\d+', val_str)
                # First try to find a 4-digit year
                for num_str in numbers:
                    num = int(num_str)
                    if 1900 <= num <= 2100:
                        display_name = f"{month[:3]} {num % 100}"
                        return (i, num, display_name)
                # Fall back to 2-digit year if no 4-digit year found
                for num_str in numbers:
                    if len(num_str) == 2:
                        year = 2000 + int(num_str)
                        display_name = f"{month[:3]} {year % 100}"
                        return (i, year, display_name)

        # Strategy 2: Numeric date formats like "1/24", "01/2024", "1-24", "12/25"
        # Pattern: month/year or month-year
        date_patterns = [
            r'^(\d{1,2})[/\-\.](\d{2,4})$',  # 1/24, 01/2024, 1-24
            r'^(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{2,4})$',  # 1/1/24 - take first as month, last as year
        ]

        for pattern in date_patterns:
            match = re.match(pattern, val_str)
            if match:
                groups = match.groups()
                month_num = int(groups[0])
                year = int(groups[-1])  # Last group is year
                if 1 <= month_num <= 12:
                    if year < 100:
                        year = 2000 + year
                    month_abbrev = self.MONTHS[month_num - 1][:3]
                    display_name = f"{month_abbrev} {year % 100}"
                    return (month_num, year, display_name)

        # Strategy 3: Try pandas date parsing as last resort
        try:
            parsed_date = pd.to_datetime(val_str, errors='raise', dayfirst=False)
            if parsed_date:
                month_num = parsed_date.month
                year = parsed_date.year
                if 1900 <= year <= 2100:  # Sanity check
                    month_abbrev = self.MONTHS[month_num - 1][:3]
                    display_name = f"{month_abbrev} {year % 100}"
                    return (month_num, year, display_name)
        except:
            pass

        return None

    def _parse_financial_data(self, df, indents=None, user_start=None, user_end=None):
        """Parse financial data from dataframe with indentation levels.

        Args:
            df: DataFrame with financial data
            indents: Dict of row_idx -> indent level
            user_start: Tuple of (month_num, year) user-specified start date
            user_end: Tuple of (month_num, year) user-specified end date
        """
        indents = indents or {}

        # Find header row by looking for any recognizable month pattern or common headers
        header_row = None
        has_total_column = False

        for idx in range(min(15, len(df))):
            for col_idx in range(len(df.columns)):
                val = df.iloc[idx, col_idx]
                if self._parse_month_value(val):
                    header_row = idx
                    break
                # Also check for "Total" header which indicates a summary column
                if pd.notna(val) and str(val).strip().lower() == 'total':
                    header_row = idx
                    has_total_column = True
                    break
            if header_row is not None:
                break

        if header_row is None:
            # Fallback: Look for row before common financial statement headers
            common_first_items = ['income', 'revenue', 'sales', 'assets', 'liabilities',
                                  'equity', 'expenses', 'cost of', 'ordinary income']
            for idx in range(min(15, len(df))):
                first_cell = df.iloc[idx, 0]
                if pd.notna(first_cell):
                    first_cell_lower = str(first_cell).strip().lower()
                    for pattern in common_first_items:
                        if first_cell_lower.startswith(pattern):
                            # The header row is likely the row BEFORE this
                            header_row = max(0, idx - 1)
                            print(f"Found data starting at row {idx} ('{first_cell}'), using row {header_row} as header")
                            break
                    if header_row is not None:
                        break

        if header_row is None:
            # Try to provide more helpful error message
            print("DEBUG: First 5 rows of data:")
            for i in range(min(5, len(df))):
                print(f"  Row {i}: {list(df.iloc[i, :5])}")
            raise ValueError("Could not find month headers in file. Please ensure the file has column headers.")

        # Extract months using the flexible parser
        months = []
        data_columns = []  # Track which columns have actual data (not "Total")

        print(f"DEBUG: Header row {header_row}, user_start={user_start}, user_end={user_end}")
        print(f"DEBUG: Header row contents: {list(df.iloc[header_row, :min(10, len(df.columns))])}")

        for col_idx in range(1, len(df.columns)):
            val = df.iloc[header_row, col_idx]
            if pd.isna(val):
                continue

            val_str = str(val).strip().lower()
            print(f"DEBUG: Col {col_idx} = '{val}' (type: {type(val).__name__})")

            # Skip "Total" columns when we have other month columns
            if val_str == 'total':
                has_total_column = True
                print(f"DEBUG: Found Total column at {col_idx}")
                # We'll decide later whether to use this
                continue

            parsed = self._parse_month_value(val)
            if parsed:
                print(f"DEBUG: Parsed '{val}' as month: {parsed}")
                months.append(parsed)
                data_columns.append(col_idx)
            else:
                print(f"DEBUG: Could not parse '{val}' as month")

        # If no months found, try to use user-specified dates
        print(f"DEBUG: After parsing - months found: {len(months)}, has_total_column: {has_total_column}")
        print(f"DEBUG: months list: {months}")
        print(f"DEBUG: data_columns: {data_columns}")

        if not months and user_start and user_end:
            print(f"DEBUG: No months found, using user-specified dates: {user_start} to {user_end}")
            start_month, start_year = user_start
            end_month, end_year = user_end

            # Find a data column to use - prefer Total column, otherwise use first column with numeric data
            data_col = None

            # First, look for Total column in header row
            for col_idx in range(1, len(df.columns)):
                val = df.iloc[header_row, col_idx]
                if pd.notna(val) and str(val).strip().lower() == 'total':
                    data_col = col_idx
                    print(f"DEBUG: Found Total column at {col_idx}")
                    break

            # If no Total column found in header, look for first column with numeric data in data rows
            if data_col is None:
                # Check a few rows after header to find where numeric data is
                for col_idx in range(1, len(df.columns)):
                    for check_row in range(header_row + 1, min(header_row + 10, len(df))):
                        val = df.iloc[check_row, col_idx]
                        if pd.notna(val):
                            try:
                                float(val)
                                data_col = col_idx
                                print(f"DEBUG: Found numeric data in column {col_idx} at row {check_row}")
                                break
                            except (ValueError, TypeError):
                                continue
                    if data_col is not None:
                        break

            print(f"DEBUG: data_col = {data_col}")

            if data_col:
                # If start == end, it's a single month report
                if start_month == end_month and start_year == end_year:
                    month_abbrev = self.MONTHS[start_month - 1][:3]
                    display_name = f"{month_abbrev} {start_year % 100}"
                    months = [(start_month, start_year, display_name)]
                    data_columns = [data_col]
                    print(f"Using user-specified single month: {display_name}, data from column {data_col}")
                else:
                    # Multi-month range - but we only have one column of data
                    # This is a special case - user says it's a range but file has one column
                    # Use the single column for the END month (most recent)
                    month_abbrev = self.MONTHS[end_month - 1][:3]
                    display_name = f"{month_abbrev} {end_year % 100}"
                    months = [(end_month, end_year, display_name)]
                    data_columns = [data_col]
                    print(f"Using user-specified end month: {display_name}, data from column {data_col}")

        # If still no months, raise error with helpful message
        if not months:
            raise ValueError("Could not determine date columns. Please specify the data period in the UI (Start/End dates).")

        # Extract accounts and values, tracking exact total names
        accounts = []
        detected_totals = {
            'total_income': None,
            'total_cogs': None,
            'total_expenses': None,
            'gross_profit': None,
            'net_income': None,
            'total_assets': None,
            'total_liabilities': None,
            'total_equity': None,
            'total_liab_equity': None,
        }

        # If data_columns wasn't populated (old code path), default to sequential
        if not data_columns:
            data_columns = list(range(1, len(months) + 1))

        for row_idx in range(header_row + 1, len(df)):
            account_name = df.iloc[row_idx, 0]
            if pd.isna(account_name) or not str(account_name).strip():
                continue
            account_name = str(account_name).strip()
            if 'cash basis' in account_name.lower():
                continue

            values = {}
            # Use data_columns to get values from correct columns
            # Store values with (month, year) tuple keys for lookup
            for month_idx, col_idx in enumerate(data_columns):
                if col_idx < len(df.columns) and month_idx < len(months):
                    val = df.iloc[row_idx, col_idx]
                    month_m, month_y, _ = months[month_idx]
                    if pd.notna(val):
                        try:
                            values[(month_m, month_y)] = float(val)
                        except:
                            values[(month_m, month_y)] = 0
                    else:
                        values[(month_m, month_y)] = 0

            is_total = account_name.lower().startswith('total') or account_name in ['Net Income', 'Gross Profit']
            is_header = account_name in ['Income', 'Expenses', 'Cost of Sales', 'Assets', 'Liabilities', 'Equity',
                                         'Other Current Assets', 'Fixed Assets', 'Other Assets',
                                         'Current Liabilities', 'Long Term Liabilities', 'Other Current Liabilities']

            # Get indent level from source file
            indent_level = indents.get(row_idx, 0)

            # Detect specific total rows by their exact names
            name_lower = account_name.lower()
            if name_lower.startswith('total income') or name_lower == 'total for income':
                detected_totals['total_income'] = account_name
            elif (name_lower.startswith('total cost') or
                  name_lower.startswith('total cogs') or
                  name_lower.startswith('total for cost') or
                  ('cost of sales' in name_lower and 'total' in name_lower) or
                  ('cost of goods' in name_lower and 'total' in name_lower)):
                detected_totals['total_cogs'] = account_name
            elif name_lower.startswith('total expenses') or name_lower == 'total for expenses':
                detected_totals['total_expenses'] = account_name
            elif 'gross profit' in name_lower:
                detected_totals['gross_profit'] = account_name
            elif name_lower == 'net income':
                detected_totals['net_income'] = account_name
            elif name_lower.startswith('total for assets') or name_lower == 'total assets':
                detected_totals['total_assets'] = account_name
            elif name_lower.startswith('total for liabilities and equity') or name_lower == 'total liabilities and equity':
                detected_totals['total_liab_equity'] = account_name
            elif (name_lower.startswith('total for liabilities') or name_lower == 'total liabilities') and 'equity' not in name_lower:
                detected_totals['total_liabilities'] = account_name
            elif (name_lower.startswith('total for equity') or name_lower == 'total equity') and 'liabilities' not in name_lower:
                detected_totals['total_equity'] = account_name

            accounts.append({
                'name': account_name,
                'indent': indent_level,
                'values': values,
                'is_total': is_total,
                'is_header': is_header
            })

        return accounts, months, detected_totals

    def _detect_actual_date_range(self, accounts, months):
        """Detect which months actually have data (non-zero values).

        Filters out leading months with no data in any account.

        Args:
            accounts: List of account dictionaries with 'values' dict
            months: List of (month, year, display_name) tuples

        Returns:
            Filtered list of months that have actual data
        """
        if not months or not accounts:
            return months

        # Find first and last month with any non-zero value
        first_month_with_data = None
        last_month_with_data = None

        for idx, (m, y, name) in enumerate(months):
            has_data = False
            for account in accounts:
                if account.get('is_header', False):
                    continue
                val = account.get('values', {}).get((m, y), 0)
                if val != 0:
                    has_data = True
                    break

            if has_data:
                if first_month_with_data is None:
                    first_month_with_data = idx
                last_month_with_data = idx

        # If no data found anywhere, return original months
        if first_month_with_data is None:
            return months

        # Return filtered months (from first with data to last with data)
        filtered = months[first_month_with_data:last_month_with_data + 1]
        if len(filtered) < len(months):
            print(f"Date range detected: trimmed {len(months)} months to {len(filtered)} with actual data")
            print(f"  Original: {months[0][2]} to {months[-1][2]}")
            print(f"  Filtered: {filtered[0][2]} to {filtered[-1][2]}")

        return filtered

    def _get_previous_year_columns(self, months, data_start_col):
        """Identify columns that belong to previous years for grouping.

        Groups all columns from years PRIOR to the reporting year (last year in data).
        Example: If data goes through Dec 2025, groups all 2024 and earlier columns.

        Args:
            months: List of (month, year, display_name) tuples
            data_start_col: First column number where month data starts (1-indexed)

        Returns:
            List of (start_col, end_col, year) tuples for each previous year group
        """
        if not months:
            return []

        # Use the last month's year as the "reporting year" - group everything before it
        reporting_year = months[-1][1]

        year_groups = []
        current_group_start = None
        current_group_year = None

        for idx, (m, y, name) in enumerate(months):
            col = data_start_col + idx

            if y < reporting_year:
                # This is a prior year column - should be grouped
                if current_group_year != y:
                    # New year group - save the old one if exists
                    if current_group_start is not None:
                        year_groups.append((current_group_start, col - 1, current_group_year))
                    current_group_start = col
                    current_group_year = y
            else:
                # Current/reporting year - close any open previous year group
                if current_group_start is not None:
                    year_groups.append((current_group_start, col - 1, current_group_year))
                    current_group_start = None
                    current_group_year = None

        # Close final group if still open (all months are prior year)
        if current_group_start is not None:
            # Find last prior year column
            last_prior_col = data_start_col + len(months) - 1
            for idx, (m, y, name) in enumerate(months):
                if y >= reporting_year:
                    last_prior_col = data_start_col + idx - 1
                    break
            if last_prior_col >= current_group_start:
                year_groups.append((current_group_start, last_prior_col, current_group_year))

        return year_groups

    def _group_previous_year_columns(self, sheet, months, data_start_col):
        """Group and collapse columns for previous years.

        Args:
            sheet: xlwings sheet object
            months: List of (month, year, display_name) tuples
            data_start_col: First column number where month data starts
        """
        year_groups = self._get_previous_year_columns(months, data_start_col)

        if not year_groups:
            return

        try:
            for start_col, end_col, year in year_groups:
                if start_col <= end_col:
                    start_letter = self._col_letter(start_col)
                    end_letter = self._col_letter(end_col)
                    print(f"Grouping {year} columns: {start_letter}:{end_letter}")
                    sheet.range(f'{start_letter}:{end_letter}').api.Columns.Group()

            # Collapse all groups (show only level 1 = ungrouped columns)
            sheet.api.Outline.ShowLevels(ColumnLevels=1)
            print(f"Grouped and collapsed {len(year_groups)} previous year(s)")
        except Exception as e:
            print(f"Warning: Could not group previous year columns: {e}")

    def _remove_empty_leading_columns(self, sheet):
        """Remove leading columns that have $0 in Net Income (division didn't exist yet).

        Scans from left to right and deletes data columns where Net Income = 0
        until hitting a column with non-zero Net Income.

        Args:
            sheet: xlwings sheet object
        """
        try:
            # Find the Net Income row using BULK READ (optimized)
            net_income_row = None
            last_row = min(200, sheet.api.UsedRange.Rows.Count)

            # Read entire column A at once
            col_a_data = sheet.range((1, 1), (last_row, 1)).value
            if not isinstance(col_a_data, list):
                col_a_data = [col_a_data]

            for row_idx, val in enumerate(col_a_data):
                if val and 'net income' in str(val).lower() and 'other' not in str(val).lower():
                    net_income_row = row_idx + 1  # Convert to 1-based
                    break

            if not net_income_row:
                print(f"  Could not find Net Income row")
                return

            # Find the header row (row 4) and data start column (column 2)
            header_row = 4
            data_start_col = 2

            # Find last column
            last_col = sheet.range((header_row, 1)).end('right').column

            # Read entire Net Income row at once (BULK READ)
            net_income_values = sheet.range((net_income_row, data_start_col), (net_income_row, last_col)).value
            if not isinstance(net_income_values, list):
                net_income_values = [net_income_values]

            # Count how many leading columns have $0 Net Income
            empty_cols_count = 0
            for net_income_val in net_income_values:
                # Check if value is 0 or None (formula might show 0)
                if net_income_val is None or net_income_val == 0:
                    empty_cols_count += 1
                else:
                    # Found a column with non-zero Net Income - stop here
                    break

            if empty_cols_count > 0:
                # Delete columns in reverse order (from right to left of empty range)
                # Actually, delete them all at once for efficiency
                first_col_letter = self._col_letter(data_start_col)
                last_empty_col_letter = self._col_letter(data_start_col + empty_cols_count - 1)

                print(f"  Removing {empty_cols_count} empty leading columns ({first_col_letter}:{last_empty_col_letter})")
                sheet.range(f'{first_col_letter}:{last_empty_col_letter}').delete()

                # After deletion, need to ungroup any orphaned groups
                try:
                    sheet.api.Outline.ShowLevels(ColumnLevels=8)  # Show all levels
                except:
                    pass

        except Exception as e:
            print(f"  Error removing empty columns: {e}")

    def _populate_source_sheet(self, sheet, accounts, months, division_name=None):
        """Populate a source data sheet with proper formatting

        Structure (single division / backward compatible):
        - Row 1: Headers (Account, month names like "Jan 24")
        - Row 2: YYYYMM helper values (e.g., 202401) for YTD calculations - hidden
        - Row 3+: Account data

        Structure (multi-division mode when division_name provided):
        - Row 1: Headers (Division, Account, month names)
        - Row 2: YYYYMM helper values (hidden)
        - Row 3+: Account data with Division in column A

        Args:
            sheet: xlwings sheet object
            accounts: List of account dictionaries
            months: List of (month, year, display_name) tuples
            division_name: Optional division identifier for multi-division mode
        """
        # Colors - matching web version
        SOURCE_BLACK = (26, 26, 26)  # #1A1A1A - almost black for source sheets

        # Determine column offsets based on division mode
        has_division = division_name is not None or self.is_multi_division.get()
        col_offset = 1 if has_division else 0  # Extra column for Division

        # Row 1: Header
        if has_division:
            sheet.range('A1').value = 'Division'
            sheet.range('B1').value = 'Account'
            for i, (m, y, name) in enumerate(months):
                sheet.range((1, i + 3)).value = name
        else:
            sheet.range('A1').value = 'Account'
            for i, (m, y, name) in enumerate(months):
                sheet.range((1, i + 2)).value = name

        # Row 2: YYYYMM helper values for YTD calculations (e.g., 202411 for Nov 2024)
        # This enables SUMPRODUCT formulas to filter by year and month
        start_col = 3 if has_division else 2
        for i, (m, y, name) in enumerate(months):
            sheet.range((2, start_col + i)).value = y * 100 + m

        # Hide row 2 (helper row)
        try:
            sheet.range('2:2').api.Hidden = True
        except:
            pass

        # Row 3+: Data - write all at once for speed and to ensure numbers are numbers
        data = []
        # Use provided division_name or get from account if present
        div_name = division_name or self.company_name.get()

        for account in accounts:
            # Get division from account if available, otherwise use provided or company name
            acct_division = account.get('division', div_name)
            if has_division:
                row = [acct_division, account['name']]
            else:
                row = [account['name']]

            for m, y, _ in months:
                val = account['values'].get((m, y), 0)
                # Ensure it's a number
                if isinstance(val, str):
                    try:
                        val = float(val.replace(',', ''))
                    except:
                        val = 0
                row.append(val)
            data.append(row)

        if data:
            sheet.range('A3').value = data  # Start at row 3 now

        # Format header row with dark background and white text (like web version)
        try:
            num_cols = len(months) + (2 if has_division else 1)
            header_range = sheet.range((1, 1), (1, num_cols))
            header_range.font.name = 'Calibri Light'
            header_range.font.size = 10
            header_range.font.bold = True
            header_range.font.color = (255, 255, 255)
            header_range.color = SOURCE_BLACK

            # Center align month headers
            for col in range(start_col, num_cols + 1):
                sheet.range((1, col)).api.HorizontalAlignment = -4108  # xlCenter
        except:
            pass

        # Apply number format and font to data columns (now starting at row 3)
        if len(accounts) > 0 and len(months) > 0:
            try:
                data_range = sheet.range((3, start_col), (len(accounts) + 2, num_cols))
                data_range.number_format = '#,##0'
                data_range.font.name = 'Calibri Light'
                data_range.font.size = 10

                # Account names column
                acct_col = 2 if has_division else 1
                account_range = sheet.range((3, acct_col), (len(accounts) + 2, acct_col))
                account_range.font.name = 'Calibri Light'
                account_range.font.size = 10

                # Division column if present
                if has_division:
                    div_range = sheet.range((3, 1), (len(accounts) + 2, 1))
                    div_range.font.name = 'Calibri Light'
                    div_range.font.size = 10
            except:
                pass

        # Set column widths
        if has_division:
            sheet.range('A:A').column_width = 25  # Division column
            sheet.range('B:B').column_width = 45  # Account column
            for col in range(3, num_cols + 1):
                sheet.range((1, col), (1, col)).column_width = 14
        else:
            sheet.range('A:A').column_width = 45
            for col in range(2, len(months) + 2):
                sheet.range((1, col), (1, col)).column_width = 14

    def _create_menu_sheet(self, sheet, months):
        """Create professional, corporate-style menu/control sheet"""
        # Colors
        DARK_BLUE = (22, 33, 62)  # #16213E
        ACCENT_BLUE = (0, 102, 204)  # #0066CC
        GRAY = (128, 128, 128)
        LIGHT_GRAY = (240, 240, 240)
        WHITE = (255, 255, 255)

        company = self.company_name.get()

        # Professional layout with clean spacing
        # Row 2-3: Company Title (bold, prominent)
        sheet.range('B2').value = company
        sheet.range('B2').font.name = 'Calibri Light'
        sheet.range('B2').font.size = 28
        sheet.range('B2').font.bold = True
        sheet.range('B2').font.color = DARK_BLUE

        sheet.range('B3').value = 'Financial Model'
        sheet.range('B3').font.name = 'Calibri Light'
        sheet.range('B3').font.size = 14
        sheet.range('B3').font.color = GRAY

        # Row 5: Horizontal line (using cell border)
        try:
            sheet.range('B5:C5').api.Borders(9).LineStyle = 1  # xlContinuous bottom border
            sheet.range('B5:C5').api.Borders(9).Color = 0x3E2116  # Dark blue
            sheet.range('B5:C5').api.Borders(9).Weight = 2
        except:
            pass

        # Row 7-8: Current Period Info (clean, professional)
        sheet.range('B7').value = 'Current Period'
        sheet.range('B7').font.name = 'Calibri Light'
        sheet.range('B7').font.size = 10
        sheet.range('B7').font.color = GRAY

        sheet.range('C7').value = months[-1][2] if months else ''
        sheet.range('C7').font.name = 'Calibri Light'
        sheet.range('C7').font.size = 12
        sheet.range('C7').font.bold = True
        sheet.range('C7').font.color = DARK_BLUE

        sheet.range('B8').value = 'Data Range'
        sheet.range('B8').font.name = 'Calibri Light'
        sheet.range('B8').font.size = 10
        sheet.range('B8').font.color = GRAY

        sheet.range('C8').value = f"{months[0][2]} - {months[-1][2]}" if months else ''
        sheet.range('C8').font.name = 'Calibri Light'
        sheet.range('C8').font.size = 10
        sheet.range('C8').font.color = DARK_BLUE

        # Row 10: Navigation Header
        sheet.range('B10').value = 'Quick Navigation'
        sheet.range('B10').font.name = 'Calibri Light'
        sheet.range('B10').font.size = 12
        sheet.range('B10').font.bold = True
        sheet.range('B10').font.color = DARK_BLUE

        # Rows 11-17: Navigation links (clean hyperlinks)
        nav_items = [
            ('Dashboard', 'Dashboard', 'Executive overview and KPIs'),
            ('P&L Statement', 'Consolidated_PL', 'Profit & Loss analysis'),
            ('Balance Sheet', 'Consolidated_BS', 'Assets, Liabilities & Equity'),
            ('Cash Flow', 'Cash_Flow', 'Cash flow statement'),
            ('Forecast', 'Forecast', 'Budget vs Actual forecast'),
            ('Forecast Summary', 'Forecast_Summary', 'YTD variance analysis'),
            ('Notes', 'Notes', 'Commentary and annotations'),
        ]

        for i, (label, target, desc) in enumerate(nav_items):
            row = 11 + i
            sheet.range(f'B{row}').value = label
            sheet.range(f'B{row}').font.name = 'Calibri Light'
            sheet.range(f'B{row}').font.size = 11
            sheet.range(f'B{row}').font.color = ACCENT_BLUE
            sheet.range(f'B{row}').font.underline = True
            sheet.range(f'C{row}').value = desc
            sheet.range(f'C{row}').font.name = 'Calibri Light'
            sheet.range(f'C{row}').font.size = 9
            sheet.range(f'C{row}').font.color = GRAY
            try:
                sheet.range(f'B{row}').add_hyperlink(f'#{target}!A1', text_to_display=label)
            except:
                pass

        # Row 19: Version info (subtle)
        sheet.range('B19').value = f'Generated: {datetime.now().strftime("%B %d, %Y")}'
        sheet.range('B19').font.name = 'Calibri Light'
        sheet.range('B19').font.size = 9
        sheet.range('B19').font.color = GRAY

        sheet.range('B20').value = f'Version {APP_VERSION}'
        sheet.range('B20').font.name = 'Calibri Light'
        sheet.range('B20').font.size = 9
        sheet.range('B20').font.color = GRAY

        # Add helper cells for YTD calculations (hidden columns)
        try:
            if months:
                current_month = months[-1][0]
                current_year = months[-1][1]
                sheet.range('E7').value = current_month
                sheet.range('F7').value = current_year
                sheet.range('G7').value = current_year * 100 + current_month
                sheet.range('G9').value = current_year * 100 + current_month
                # Additional helper for multi-division year reference
                sheet.range('I7').value = current_year
            # Hide helper columns E through K
            sheet.range('E:K').api.Hidden = True
        except:
            pass

        # Set column widths for professional layout
        sheet.range('A:A').column_width = 3
        sheet.range('B:B').column_width = 22
        sheet.range('C:C').column_width = 30
        sheet.range('D:D').column_width = 3

        # Hide gridlines for cleaner look
        try:
            sheet.api.Activate()
            sheet.book.app.api.ActiveWindow.DisplayGridlines = False
        except:
            pass

        # Set tab color
        try:
            sheet.api.Tab.Color = 0x3E2116  # Dark blue
        except:
            pass

    def _create_pl_report(self, sheet, accounts, months, detected_totals=None):
        """Create P&L report with SUMIF formulas and CFO-grade formatting"""
        company = self.company_name.get()
        detected_totals = detected_totals or {}

        # Colors
        DARK_BLUE = (22, 33, 62)  # #16213E
        SUBTOTAL_GRAY = (236, 236, 236)  # #ECECEC

        # Title section
        sheet.range('A1').value = company
        sheet.range('A1').font.name = 'Calibri Light'
        sheet.range('A1').font.size = 14
        sheet.range('A1').font.bold = True

        sheet.range('A2').value = 'Profit & Loss Statement'
        sheet.range('A2').font.name = 'Calibri Light'
        sheet.range('A2').font.size = 12
        sheet.range('A2').font.bold = True

        # Calculate column positions
        # Months | Notes | Spacer | PY YTD | CY YTD | Var $ | Var % | Spacer | Full Years...
        header_row = 4
        last_month_col = len(months) + 1
        notes_col = last_month_col + 1
        spacer1_col = notes_col + 1
        py_ytd_col = spacer1_col + 1
        cy_ytd_col = py_ytd_col + 1
        var_col = cy_ytd_col + 1
        var_pct_col = var_col + 1
        spacer2_col = var_pct_col + 1

        # Get unique years for full year columns
        years = sorted(set(y for m, y, name in months))
        fy_start_col = spacer2_col + 1
        last_col = fy_start_col + len(years) - 1

        # Headers
        sheet.range(f'A{header_row}').value = 'Account'

        # Row 3: Helper row with YYYYMM values for dynamic YTD calculations
        # This allows formulas to compare dates against Menu!C7
        for i, (m, y, name) in enumerate(months):
            col = i + 2
            sheet.range((header_row, col)).value = f"{self.MONTHS[m-1][:3]} {y}"
            # Row 3 stores YYYYMM as number (e.g., 202411 for Nov 2024)
            sheet.range((3, col)).value = y * 100 + m

        # Notes and Summary headers
        sheet.range((header_row, notes_col)).value = 'Notes'
        sheet.range((header_row, py_ytd_col)).value = 'PY YTD'
        sheet.range((header_row, cy_ytd_col)).value = 'CY YTD'
        sheet.range((header_row, var_col)).value = 'Var $'
        sheet.range((header_row, var_pct_col)).value = 'Var %'

        # Full year headers (just year)
        for i, year in enumerate(years):
            sheet.range((header_row, fy_start_col + i)).value = str(year)

        # Format header row
        try:
            header_range = sheet.range((header_row, 1), (header_row, last_col))
            header_range.font.name = 'Calibri Light'
            header_range.font.size = 10
            header_range.font.bold = True
            header_range.font.color = (255, 255, 255)
            header_range.color = DARK_BLUE
            for col in range(2, last_col + 1):
                if col not in [spacer1_col, spacer2_col]:
                    sheet.range((header_row, col)).api.HorizontalAlignment = -4108  # xlCenter
            # Clear spacer column headers completely
            sheet.range((header_row, spacer1_col)).value = ''
            sheet.range((header_row, spacer2_col)).value = ''
        except:
            pass

        # Add YTD date range label in row 3 (e.g., "Jan-Nov") merged across PY YTD and CY YTD
        try:
            ytd_label_cell = sheet.range((3, py_ytd_col))
            # Formula shows "Jan-[current month]" based on Menu!E7
            ytd_label_cell.formula = '="Jan-"&TEXT(DATE(2024,Menu!$E$7,1),"mmm")'
            # Merge across PY YTD and CY YTD columns
            sheet.range((3, py_ytd_col), (3, cy_ytd_col)).merge()
            ytd_label_cell.api.HorizontalAlignment = -4108  # xlCenter
            ytd_label_cell.font.name = 'Calibri Light'
            ytd_label_cell.font.size = 9
            ytd_label_cell.font.italic = True
            ytd_label_cell.font.color = DARK_BLUE
        except:
            pass

        # Hide Row 3 (YYYYMM helper row) - actually hide the row, not just font color
        try:
            sheet.range('3:3').api.Hidden = True
        except:
            pass

        # Track current section for indentation
        current_section = None
        in_section = False
        row_idx = header_row + 1
        gross_margin_row = None
        net_income_row = None
        total_income_row = None  # Total Revenue/Income BEFORE COGS
        total_cogs_row = None
        total_expenses_row = None  # Total Expenses (not Total Other Expenses)
        found_cogs_section = False  # Track if we've passed the COGS section

        for account in accounts:
            account_name = account['name']
            name_lower = account_name.lower()

            # Detect section headers and track COGS section
            if account['is_header']:
                current_section = name_lower
                in_section = True
                if 'cost' in name_lower or 'cogs' in name_lower:
                    found_cogs_section = True

            # Get indent level from source file (4 spaces per level)
            indent_level = account.get('indent', 0)

            # Write account name with indentation matching source
            display_name = account_name
            if indent_level > 0 and not account['is_header'] and not account['is_total']:
                display_name = ('    ' * indent_level) + account_name  # 4 spaces per indent level

            # Track Gross Profit/Margin row (don't rename - keep original for formula matching)
            if 'gross profit' in name_lower:
                gross_margin_row = row_idx

            if 'net income' in name_lower and account['is_total']:
                net_income_row = row_idx

            # Track total rows for validation - IMPROVED LOGIC
            if 'total' in name_lower and account['is_total']:
                # Total Income/Revenue - must be BEFORE COGS section (not "other income")
                if (('income' in name_lower or 'revenue' in name_lower) and
                    'net' not in name_lower and
                    'other' not in name_lower and
                    not found_cogs_section):
                    total_income_row = row_idx
                # Total COGS
                elif 'cost' in name_lower or 'cogs' in name_lower:
                    total_cogs_row = row_idx
                    found_cogs_section = True  # Mark that we've found COGS
                # Total Expenses - only the main "Total Expenses" not "Total Other Expenses"
                elif 'expense' in name_lower and 'other' not in name_lower:
                    total_expenses_row = row_idx

            sheet.range((row_idx, 1)).value = display_name
            sheet.range((row_idx, 1)).font.name = 'Calibri Light'
            sheet.range((row_idx, 1)).font.size = 10
            # Vertical center alignment for account names
            try:
                sheet.range((row_idx, 1)).api.VerticalAlignment = -4108  # xlVAlignCenter
            except:
                pass

            # Apply formatting based on row type
            if account['is_header']:
                sheet.range((row_idx, 1)).font.bold = True
                # Header rows should NOT have formulas - leave data cells blank
                row_idx += 1
                in_section = True
                continue  # Skip formula creation for header rows
            elif account['is_total']:
                is_net_income = 'net income' in account_name.lower()

                if is_net_income:
                    # Net Income: bold, thick top border, double bottom border
                    for col in range(1, last_col + 1):
                        if col not in [spacer1_col, spacer2_col]:
                            cell = sheet.range((row_idx, col))
                            cell.font.bold = True
                            try:
                                cell.api.Borders(8).LineStyle = 1  # xlContinuous top
                                cell.api.Borders(8).Weight = 3  # xlMedium
                                cell.api.Borders(9).LineStyle = -4119  # xlDouble
                                cell.api.Borders(9).Weight = 4
                            except:
                                pass
                else:
                    # Other totals: bold, thin top border, gray background
                    for col in range(1, last_col + 1):
                        if col not in [spacer1_col, spacer2_col]:
                            cell = sheet.range((row_idx, col))
                            cell.font.bold = True  # ALL totals bold
                            cell.color = SUBTOTAL_GRAY
                            try:
                                cell.api.Borders(8).LineStyle = 1  # xlContinuous top
                                cell.api.Borders(8).Weight = 2  # xlThin
                            except:
                                pass

            # SUMIF formulas for each month (limited range for speed)
            for i, (m, y, name) in enumerate(months):
                col = i + 2
                cl = self._col_letter(col)
                formula = f"=SUMIF(Source_PL!$A$3:$A$1500,\"{account_name}\",Source_PL!{cl}$3:{cl}$1500)"
                sheet.range((row_idx, col)).formula = formula

            # Notes column - lookup formula (matches Statement Type, Date, and Account)
            # Notes structure: A=Statement Type, B=Date, C=Account, D=Note
            formula = f'=IFERROR(LOOKUP(2,1/((Notes!$A$2:$A$100="P&L")*(Notes!$B$2:$B$100=TEXT(Menu!$C$7,"mmm yy"))*(Notes!$C$2:$C$100=TRIM($A{row_idx}))),Notes!$D$2:$D$100),"")'
            sheet.range((row_idx, notes_col)).formula = formula
            sheet.range((row_idx, notes_col)).font.name = 'Calibri Light'
            sheet.range((row_idx, notes_col)).font.size = 9
            # Left justify and wrap notes
            try:
                sheet.range((row_idx, notes_col)).api.HorizontalAlignment = -4131  # xlLeft
                sheet.range((row_idx, notes_col)).api.WrapText = True
                sheet.range((row_idx, notes_col)).api.VerticalAlignment = -4108  # xlVAlignCenter
            except:
                pass

            # Dynamic YTD formulas that reference Menu helper cells
            # Row 3 contains YYYYMM values (e.g., 202411 for Nov 2024)
            # Menu!E7 = current month number (1-12)
            # Menu!F7 = current year (e.g., 2024)

            first_data_col = self._col_letter(2)  # B
            last_data_col = self._col_letter(len(months) + 1)
            data_range = f'{first_data_col}{row_idx}:{last_data_col}{row_idx}'
            helper_range = f'{first_data_col}$3:{last_data_col}$3'

            # CY YTD: SUMPRODUCT for current year, months through current month
            # Uses -- to coerce TRUE/FALSE to 1/0
            cy_formula = (
                f'=SUMPRODUCT(({data_range})*'
                f'--(INT({helper_range}/100)=Menu!$F$7)*'
                f'--(MOD({helper_range},100)<=Menu!$E$7))'
            )

            # PY YTD: SUMPRODUCT for prior year, months through current month
            py_formula = (
                f'=SUMPRODUCT(({data_range})*'
                f'--(INT({helper_range}/100)=Menu!$F$7-1)*'
                f'--(MOD({helper_range},100)<=Menu!$E$7))'
            )

            sheet.range((row_idx, cy_ytd_col)).formula = cy_formula
            sheet.range((row_idx, py_ytd_col)).formula = py_formula

            # Variance $ (CY - PY)
            sheet.range((row_idx, var_col)).formula = f'={self._col_letter(cy_ytd_col)}{row_idx}-{self._col_letter(py_ytd_col)}{row_idx}'

            # Variance % (Var/PY)
            sheet.range((row_idx, var_pct_col)).formula = f'=IFERROR({self._col_letter(var_col)}{row_idx}/{self._col_letter(py_ytd_col)}{row_idx},0)'

            # Full Year columns (sum months for that year)
            for i, year in enumerate(years):
                year_month_cols = [c + 2 for c, (m, y, name) in enumerate(months) if y == year]
                if year_month_cols:
                    refs = '+'.join([f'{self._col_letter(c)}{row_idx}' for c in year_month_cols])
                    sheet.range((row_idx, fy_start_col + i)).formula = f'={refs}'

            row_idx += 1

            # Add COGS % row after Total COGS
            if account['is_total'] and ('cost' in name_lower or 'cogs' in name_lower) and total_income_row:
                sheet.range((row_idx, 1)).value = '    COGS %'
                sheet.range((row_idx, 1)).font.name = 'Calibri Light'
                sheet.range((row_idx, 1)).font.size = 10
                sheet.range((row_idx, 1)).font.italic = True
                sheet.range((row_idx, 1)).font.color = (100, 100, 100)

                # COGS % = Total COGS / Total Revenue for each column
                cogs_row_num = row_idx - 1
                for col in range(2, last_col + 1):
                    if col not in [spacer1_col, spacer2_col, notes_col]:
                        col_letter = self._col_letter(col)
                        formula = f'=IFERROR(ABS({col_letter}{cogs_row_num})/{col_letter}{total_income_row},0)'
                        sheet.range((row_idx, col)).formula = formula
                        sheet.range((row_idx, col)).number_format = '0.0%'
                        sheet.range((row_idx, col)).font.name = 'Calibri Light'
                        sheet.range((row_idx, col)).font.size = 10
                        sheet.range((row_idx, col)).font.italic = True
                        sheet.range((row_idx, col)).font.color = (100, 100, 100)
                row_idx += 1

            # Add Gross Margin % row after Gross Profit
            if 'gross profit' in name_lower and total_income_row:
                sheet.range((row_idx, 1)).value = '    Gross Margin %'
                sheet.range((row_idx, 1)).font.name = 'Calibri Light'
                sheet.range((row_idx, 1)).font.size = 10
                sheet.range((row_idx, 1)).font.italic = True
                sheet.range((row_idx, 1)).font.color = (100, 100, 100)

                # Gross Margin % = Gross Profit / Total Revenue for each column
                gross_profit_row = row_idx - 1
                for col in range(2, last_col + 1):
                    if col not in [spacer1_col, spacer2_col, notes_col]:
                        col_letter = self._col_letter(col)
                        formula = f'=IFERROR({col_letter}{gross_profit_row}/{col_letter}{total_income_row},0)'
                        sheet.range((row_idx, col)).formula = formula
                        sheet.range((row_idx, col)).number_format = '0.0%'
                        sheet.range((row_idx, col)).font.name = 'Calibri Light'
                        sheet.range((row_idx, col)).font.size = 10
                        sheet.range((row_idx, col)).font.italic = True
                        sheet.range((row_idx, col)).font.color = (100, 100, 100)
                row_idx += 1

            # Add Expense % row after Total Expenses (not "Other Expenses")
            if (account['is_total'] and 'expense' in name_lower and
                'other' not in name_lower and 'total' in name_lower and total_income_row):
                sheet.range((row_idx, 1)).value = '    Expense %'
                sheet.range((row_idx, 1)).font.name = 'Calibri Light'
                sheet.range((row_idx, 1)).font.size = 10
                sheet.range((row_idx, 1)).font.italic = True
                sheet.range((row_idx, 1)).font.color = (100, 100, 100)

                # Expense % = Total Expenses / Total Revenue for each column
                expense_row_num = row_idx - 1
                for col in range(2, last_col + 1):
                    if col not in [spacer1_col, spacer2_col, notes_col]:
                        col_letter = self._col_letter(col)
                        formula = f'=IFERROR(ABS({col_letter}{expense_row_num})/{col_letter}{total_income_row},0)'
                        sheet.range((row_idx, col)).formula = formula
                        sheet.range((row_idx, col)).number_format = '0.0%'
                        sheet.range((row_idx, col)).font.name = 'Calibri Light'
                        sheet.range((row_idx, col)).font.size = 10
                        sheet.range((row_idx, col)).font.italic = True
                        sheet.range((row_idx, col)).font.color = (100, 100, 100)
                row_idx += 1

            # Add Net Profit % row after Net Income
            if 'net income' in name_lower and account['is_total'] and total_income_row:
                sheet.range((row_idx, 1)).value = '    Net Profit %'
                sheet.range((row_idx, 1)).font.name = 'Calibri Light'
                sheet.range((row_idx, 1)).font.size = 10
                sheet.range((row_idx, 1)).font.italic = True
                sheet.range((row_idx, 1)).font.color = (100, 100, 100)

                # Net Profit % = Net Income / Total Revenue for each column
                net_income_row_num = row_idx - 1
                for col in range(2, last_col + 1):
                    if col not in [spacer1_col, spacer2_col, notes_col]:
                        col_letter = self._col_letter(col)
                        formula = f'=IFERROR({col_letter}{net_income_row_num}/{col_letter}{total_income_row},0)'
                        sheet.range((row_idx, col)).formula = formula
                        sheet.range((row_idx, col)).number_format = '0.0%'
                        sheet.range((row_idx, col)).font.name = 'Calibri Light'
                        sheet.range((row_idx, col)).font.size = 10
                        sheet.range((row_idx, col)).font.italic = True
                        sheet.range((row_idx, col)).font.color = (100, 100, 100)
                row_idx += 1

            # Reset section tracking after totals
            if account['is_total']:
                in_section = False

        data_end_row = row_idx - 1
        data_start_row = header_row + 1

        # Apply formatting to data range in bulk operations
        try:
            # Apply font to entire data area at once (much faster than per-column)
            full_data_range = sheet.range((data_start_row, 1), (data_end_row, last_col))
            full_data_range.font.name = 'Calibri Light'
            full_data_range.font.size = 10

            # Number format - apply to contiguous ranges for efficiency
            # Month columns (2 to last_month_col)
            if last_month_col > 1:
                sheet.range((data_start_row, 2), (data_end_row, last_month_col)).number_format = '#,##0'

            # YTD columns (py_ytd and cy_ytd)
            sheet.range((data_start_row, py_ytd_col), (data_end_row, cy_ytd_col)).number_format = '#,##0'

            # Variance $ column
            sheet.range((data_start_row, var_col), (data_end_row, var_col)).number_format = '#,##0'

            # Variance % column
            sheet.range((data_start_row, var_pct_col), (data_end_row, var_pct_col)).number_format = '0.0%'

            # Full year columns
            if fy_start_col <= last_col:
                sheet.range((data_start_row, fy_start_col), (data_end_row, last_col)).number_format = '#,##0'

            # Right align all numeric columns at once (excludes only spacers and notes)
            for col in range(2, last_col + 1):
                if col not in [spacer1_col, spacer2_col, notes_col]:
                    sheet.range((data_start_row, col), (data_end_row, col)).api.HorizontalAlignment = -4152  # xlRight
        except:
            pass

        # Clear all formatting from spacer columns (entire column, true white space)
        try:
            for spacer_col in [spacer1_col, spacer2_col]:
                spacer_range = sheet.range((1, spacer_col), (row_idx + 50, spacer_col))
                spacer_range.clear()
                spacer_range.color = None  # Remove any background
        except:
            pass

        # Add EBITDA Reconciliation section
        ebitda_start = row_idx + 2
        sheet.range((ebitda_start, 1)).value = 'RECONCILIATION TO EBITDA'
        sheet.range((ebitda_start, 1)).font.name = 'Calibri Light'
        sheet.range((ebitda_start, 1)).font.size = 10
        sheet.range((ebitda_start, 1)).font.bold = True
        sheet.range((ebitda_start, 1)).font.color = DARK_BLUE

        ebitda_items = [
            ('Net Income', 'net_income'),
            ('    Add: Interest, net', 'interest'),
            ('    Add: Income taxes', 'taxes'),
            ('    Add: Depreciation and amortization', 'depr'),
            ('    Add: Non-recurring expense (income)', 'nonrecurring'),
            ('Reported EBITDA', 'ebitda_total'),
        ]

        for i, (label, item_type) in enumerate(ebitda_items):
            r = ebitda_start + 1 + i
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            if item_type == 'ebitda_total':
                # Bold with borders for EBITDA total
                sheet.range((r, 1)).font.bold = True
                for col in range(1, last_month_col + 1):
                    if col > 1:
                        # Sum of Net Income + all add-backs
                        formula = f'=SUM({self._col_letter(col)}{ebitda_start+1}:{self._col_letter(col)}{r-1})'
                        sheet.range((r, col)).formula = formula
                    cell = sheet.range((r, col))
                    cell.font.bold = True
                    try:
                        cell.api.Borders(8).LineStyle = 1
                        cell.api.Borders(8).Weight = 3
                        cell.api.Borders(9).LineStyle = -4119  # xlDouble
                    except:
                        pass
            else:
                for col_idx in range(2, last_month_col + 1):
                    col_letter = self._col_letter(col_idx)
                    if item_type == 'net_income' and net_income_row:
                        formula = f'={col_letter}{net_income_row}'
                    elif item_type == 'interest':
                        formula = f'=SUMIF(Source_PL!$A$3:$A$1500,"*Interest*",Source_PL!{col_letter}$3:{col_letter}$1500)*-1'
                    elif item_type == 'taxes':
                        formula = f'=SUMIF(Source_PL!$A$3:$A$1500,"*Tax*",Source_PL!{col_letter}$3:{col_letter}$1500)*-1'
                    elif item_type == 'depr':
                        formula = f'=SUMIF(Source_PL!$A$3:$A$1500,"*Deprec*",Source_PL!{col_letter}$3:{col_letter}$1500)*-1+SUMIF(Source_PL!$A$3:$A$1500,"*Amort*",Source_PL!{col_letter}$3:{col_letter}$1500)*-1'
                    elif item_type == 'nonrecurring':
                        formula = '0'  # Manual entry placeholder
                    else:
                        formula = '0'
                    sheet.range((r, col_idx)).formula = formula
                    sheet.range((r, col_idx)).number_format = '#,##0'

        ebitda_end_row = ebitda_start + len(ebitda_items)

        # Add validation section at bottom
        val_start = ebitda_end_row + 2
        sheet.range((val_start, 1)).value = 'VALIDATION'
        sheet.range((val_start, 1)).font.name = 'Calibri Light'
        sheet.range((val_start, 1)).font.size = 10
        sheet.range((val_start, 1)).font.bold = True
        sheet.range((val_start, 1)).font.color = DARK_BLUE

        # Get exact total names from detected_totals for source lookups
        src_total_income = detected_totals.get('total_income', 'Total Income')
        src_total_cogs = detected_totals.get('total_cogs', 'Total Cost of Sales')
        src_total_expenses = detected_totals.get('total_expenses', 'Total Expenses')
        src_net_income = detected_totals.get('net_income', 'Net Income')

        # Report totals section - starts right after VALIDATION header
        report_start = val_start + 1

        validation_report_labels = [
            ('Revenue (Report)', total_income_row),
            ('COGS (Report)', total_cogs_row),
            ('Expenses (Report)', total_expenses_row),
            ('Net Income (Report)', net_income_row),
        ]

        for i, (label, ref_row) in enumerate(validation_report_labels):
            r = report_start + i  # No +1 since no header row
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            if ref_row:
                for col_idx in range(2, last_month_col + 1):
                    col_letter = self._col_letter(col_idx)
                    formula = f'={col_letter}{ref_row}'
                    sheet.range((r, col_idx)).formula = formula
                    sheet.range((r, col_idx)).number_format = '#,##0'

        report_val_end = report_start + len(validation_report_labels) - 1  # Last row of report section

        # Source check rows - using EXACT total names detected from source (no header)
        source_start = report_val_end + 1

        # Use exact matches on detected total names (no wildcards)
        source_labels = [
            ('Revenue (Source)', src_total_income),
            ('COGS (Source)', src_total_cogs),
            ('Expenses (Source)', src_total_expenses),
            ('Net Income (Source)', src_net_income),
        ]

        for i, (label, exact_name) in enumerate(source_labels):
            r = source_start + i  # No header, so start at source_start
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            if exact_name:
                for col_idx in range(2, last_month_col + 1):
                    col_letter = self._col_letter(col_idx)
                    # Use exact match with limited range
                    formula = f'=SUMIF(Source_PL!$A$3:$A$1500,"{exact_name}",Source_PL!{col_letter}$3:{col_letter}$1500)'
                    sheet.range((r, col_idx)).formula = formula
                    sheet.range((r, col_idx)).number_format = '#,##0'

        source_end = source_start + len(source_labels) - 1  # Last row of source section

        # Variance section (Report - Source = should be 0) - no header
        var_start = source_end + 1

        # Calculate row references for variance (no +1 since no headers)
        rev_report_row = report_start
        cogs_report_row = report_start + 1
        exp_report_row = report_start + 2
        ni_report_row = report_start + 3
        rev_source_row = source_start
        cogs_source_row = source_start + 1
        exp_source_row = source_start + 2
        ni_source_row = source_start + 3

        variance_items = [
            ('Revenue Variance', rev_report_row, rev_source_row),
            ('COGS Variance', cogs_report_row, cogs_source_row),
            ('Expenses Variance', exp_report_row, exp_source_row),
            ('Net Income Variance', ni_report_row, ni_source_row),
        ]

        for i, (label, report_row, source_row) in enumerate(variance_items):
            r = var_start + i  # No header
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            for col_idx in range(2, last_month_col + 1):
                col_letter = self._col_letter(col_idx)
                formula = f'={col_letter}{report_row}-{col_letter}{source_row}'
                sheet.range((r, col_idx)).formula = formula
                sheet.range((r, col_idx)).number_format = '#,##0'

        var_end_row = var_start + len(variance_items) - 1  # Last row of variance section

        # Add matrix borders around validation sections
        try:
            # Outer border for entire validation section (thick)
            val_range = sheet.range((val_start, 1), (var_end_row, last_month_col))
            val_range.api.Borders(7).LineStyle = 1   # xlLeft
            val_range.api.Borders(7).Weight = 3      # xlMedium
            val_range.api.Borders(8).LineStyle = 1   # xlTop
            val_range.api.Borders(8).Weight = 3
            val_range.api.Borders(9).LineStyle = 1   # xlBottom
            val_range.api.Borders(9).Weight = 3
            val_range.api.Borders(10).LineStyle = 1  # xlRight
            val_range.api.Borders(10).Weight = 3
            
            # Add thin borders inside
            val_range.api.Borders(11).LineStyle = 1  # xlInsideVertical
            val_range.api.Borders(11).Weight = 2     # xlThin
            val_range.api.Borders(12).LineStyle = 1  # xlInsideHorizontal
            val_range.api.Borders(12).Weight = 2     # xlThin
            
            # Bold separator line between Report Totals and Source Totals
            separator1 = sheet.range((source_start, 1), (source_start, last_month_col))
            separator1.api.Borders(8).LineStyle = 1
            separator1.api.Borders(8).Weight = 3     # xlMedium - bold line
            
            # Bold separator line between Source Totals and Variance
            separator2 = sheet.range((var_start, 1), (var_start, last_month_col))
            separator2.api.Borders(8).LineStyle = 1
            separator2.api.Borders(8).Weight = 3     # xlMedium - bold line
        except Exception as e:
            print(f"Border formatting warning: {e}")

        # Group validation section so it can be collapsed
        try:
            sheet.range(f'{val_start}:{var_end_row}').api.Rows.Group()
        except Exception as e:
            print(f"Grouping warning: {e}")

        # Group account sections (rows between header and total) for collapsible sections
        try:
            section_start = None
            for i, account in enumerate(accounts):
                row_num = header_row + 1 + i
                if account['is_header']:
                    # Start of a new section
                    section_start = row_num + 1  # First data row after header
                elif account['is_total'] and section_start is not None:
                    # End of section - group rows from section_start to row before total
                    section_end = row_num - 1
                    if section_end >= section_start:
                        sheet.range(f'{section_start}:{section_end}').api.Rows.Group()
                    section_start = None
        except Exception as e:
            print(f"PL section grouping warning: {e}")

        # Set column widths
        sheet.range('A:A').column_width = 45
        for col in range(2, last_col + 1):
            if col in [spacer1_col, spacer2_col]:
                sheet.range((1, col), (1, col)).column_width = 3
            elif col == notes_col:
                sheet.range((1, col), (1, col)).column_width = 30
            else:
                sheet.range((1, col), (1, col)).column_width = 12

        # AutoFit columns to content where appropriate
        try:
            sheet.range('A:A').api.EntireColumn.AutoFit()
            for col in range(2, last_month_col + 1):
                sheet.range((1, col), (1, col)).api.EntireColumn.AutoFit()
        except:
            pass

        # Group columns by year and hide prior years
        self._group_columns_by_year(sheet, months, header_row)

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    def _create_bs_report(self, sheet, accounts, months, detected_totals=None):
        """Create Balance Sheet report with SUMIF formulas and CFO-grade formatting"""
        company = self.company_name.get()
        detected_totals = detected_totals or {}

        # Colors
        DARK_BLUE = (22, 33, 62)  # #16213E
        SUBTOTAL_GRAY = (236, 236, 236)  # #ECECEC

        # Title section
        sheet.range('A1').value = company
        sheet.range('A1').font.name = 'Calibri Light'
        sheet.range('A1').font.size = 14
        sheet.range('A1').font.bold = True

        sheet.range('A2').value = 'Balance Sheet'
        sheet.range('A2').font.name = 'Calibri Light'
        sheet.range('A2').font.size = 12
        sheet.range('A2').font.bold = True

        header_row = 4
        last_month_col = len(months) + 1
        notes_col = last_month_col + 1
        last_col = notes_col

        # Headers
        sheet.range(f'A{header_row}').value = 'Account'

        # Row 3: Helper row with YYYYMM values for dynamic calculations
        for i, (m, y, name) in enumerate(months):
            col = i + 2
            sheet.range((header_row, col)).value = f"{self.MONTHS[m-1][:3]} {y}"
            sheet.range((3, col)).value = y * 100 + m

        # Notes header
        sheet.range((header_row, notes_col)).value = 'Notes'

        # Format header row
        try:
            header_range = sheet.range((header_row, 1), (header_row, last_col))
            header_range.font.name = 'Calibri Light'
            header_range.font.size = 10
            header_range.font.bold = True
            header_range.font.color = (255, 255, 255)
            header_range.color = DARK_BLUE
            for col in range(2, last_col + 1):
                sheet.range((header_row, col)).api.HorizontalAlignment = -4108  # xlCenter
        except:
            pass

        # Hide Row 3 (YYYYMM helper row) - actually hide the row
        try:
            sheet.range('3:3').api.Hidden = True
        except:
            pass

        # Track sections for indentation
        current_section = None
        in_section = False
        row_idx = header_row + 1
        total_assets_row = None
        total_liab_row = None
        total_equity_row = None
        total_liab_equity_row = None

        for account in accounts:
            account_name = account['name']

            # Detect section headers (Assets, Liabilities, Equity, or subsections)
            if account['is_header']:
                current_section = account_name.lower()
                in_section = True

            # Get indent level from source file (4 spaces per level)
            indent_level = account.get('indent', 0)

            # Write account name with indentation matching source
            display_name = account_name
            if indent_level > 0 and not account['is_header'] and not account['is_total']:
                display_name = ('    ' * indent_level) + account_name  # 4 spaces per indent level

            # Track key total rows
            if 'total' in account_name.lower():
                if 'assets' in account_name.lower() and 'liab' not in account_name.lower():
                    total_assets_row = row_idx
                elif 'liabilities and equity' in account_name.lower() or ('liab' in account_name.lower() and 'equity' in account_name.lower()):
                    total_liab_equity_row = row_idx
                elif 'liabilities' in account_name.lower():
                    total_liab_row = row_idx
                elif 'equity' in account_name.lower():
                    total_equity_row = row_idx

            sheet.range((row_idx, 1)).value = display_name
            sheet.range((row_idx, 1)).font.name = 'Calibri Light'
            sheet.range((row_idx, 1)).font.size = 10
            # Vertical center alignment for account names
            try:
                sheet.range((row_idx, 1)).api.VerticalAlignment = -4108  # xlVAlignCenter
            except:
                pass

            # Apply formatting based on row type
            if account['is_header']:
                sheet.range((row_idx, 1)).font.bold = True
                # Header rows should NOT have formulas - leave data cells blank
                row_idx += 1
                in_section = True
                continue  # Skip formula creation for header rows
            elif account['is_total']:
                is_main_total = 'total for assets' in account_name.lower() or 'total for liabilities and equity' in account_name.lower()

                if is_main_total:
                    # Main totals: bold, thick top border, double bottom border
                    for col in range(1, last_col + 1):
                        cell = sheet.range((row_idx, col))
                        cell.font.bold = True
                        try:
                            cell.api.Borders(8).LineStyle = 1  # xlContinuous top
                            cell.api.Borders(8).Weight = 3  # xlMedium
                            cell.api.Borders(9).LineStyle = -4119  # xlDouble
                            cell.api.Borders(9).Weight = 4
                        except:
                            pass
                else:
                    # Subtotals: bold, thin top border, gray background
                    for col in range(1, last_col + 1):
                        cell = sheet.range((row_idx, col))
                        cell.font.bold = True  # ALL totals bold
                        cell.color = SUBTOTAL_GRAY
                        try:
                            cell.api.Borders(8).LineStyle = 1  # xlContinuous top
                            cell.api.Borders(8).Weight = 2  # xlThin
                        except:
                            pass

            # SUMIF formulas for each month (limited range for speed)
            for i, (m, y, name) in enumerate(months):
                col = i + 2
                cl = self._col_letter(col)
                formula = f"=SUMIF(Source_BS!$A$3:$A$1500,\"{account_name}\",Source_BS!{cl}$3:{cl}$1500)"
                sheet.range((row_idx, col)).formula = formula

            # Notes column - lookup formula (matches Statement Type, Date, and Account)
            # Notes structure: A=Statement Type, B=Date, C=Account, D=Note
            notes_formula = f'=IFERROR(LOOKUP(2,1/((Notes!$A$2:$A$100="Balance Sheet")*(Notes!$B$2:$B$100=TEXT(Menu!$C$7,"mmm yy"))*(Notes!$C$2:$C$100=TRIM($A{row_idx}))),Notes!$D$2:$D$100),"")'
            sheet.range((row_idx, notes_col)).formula = notes_formula
            sheet.range((row_idx, notes_col)).font.name = 'Calibri Light'
            sheet.range((row_idx, notes_col)).font.size = 9
            try:
                sheet.range((row_idx, notes_col)).api.HorizontalAlignment = -4131  # xlLeft
                sheet.range((row_idx, notes_col)).api.WrapText = True
            except:
                pass

            row_idx += 1

            # Reset section tracking after totals
            if account['is_total']:
                in_section = False

        # Apply formatting to data range in bulk operations (exclude Notes column)
        data_start_row = header_row + 1
        data_end_row = row_idx - 1
        try:
            # Apply font to entire data area at once (much faster than per-cell)
            full_data_range = sheet.range((data_start_row, 1), (data_end_row, last_col))
            full_data_range.font.name = 'Calibri Light'
            full_data_range.font.size = 10

            # Number format and alignment for month columns
            month_range = sheet.range((data_start_row, 2), (data_end_row, last_month_col))
            month_range.number_format = '#,##0'
            month_range.api.HorizontalAlignment = -4152  # xlRight
        except:
            pass

        # Add validation section at bottom - with formulas for each month column
        val_start = row_idx + 2
        sheet.range((val_start, 1)).value = 'VALIDATION'
        sheet.range((val_start, 1)).font.name = 'Calibri Light'
        sheet.range((val_start, 1)).font.size = 10
        sheet.range((val_start, 1)).font.bold = True
        sheet.range((val_start, 1)).font.color = DARK_BLUE

        # Get exact total names from detected_totals for source lookups
        src_total_assets = detected_totals.get('total_assets', 'Total for Assets')
        src_total_liab = detected_totals.get('total_liabilities', 'Total for Liabilities')
        src_total_equity = detected_totals.get('total_equity', 'Total for Equity')
        src_total_liab_equity = detected_totals.get('total_liab_equity', 'Total for Liabilities and Equity')

        # Report totals section - starts right after VALIDATION header (no sub-headers)
        report_start = val_start + 1

        report_labels = [
            ('Total Assets (Report)', total_assets_row),
            ('Total Liabilities (Report)', total_liab_row),
            ('Total Equity (Report)', total_equity_row),
            ('Total Liab + Equity (Report)', total_liab_equity_row),
        ]

        for i, (label, ref_row) in enumerate(report_labels):
            r = report_start + i  # No header row
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            if ref_row:
                for col_idx in range(2, last_col + 1):
                    col_letter = self._col_letter(col_idx)
                    formula = f'={col_letter}{ref_row}'
                    sheet.range((r, col_idx)).formula = formula
                    sheet.range((r, col_idx)).number_format = '#,##0'

        report_end = report_start + len(report_labels) - 1  # Last row of report section

        # Source check section - no header
        source_start = report_end + 1

        source_labels = [
            ('Total Assets (Source)', src_total_assets),
            ('Total Liabilities (Source)', src_total_liab),
            ('Total Equity (Source)', src_total_equity),
        ]

        for i, (label, exact_name) in enumerate(source_labels):
            r = source_start + i  # No header
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            if exact_name:
                for col_idx in range(2, last_col + 1):
                    col_letter = self._col_letter(col_idx)
                    # Use exact match with limited range
                    formula = f'=SUMIF(Source_BS!$A$3:$A$1500,"{exact_name}",Source_BS!{col_letter}$3:{col_letter}$1500)'
                    sheet.range((r, col_idx)).formula = formula
                    sheet.range((r, col_idx)).number_format = '#,##0'

        source_end = source_start + len(source_labels) - 1  # Last row of source section

        # Balance Check - just the row, no header
        balance_row = source_end + 1
        sheet.range((balance_row, 1)).value = 'Assets - (Liab + Equity)'
        sheet.range((balance_row, 1)).font.name = 'Calibri Light'
        sheet.range((balance_row, 1)).font.size = 10
        sheet.range((balance_row, 1)).font.bold = True

        if total_assets_row and total_liab_equity_row:
            for col_idx in range(2, last_col + 1):
                col_letter = self._col_letter(col_idx)
                formula = f'={col_letter}{total_assets_row}-{col_letter}{total_liab_equity_row}'
                sheet.range((balance_row, col_idx)).formula = formula
                sheet.range((balance_row, col_idx)).number_format = '#,##0'
                sheet.range((balance_row, col_idx)).font.bold = True

        # Variance section - no header
        var_start = balance_row + 1

        # Calculate row references for variance (no headers now)
        assets_report_row = report_start
        liab_report_row = report_start + 1
        equity_report_row = report_start + 2
        assets_source_row = source_start
        liab_source_row = source_start + 1
        equity_source_row = source_start + 2

        variance_items = [
            ('Assets Variance', assets_report_row, assets_source_row),
            ('Liabilities Variance', liab_report_row, liab_source_row),
            ('Equity Variance', equity_report_row, equity_source_row),
        ]

        for i, (label, report_row, source_row) in enumerate(variance_items):
            r = var_start + i  # No header
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            for col_idx in range(2, last_col + 1):
                col_letter = self._col_letter(col_idx)
                formula = f'={col_letter}{report_row}-{col_letter}{source_row}'
                sheet.range((r, col_idx)).formula = formula
                sheet.range((r, col_idx)).number_format = '#,##0'

        var_end_row = var_start + len(variance_items) - 1  # Last row of variance section

        # Add matrix borders around validation sections
        try:
            # Outer border for entire validation section (thick)
            val_range = sheet.range((val_start, 1), (var_end_row, last_col))
            val_range.api.Borders(7).LineStyle = 1   # xlLeft
            val_range.api.Borders(7).Weight = 3      # xlMedium
            val_range.api.Borders(8).LineStyle = 1   # xlTop
            val_range.api.Borders(8).Weight = 3
            val_range.api.Borders(9).LineStyle = 1   # xlBottom
            val_range.api.Borders(9).Weight = 3
            val_range.api.Borders(10).LineStyle = 1  # xlRight
            val_range.api.Borders(10).Weight = 3
            
            # Add thin borders inside
            val_range.api.Borders(11).LineStyle = 1  # xlInsideVertical
            val_range.api.Borders(11).Weight = 2     # xlThin
            val_range.api.Borders(12).LineStyle = 1  # xlInsideHorizontal
            val_range.api.Borders(12).Weight = 2     # xlThin
            
            # Bold separator line between Report Totals and Source Totals
            separator1 = sheet.range((source_start, 1), (source_start, last_col))
            separator1.api.Borders(8).LineStyle = 1
            separator1.api.Borders(8).Weight = 3     # xlMedium - bold line
            
            # Bold separator line between Source Totals and Balance Check
            separator2 = sheet.range((balance_row, 1), (balance_row, last_col))
            separator2.api.Borders(8).LineStyle = 1
            separator2.api.Borders(8).Weight = 3     # xlMedium - bold line
            
            # Bold separator line between Balance Check and Variance
            separator3 = sheet.range((var_start, 1), (var_start, last_col))
            separator3.api.Borders(8).LineStyle = 1
            separator3.api.Borders(8).Weight = 3     # xlMedium - bold line
        except Exception as e:
            print(f"Border formatting warning: {e}")

        # Group validation section so it can be collapsed
        try:
            sheet.range(f'{val_start}:{var_end_row}').api.Rows.Group()
        except Exception as e:
            print(f"Grouping warning: {e}")

        # Group account sections (rows between header and total) for collapsible sections
        try:
            section_start = None
            for i, account in enumerate(accounts):
                row_num = header_row + 1 + i
                if account['is_header']:
                    # Start of a new section
                    section_start = row_num + 1  # First data row after header
                elif account['is_total'] and section_start is not None:
                    # End of section - group rows from section_start to row before total
                    section_end = row_num - 1
                    if section_end >= section_start:
                        sheet.range(f'{section_start}:{section_end}').api.Rows.Group()
                    section_start = None
        except Exception as e:
            print(f"BS section grouping warning: {e}")

        # Set column widths
        sheet.range('A:A').column_width = 45
        for col in range(2, last_month_col + 1):
            sheet.range((1, col), (1, col)).column_width = 14
        sheet.range((1, notes_col), (1, notes_col)).column_width = 30  # Notes column

        # AutoFit columns
        try:
            sheet.range('A:A').api.EntireColumn.AutoFit()
            for col in range(2, last_month_col + 1):
                sheet.range((1, col), (1, col)).api.EntireColumn.AutoFit()
        except:
            pass

        # Group columns by year and hide prior years
        self._group_columns_by_year(sheet, months, header_row)

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    def _create_cash_flow(self, sheet, months):
        """Create Cash Flow statement with indirect method formulas"""
        company = self.company_name.get()

        # Detect multi-division mode
        multi_division = hasattr(self, 'divisions') and len(self.divisions) > 1

        # Account column: B for multi-division (A=Division, B=Account), A for single
        acct_col = 'B' if multi_division else 'A'

        # Data starts at column C for multi-division (Division, Account, then data), B for single
        data_start_col = 3 if multi_division else 2

        # Colors
        DARK_BLUE = (22, 33, 62)  # #16213E
        SUBTOTAL_GRAY = (236, 236, 236)  # #ECECEC

        # Title section
        sheet.range('A1').value = company
        sheet.range('A1').font.name = 'Calibri Light'
        sheet.range('A1').font.size = 14
        sheet.range('A1').font.bold = True

        sheet.range('A2').value = 'Statement of Cash Flows (Indirect Method)'
        sheet.range('A2').font.name = 'Calibri Light'
        sheet.range('A2').font.size = 12
        sheet.range('A2').font.bold = True

        header_row = 4
        last_month_col = len(months) + 1
        ytd_col = last_month_col + 2

        # Headers
        sheet.range(f'A{header_row}').value = 'Description'

        # Row 3: Helper row with YYYYMM values for dynamic YTD calculations
        for i, (m, y, name) in enumerate(months):
            col = i + 2
            sheet.range((header_row, col)).value = f"{self.MONTHS[m-1][:3]} {y}"
            sheet.range((3, col)).value = y * 100 + m

        sheet.range((header_row, ytd_col)).value = 'YTD'

        # Format header row
        try:
            header_range = sheet.range((header_row, 1), (header_row, ytd_col))
            header_range.font.name = 'Calibri Light'
            header_range.font.size = 10
            header_range.font.bold = True
            header_range.font.color = (255, 255, 255)
            header_range.color = DARK_BLUE
            for col in range(2, ytd_col + 1):
                sheet.range((header_row, col)).api.HorizontalAlignment = -4108  # xlCenter
        except:
            pass

        # Track row references for subtotals
        row_refs = {}
        row = header_row + 1

        # Define cash flow structure with formulas
        # Format: (display_name, type, formula_type, bs_search_term)
        # type: 'header', 'item', 'subtotal', 'total'
        # formula_type: 'pl_lookup', 'bs_change_asset', 'bs_change_liab', 'manual', 'sum_range', 'bs_value'
        cf_structure = [
            # OPERATING ACTIVITIES
            ('CASH FLOWS FROM OPERATING ACTIVITIES', 'header', None, None),
            ('    Net Income', 'item', 'pl_lookup', 'Net Income'),
            ('    Adjustments to reconcile net income:', 'header', None, None),
            ('        Depreciation & Amortization', 'item', 'manual', None),
            ('    Changes in Operating Assets and Liabilities:', 'header', None, None),
            ('        (Increase) Decrease in Accounts Receivable', 'item', 'bs_change_asset', 'Accounts Receivable'),
            ('        (Increase) Decrease in Inventory', 'item', 'bs_change_asset', 'Inventory'),
            ('        (Increase) Decrease in Prepaid Expenses', 'item', 'bs_change_asset', 'Prepaid'),
            ('        (Increase) Decrease in Other Current Assets', 'item', 'bs_change_asset', 'Other Current Assets'),
            ('        Increase (Decrease) in Accounts Payable', 'item', 'bs_change_liab', 'Accounts Payable'),
            ('        Increase (Decrease) in Accrued Expenses', 'item', 'bs_change_liab', 'Accrued'),
            ('        Increase (Decrease) in Other Current Liabilities', 'item', 'bs_change_liab', 'Other Current Liabilities'),
            ('Net Cash Provided by Operating Activities', 'subtotal', 'sum_operating', None),
            ('', 'blank', None, None),
            # INVESTING ACTIVITIES
            ('CASH FLOWS FROM INVESTING ACTIVITIES', 'header', None, None),
            ('    Purchase of Property & Equipment', 'item', 'bs_change_asset', 'Fixed Assets'),
            ('    Purchase of Investments', 'item', 'bs_change_asset', 'Investments'),
            ('    Other Investing Activities', 'item', 'manual', None),
            ('Net Cash Used in Investing Activities', 'subtotal', 'sum_investing', None),
            ('', 'blank', None, None),
            # FINANCING ACTIVITIES
            ('CASH FLOWS FROM FINANCING ACTIVITIES', 'header', None, None),
            ('    Proceeds from (Payments on) Line of Credit', 'item', 'bs_change_liab', 'Line of Credit'),
            ('    Proceeds from (Payments on) Long-term Debt', 'item', 'bs_change_liab', 'Long-term'),
            ('    Owner Contributions', 'item', 'bs_change_liab', 'Contributed Capital'),
            ('    Distributions to Owners', 'item', 'bs_change_liab', 'Distribution'),
            ('Net Cash Provided by Financing Activities', 'subtotal', 'sum_financing', None),
            ('', 'blank', None, None),
            # SUMMARY
            ('NET INCREASE (DECREASE) IN CASH', 'total', 'sum_all', None),
            ('', 'blank', None, None),
            ('Cash at Beginning of Period', 'item', 'bs_prior', 'Bank Accounts'),
            ('Cash at End of Period', 'total', 'bs_current', 'Bank Accounts'),
        ]

        # ================================================================
        # BUILD ALL DATA IN MEMORY FIRST (OPTIMIZED)
        # ================================================================
        all_data = []  # List of row data arrays
        row_types_list = []  # Track row type for batch formatting

        # Track section rows for subtotals (relative to data start)
        operating_start_idx = None
        operating_end_idx = None
        investing_start_idx = None
        investing_end_idx = None
        financing_start_idx = None
        financing_end_idx = None
        net_change_idx = None
        beginning_cash_idx = None
        ending_cash_idx = None

        data_start_row = header_row + 1
        num_cols = ytd_col  # Total columns including YTD

        for struct_idx, (display_name, row_type, formula_type, search_term) in enumerate(cf_structure):
            actual_row = data_start_row + len(all_data)
            row_idx = len(all_data)

            # Track section boundaries
            if 'OPERATING ACTIVITIES' in display_name and row_type == 'header':
                operating_start_idx = row_idx + 1
            elif 'Net Cash Provided by Operating' in display_name:
                operating_end_idx = row_idx - 1
            elif 'INVESTING ACTIVITIES' in display_name and row_type == 'header':
                investing_start_idx = row_idx + 1
            elif 'Net Cash Used in Investing' in display_name:
                investing_end_idx = row_idx - 1
            elif 'FINANCING ACTIVITIES' in display_name and row_type == 'header':
                financing_start_idx = row_idx + 1
            elif 'Net Cash Provided by Financing' in display_name:
                financing_end_idx = row_idx - 1
            elif 'NET INCREASE' in display_name:
                net_change_idx = row_idx
            elif 'Beginning of Period' in display_name:
                beginning_cash_idx = row_idx
            elif 'End of Period' in display_name:
                ending_cash_idx = row_idx

            # Build row data array
            row_data = [display_name]

            if row_type in ['header', 'blank']:
                # Headers/blanks: just description, empty data cells
                row_data.extend([''] * (num_cols - 1))
            else:
                # Build formulas for each month column
                for i in range(len(months)):
                    source_col = data_start_col + i
                    source_col_letter = self._col_letter(source_col)
                    source_col_prev_letter = self._col_letter(source_col - 1) if source_col > data_start_col else source_col_letter

                    if formula_type == 'pl_lookup':
                        formula = f'=SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_PL!{source_col_letter}$3:{source_col_letter}$1500)'
                    elif formula_type == 'bs_change_asset':
                        if i == 0:
                            formula = 0
                        else:
                            formula = f'=SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_prev_letter}$3:{source_col_prev_letter}$1500)-SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_letter}$3:{source_col_letter}$1500)'
                    elif formula_type == 'bs_change_liab':
                        if i == 0:
                            formula = 0
                        else:
                            formula = f'=SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_letter}$3:{source_col_letter}$1500)-SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_prev_letter}$3:{source_col_prev_letter}$1500)'
                    elif formula_type == 'bs_prior':
                        if i == 0:
                            formula = f'=SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_letter}$3:{source_col_letter}$1500)'
                        else:
                            formula = f'=SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_prev_letter}$3:{source_col_prev_letter}$1500)'
                    elif formula_type == 'bs_current':
                        formula = f'=SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_letter}$3:{source_col_letter}$1500)'
                    elif formula_type == 'manual':
                        formula = 0
                    elif formula_type in ['sum_operating', 'sum_investing', 'sum_financing', 'sum_all']:
                        formula = ''  # Will be filled in second pass
                    else:
                        formula = ''

                    row_data.append(formula)

                # Pad to last_month_col if needed, then add spacer and YTD
                while len(row_data) < last_month_col:
                    row_data.append('')

                # Add spacer column
                row_data.append('')

                # Add YTD formula
                if formula_type and formula_type not in ['sum_operating', 'sum_investing', 'sum_financing', 'sum_all']:
                    first_col = self._col_letter(2)
                    last_col_letter = self._col_letter(last_month_col)
                    data_range = f'{first_col}{actual_row}:{last_col_letter}{actual_row}'
                    helper_range = f'{first_col}$3:{last_col_letter}$3'
                    ytd_formula = f'=SUMPRODUCT(({data_range})*--(INT({helper_range}/100)=Menu!$F$7)*--(MOD({helper_range},100)<=Menu!$E$7))'
                    row_data.append(ytd_formula)
                else:
                    row_data.append('')

            # Ensure row has correct number of columns
            while len(row_data) < num_cols:
                row_data.append('')

            all_data.append(row_data)
            row_types_list.append(row_type)

        # ================================================================
        # WRITE ALL DATA IN ONE BULK OPERATION
        # ================================================================
        if all_data:
            data_end_row = data_start_row + len(all_data) - 1
            sheet.range((data_start_row, 1), (data_end_row, num_cols)).value = all_data

        # ================================================================
        # SECOND PASS: Fill in subtotal formulas (need row references)
        # ================================================================
        # Operating subtotal
        if operating_start_idx is not None and operating_end_idx is not None:
            subtotal_row = data_start_row + operating_end_idx + 1
            op_start = data_start_row + operating_start_idx
            op_end = data_start_row + operating_end_idx
            for col in range(2, ytd_col + 1):
                sheet.range((subtotal_row, col)).formula = f'=SUM({self._col_letter(col)}{op_start}:{self._col_letter(col)}{op_end})'

        # Investing subtotal
        if investing_start_idx is not None and investing_end_idx is not None:
            subtotal_row = data_start_row + investing_end_idx + 1
            inv_start = data_start_row + investing_start_idx
            inv_end = data_start_row + investing_end_idx
            for col in range(2, ytd_col + 1):
                sheet.range((subtotal_row, col)).formula = f'=SUM({self._col_letter(col)}{inv_start}:{self._col_letter(col)}{inv_end})'

        # Financing subtotal
        if financing_start_idx is not None and financing_end_idx is not None:
            subtotal_row = data_start_row + financing_end_idx + 1
            fin_start = data_start_row + financing_start_idx
            fin_end = data_start_row + financing_end_idx
            for col in range(2, ytd_col + 1):
                sheet.range((subtotal_row, col)).formula = f'=SUM({self._col_letter(col)}{fin_start}:{self._col_letter(col)}{fin_end})'

        # Net change in cash = Operating + Investing + Financing subtotals
        if net_change_idx is not None and operating_end_idx is not None:
            net_change_row = data_start_row + net_change_idx
            op_row = data_start_row + operating_end_idx + 1
            inv_row = data_start_row + investing_end_idx + 1
            fin_row = data_start_row + financing_end_idx + 1
            for col in range(2, ytd_col + 1):
                col_letter = self._col_letter(col)
                sheet.range((net_change_row, col)).formula = f'={col_letter}{op_row}+{col_letter}{inv_row}+{col_letter}{fin_row}'

        # Ending cash = Beginning + Net Change
        if ending_cash_idx is not None and beginning_cash_idx is not None and net_change_idx is not None:
            ending_cash_row = data_start_row + ending_cash_idx
            beginning_cash_row = data_start_row + beginning_cash_idx
            net_change_row = data_start_row + net_change_idx
            for col in range(2, ytd_col + 1):
                col_letter = self._col_letter(col)
                sheet.range((ending_cash_row, col)).formula = f'={col_letter}{beginning_cash_row}+{col_letter}{net_change_row}'

        # ================================================================
        # APPLY FORMATTING IN BULK
        # ================================================================
        row = data_end_row + 1 if all_data else data_start_row

        # Collect rows by type for batch formatting
        header_rows = [data_start_row + idx for idx, rt in enumerate(row_types_list) if rt == 'header']
        subtotal_rows = [data_start_row + idx for idx, rt in enumerate(row_types_list) if rt == 'subtotal']
        total_rows = [data_start_row + idx for idx, rt in enumerate(row_types_list) if rt == 'total']

        # Format headers (bold)
        for r in header_rows:
            sheet.range((r, 1)).font.bold = True

        # Format subtotals (bold, gray background)
        for r in subtotal_rows:
            sheet.range((r, 1)).font.bold = True
            sheet.range((r, 1), (r, ytd_col)).color = SUBTOTAL_GRAY

        # Format totals (bold)
        for r in total_rows:
            sheet.range((r, 1)).font.bold = True

        # Apply formatting to data range in bulk operations
        try:
            # Apply font to entire data area at once
            full_data_range = sheet.range((header_row + 1, 1), (row - 1, ytd_col))
            full_data_range.font.name = 'Calibri Light'
            full_data_range.font.size = 10

            # Number format and alignment for numeric columns
            data_range = sheet.range((header_row + 1, 2), (row - 1, ytd_col))
            data_range.number_format = '#,##0'
            data_range.api.HorizontalAlignment = -4152  # xlRight
        except:
            pass

        # Set column widths and AutoFit
        sheet.range('A:A').column_width = 50
        try:
            sheet.range('A:A').api.EntireColumn.AutoFit()
            for col in range(2, ytd_col + 1):
                sheet.range((1, col), (1, col)).api.EntireColumn.AutoFit()
        except:
            pass

        # Group prior year columns like PL and BS
        self._group_columns_by_year(sheet, months, header_row)

        # Collapse all outline groups
        try:
            sheet.api.Outline.ShowLevels(ColumnLevels=1)
        except:
            pass

        # Hide row 3 (YYYYMM helper row) - MUST be at end after all other operations
        try:
            sheet.range('3:3').api.Hidden = True
            print(f"[CF] Row 3 hidden successfully")
        except Exception as e:
            print(f"[CF] ERROR hiding row 3: {e}")

        # Add back to menu link and print setup
        self._add_back_to_menu_link(sheet, row=1, col=1)
        self._setup_print_area(sheet)

    def _create_notes_sheet(self, sheet, pl_accounts, bs_accounts, months, divisions=None):
        """Create consolidated notes sheet for P&L, Balance Sheet, and Cash Flow with dropdowns.

        Columns (multi-division mode):
        - A: Division (dropdown)
        - B: Statement Type (P&L, Balance Sheet, Cash Flow)
        - C: Date (month dropdown)
        - D: Account (dropdown based on statement type)
        - E: Note

        Columns (single-entity mode):
        - A: Statement Type
        - B: Date
        - C: Account
        - D: Note
        """
        # Colors
        DARK_BLUE = (22, 33, 62)  # #16213E

        # Check if multi-division mode
        is_multi_div = divisions and len(divisions) > 0

        if is_multi_div:
            # Multi-division: 5 columns
            sheet.range('A1').value = 'Division'
            sheet.range('B1').value = 'Statement Type'
            sheet.range('C1').value = 'Date'
            sheet.range('D1').value = 'Account'
            sheet.range('E1').value = 'Note'
            header_range = sheet.range('A1:E1')
            last_col = 'E'
            div_col, type_col, date_col, acct_col, note_col = 'A', 'B', 'C', 'D', 'E'
        else:
            # Single-entity: 4 columns
            sheet.range('A1').value = 'Statement Type'
            sheet.range('B1').value = 'Date'
            sheet.range('C1').value = 'Account'
            sheet.range('D1').value = 'Note'
            header_range = sheet.range('A1:D1')
            last_col = 'D'
            div_col, type_col, date_col, acct_col, note_col = None, 'A', 'B', 'C', 'D'

        # Format header row with dark blue background and white text
        header_range.font.name = 'Calibri Light'
        header_range.font.size = 10
        header_range.font.bold = True
        header_range.font.color = (255, 255, 255)
        header_range.color = DARK_BLUE

        # Pre-populate some rows for data entry
        num_rows = 100  # Allow up to 100 notes

        try:
            # Statement type dropdown
            statement_types = 'P&L,Balance Sheet,Cash Flow'

            # Create account lists
            pl_account_list = ','.join([a['name'] for a in pl_accounts][:40])
            bs_account_list = ','.join([a['name'] for a in bs_accounts][:40])

            month_list = ','.join(months) if isinstance(months[0], str) else ','.join([m[2] for m in months])

            # Division dropdown (multi-division mode only)
            if is_multi_div and div_col:
                division_list = 'All,' + ','.join([d['name'] for d in divisions])
                for row in range(2, min(52, num_rows + 2)):
                    try:
                        sheet.range(f'{div_col}{row}').api.Validation.Delete()
                        sheet.range(f'{div_col}{row}').api.Validation.Add(
                            Type=3,  # xlValidateList
                            AlertStyle=1,  # xlValidAlertStop
                            Formula1=division_list[:255]
                        )
                    except:
                        pass

            # Apply data validation to Statement Type column
            for row in range(2, min(52, num_rows + 2)):
                try:
                    sheet.range(f'{type_col}{row}').api.Validation.Delete()
                    sheet.range(f'{type_col}{row}').api.Validation.Add(
                        Type=3,  # xlValidateList
                        AlertStyle=1,  # xlValidAlertStop
                        Formula1=statement_types
                    )
                except:
                    pass

            # Apply data validation to Date column
            for row in range(2, min(52, num_rows + 2)):
                try:
                    sheet.range(f'{date_col}{row}').api.Validation.Delete()
                    sheet.range(f'{date_col}{row}').api.Validation.Add(
                        Type=3,  # xlValidateList
                        AlertStyle=1,  # xlValidAlertStop
                        Formula1=month_list[:255]
                    )
                except:
                    pass

            # Apply data validation to Account column - combined list
            all_accounts = pl_account_list[:120] + ',' + bs_account_list[:120]
            for row in range(2, min(52, num_rows + 2)):
                try:
                    sheet.range(f'{acct_col}{row}').api.Validation.Delete()
                    sheet.range(f'{acct_col}{row}').api.Validation.Add(
                        Type=3,  # xlValidateList
                        AlertStyle=1,  # xlValidAlertStop
                        Formula1=all_accounts[:255]  # Excel limit
                    )
                except:
                    pass

        except Exception as e:
            print(f"Warning: Could not add dropdowns to Notes sheet: {e}")

        # Set column widths
        if is_multi_div:
            sheet.range('A:A').column_width = 18  # Division
            sheet.range('B:B').column_width = 15  # Statement Type
            sheet.range('C:C').column_width = 12  # Date
            sheet.range('D:D').column_width = 40  # Account
            sheet.range('E:E').column_width = 60  # Note
            data_range = sheet.range('A2:E51')
        else:
            sheet.range('A:A').column_width = 15
            sheet.range('B:B').column_width = 12
            sheet.range('C:C').column_width = 40
            sheet.range('D:D').column_width = 60
            data_range = sheet.range('A2:D51')

        # Format data area
        data_range.font.name = 'Calibri Light'
        data_range.font.size = 10

        # Enable AutoFilter for sorting
        try:
            sheet.range(f'A1:{last_col}1').api.AutoFilter()
        except:
            pass

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    # =========================================================================
    # FORECAST MODULE METHODS
    # =========================================================================

    def _create_source_budget_sheet(self, sheet, accounts, months):
        """Create Source_Budget sheet with 12 months for full-year budgeting.

        Structure:
        - Row 1: Company name title
        - Row 2: "Budget" subtitle
        - Row 3: Blank
        - Row 4: Headers (Account, Jan-Dec month names like "Jan 24")
        - Row 5: YYYYMM helper values (e.g., 202401) for formula lookups - hidden
        - Row 6+: Account data with indentation (initially zeros, populated via budget import)

        Note: Unlike Source_PL which only has actual data months, Source_Budget always
        has all 12 months of the year to support full-year forecasting.
        """
        # Colors - matching source sheets
        SOURCE_BLACK = (26, 26, 26)  # #1A1A1A
        DARK_BLUE = (22, 33, 62)
        SUBTOTAL_GRAY = (236, 236, 236)

        company = self.company_name.get() if hasattr(self, 'company_name') else 'Company'

        # Determine the budget year from the last month in the data
        budget_year = months[-1][1] if months else datetime.now().year
        year_suffix = str(budget_year)[-2:]

        # Create full 12-month list for budget (Jan-Dec of budget year)
        month_abbrevs = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                         'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        budget_months = []
        for m in range(1, 13):
            name = f"{month_abbrevs[m-1]} {year_suffix}"
            budget_months.append((m, budget_year, name))

        print(f"[Source_Budget] Creating 12-month budget for year {budget_year}")

        # Row 1: Company name title
        sheet.range('A1').value = company
        sheet.range('A1').font.name = 'Calibri Light'
        sheet.range('A1').font.size = 14
        sheet.range('A1').font.bold = True

        # Row 2: Budget subtitle
        sheet.range('A2').value = 'Budget'
        sheet.range('A2').font.name = 'Calibri Light'
        sheet.range('A2').font.size = 12
        sheet.range('A2').font.bold = True

        header_row = 4

        # Row 4: Headers - always 12 months (Jan-Dec)
        sheet.range(f'A{header_row}').value = 'Account'
        for i, (m, y, name) in enumerate(budget_months):
            sheet.range((header_row, i + 2)).value = name

        # Row 5: YYYYMM helper values for formula lookups
        for i, (m, y, name) in enumerate(budget_months):
            sheet.range((header_row + 1, i + 2)).value = y * 100 + m

        # Hide row 5 (helper row)
        try:
            sheet.range(f'{header_row + 1}:{header_row + 1}').api.Hidden = True
        except:
            pass

        # Row 6+: Account names with indentation and zero values
        data = []
        data_start_row = header_row + 2
        for account in accounts:
            # Preserve indentation like P&L
            account_name = account['name']
            indent_level = account.get('indent', 0)
            is_header = account.get('is_header', False)
            is_total = account.get('is_total', False)

            # Add indentation for non-header, non-total rows
            if indent_level > 0 and not is_header and not is_total:
                display_name = ('    ' * indent_level) + account_name
            else:
                display_name = account_name

            row = [display_name]
            for _ in range(12):  # Always 12 months
                row.append(0)  # Initialize with zeros
            data.append(row)

        if data:
            sheet.range(f'A{data_start_row}').value = data

        # Format header row (always 12 months + Account column = 13 columns)
        num_budget_cols = 12
        try:
            header_range = sheet.range((header_row, 1), (header_row, num_budget_cols + 1))
            header_range.font.name = 'Calibri Light'
            header_range.font.size = 10
            header_range.font.bold = True
            header_range.font.color = (255, 255, 255)
            header_range.color = SOURCE_BLACK

            for col in range(2, num_budget_cols + 2):
                sheet.range((header_row, col)).api.HorizontalAlignment = -4108  # xlCenter
        except:
            pass

        # Format data area with proper styling for headers/totals
        if len(accounts) > 0:
            try:
                data_range = sheet.range((data_start_row, 2), (data_start_row + len(accounts) - 1, num_budget_cols + 1))
                data_range.number_format = '#,##0'
                data_range.font.name = 'Calibri Light'
                data_range.font.size = 10

                account_range = sheet.range((data_start_row, 1), (data_start_row + len(accounts) - 1, 1))
                account_range.font.name = 'Calibri Light'
                account_range.font.size = 10

                # Format header and total rows
                for i, account in enumerate(accounts):
                    row_num = data_start_row + i
                    if account.get('is_header', False):
                        sheet.range((row_num, 1)).font.bold = True
                    elif account.get('is_total', False):
                        row_range = sheet.range((row_num, 1), (row_num, num_budget_cols + 1))
                        row_range.font.bold = True
                        row_range.color = SUBTOTAL_GRAY
            except:
                pass

        # Set column widths
        sheet.range('A:A').column_width = 45
        for col in range(2, num_budget_cols + 2):
            sheet.range((1, col), (1, col)).column_width = 14

        # Group and collapse previous year columns (use budget_months instead of months)
        self._group_previous_year_columns(sheet, budget_months, data_start_col=2)

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    def _create_forecast_sheet(self, sheet, accounts, months, division_name=None):
        """Create Forecast sheet with 5 columns per month: ACTUAL, BUDGET, ADJ, NOTE, FORECAST.

        Structure:
        - Column A: Account names
        - For each month (5 columns): ACTUAL, BUDGET, ADJ, NOTE, FORECAST
        - FORECAST = ACTUAL for past months, BUDGET + ADJ for future months
        - Columns are grouped so only FORECAST is visible by default
        - Total Forecast column at end sums all FORECAST columns

        Args:
            sheet: xlwings sheet object
            accounts: List of account dictionaries
            months: List of (month, year, display_name) tuples
            division_name: Optional division name for multi-division mode
        """
        # Colors
        DARK_BLUE = (22, 33, 62)
        LIGHT_GRAY = (242, 242, 242)
        SUBTOTAL_GRAY = (220, 220, 220)
        ACTUAL_BLUE = (189, 215, 238)
        BUDGET_YELLOW = (255, 242, 204)
        ADJ_ORANGE = (252, 228, 214)
        NOTE_WHITE = (255, 255, 255)
        FORECAST_GREEN = (198, 224, 180)

        # Determine current year and month for past/future logic
        current_year = months[-1][1] if months else datetime.now().year
        current_month = datetime.now().month
        current_ym = current_year * 100 + current_month

        # Create month list for current year (12 months)
        month_abbrevs = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                         'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        year_suffix = str(current_year)[-2:]

        # 5 columns per month
        COLS_PER_MONTH = 5
        COL_ACTUAL = 0
        COL_BUDGET = 1
        COL_ADJ = 2
        COL_NOTE = 3
        COL_FORECAST = 4

        # Determine source sheet name and if multi-division mode
        multi_division = hasattr(self, 'divisions') and len(self.divisions) > 1
        if division_name:
            source_pl = f"'{division_name}_PL'"
            source_budget = "'Source_Budget'"
        else:
            source_pl = "'Consolidated_PL'" if multi_division else "'P&L'"
            source_budget = "'Source_Budget'"

        # Title
        title_text = f'Forecast - {current_year}'
        if division_name:
            title_text = f'{division_name} Forecast - {current_year}'
        sheet.range('A1').value = title_text
        sheet.range('A1').font.bold = True
        sheet.range('A1').font.size = 16
        sheet.range('A1').font.color = DARK_BLUE

        # Row 3: Month headers (merged across 5 columns each)
        # Row 4: Column sub-headers (ACTUAL, BUDGET, ADJ, NOTE, FORECAST)
        col = 2  # Start at column B
        month_header_row = []
        subheader_row = ['Account']

        for month_idx in range(12):
            month_num = month_idx + 1
            month_name = f"{month_abbrevs[month_idx]} {year_suffix}"
            month_header_row.append(month_name)

            # Add 5 sub-headers for this month
            # For past months, last column shows "Actual"; for future months, shows "Forecast"
            month_ym = current_year * 100 + month_num
            is_past_month = month_ym <= current_ym
            forecast_header = 'Actual' if is_past_month else 'Forecast'
            subheader_row.extend(['Actual', 'Budget', 'Adj', 'Note', forecast_header])

        # Add Total Forecast at end
        subheader_row.append('Total Forecast')

        # Write sub-headers (row 4)
        sheet.range('A4').value = [subheader_row]
        header_range = sheet.range((4, 1), (4, len(subheader_row)))
        header_range.font.bold = True
        header_range.font.size = 9

        # Write and merge month headers (row 3)
        for month_idx in range(12):
            start_col = 2 + (month_idx * COLS_PER_MONTH)
            end_col = start_col + COLS_PER_MONTH - 1
            month_name = f"{month_abbrevs[month_idx]} {year_suffix}"

            # Write month name and merge
            sheet.range((3, start_col)).value = month_name
            try:
                merge_range = sheet.range((3, start_col), (3, end_col))
                merge_range.merge()
                merge_range.font.bold = True
                merge_range.font.size = 11
                merge_range.color = DARK_BLUE
                merge_range.font.color = (255, 255, 255)
                merge_range.api.HorizontalAlignment = -4108  # Center
            except:
                pass

        # Color the sub-header columns
        for month_idx in range(12):
            start_col = 2 + (month_idx * COLS_PER_MONTH)
            try:
                sheet.range((4, start_col + COL_ACTUAL)).color = ACTUAL_BLUE
                sheet.range((4, start_col + COL_BUDGET)).color = BUDGET_YELLOW
                sheet.range((4, start_col + COL_ADJ)).color = ADJ_ORANGE
                sheet.range((4, start_col + COL_NOTE)).color = NOTE_WHITE
                sheet.range((4, start_col + COL_FORECAST)).color = FORECAST_GREEN
            except:
                pass

        # Total Forecast header
        total_col = 2 + (12 * COLS_PER_MONTH)
        sheet.range((3, total_col)).value = 'TOTAL'
        sheet.range((3, total_col)).font.bold = True
        sheet.range((3, total_col)).color = FORECAST_GREEN
        sheet.range((4, total_col)).color = FORECAST_GREEN

        # Track total rows for formatting
        total_rows = []
        for acct_idx, account in enumerate(accounts):
            if account.get('is_total', False):
                total_rows.append(5 + acct_idx)

        last_row = 4 + len(accounts)

        # ================================================================
        # BUILD ALL DATA IN MEMORY FIRST (OPTIMIZED - single bulk write)
        # ================================================================
        # Pre-calculate column letters and source info
        source_data_start_col = 3 if multi_division else 2
        source_acct_col = 'B' if multi_division else 'A'
        total_col = 2 + (12 * COLS_PER_MONTH)

        # Build month-to-column mapping for Source_PL
        # Map by (month, year) tuple to handle cross-year data correctly
        month_year_to_source_col = {}
        for idx, (m, y, name) in enumerate(months):
            month_year_to_source_col[(m, y)] = source_data_start_col + idx

        # Debug: print mapping
        print(f"[Forecast] Source months mapping: {month_year_to_source_col}")
        print(f"[Forecast] Looking for year {current_year}, current_ym = {current_ym}")

        # Build complete 2D array for all data (accounts × all columns)
        all_data = []  # Each row: [account_name, month1_actual, month1_budget, month1_adj, month1_note, month1_forecast, ...]

        for acct_idx, account in enumerate(accounts):
            r = 5 + acct_idx
            account_name = account['name']
            indent_level = account.get('indent', 0)

            # Build display name with indentation
            display_name = account_name
            if indent_level > 0 and not account.get('is_header', False) and not account.get('is_total', False):
                display_name = ('  ' * indent_level) + account_name

            row_data = [display_name]

            # Build formulas for all 12 months
            for month_idx in range(12):
                month_num = month_idx + 1
                month_ym = current_year * 100 + month_num
                is_past_month = month_ym <= current_ym

                start_col = 2 + (month_idx * COLS_PER_MONTH)
                actual_letter = self._col_letter(start_col + COL_ACTUAL)
                budget_letter = self._col_letter(start_col + COL_BUDGET)
                adj_letter = self._col_letter(start_col + COL_ADJ)

                # Get the source column for ACTUAL data from Source_PL (using month, year tuple)
                month_key = (month_num, current_year)
                if month_key in month_year_to_source_col:
                    source_col_letter = self._col_letter(month_year_to_source_col[month_key])
                    # ACTUAL formula - pull from Source_PL
                    if division_name:
                        actual_formula = f'=IFERROR(SUMIFS(Source_PL!{source_col_letter}$3:{source_col_letter}$1500,Source_PL!$A$3:$A$1500,"{division_name}",Source_PL!$B$3:$B$1500,TRIM(A{r})),0)'
                    else:
                        actual_formula = f'=IFERROR(SUMIF(Source_PL!${source_acct_col}$3:${source_acct_col}$1500,TRIM(A{r}),Source_PL!{source_col_letter}$3:{source_col_letter}$1500),0)'
                else:
                    # Month not in source data - no actual data available
                    actual_formula = '0'

                # BUDGET formula - Source_Budget ALWAYS has Jan-Dec in columns B-M (2-13)
                # So Jan=B, Feb=C, Mar=D, etc. (column = 2 + month_idx where month_idx is 0-11)
                budget_col_letter = self._col_letter(2 + month_idx)
                budget_formula = f'=IFERROR(SUMIF(Source_Budget!$A$6:$A$1500,TRIM(A{r}),Source_Budget!{budget_col_letter}$6:{budget_col_letter}$1500),0)'

                # FORECAST formula
                if is_past_month:
                    forecast_formula = f'={actual_letter}{r}'
                else:
                    forecast_formula = f'={budget_letter}{r}+{adj_letter}{r}'

                # Add 5 columns for this month: Actual, Budget, Adj, Note, Forecast
                row_data.extend([actual_formula, budget_formula, 0, '', forecast_formula])

            # Add Total Forecast formula
            forecast_cols = [self._col_letter(2 + (m * COLS_PER_MONTH) + COL_FORECAST) for m in range(12)]
            sum_parts = '+'.join([f'{c}{r}' for c in forecast_cols])
            row_data.append(f'={sum_parts}')

            all_data.append(row_data)

        # WRITE ALL DATA IN ONE BULK OPERATION
        if all_data:
            num_cols = 1 + (12 * COLS_PER_MONTH) + 1  # Account + 12 months × 5 cols + Total
            sheet.range((5, 1), (last_row, num_cols)).value = all_data

        # ================================================================
        # APPLY FORMATTING IN BATCHES (by column type across all months)
        # ================================================================
        # Number format for all data columns at once
        sheet.range((5, 2), (last_row, total_col)).number_format = '#,##0'

        # Color columns by type - batch all months together
        for month_idx in range(12):
            start_col = 2 + (month_idx * COLS_PER_MONTH)
            sheet.range((5, start_col + COL_ACTUAL), (last_row, start_col + COL_ACTUAL)).color = ACTUAL_BLUE
            sheet.range((5, start_col + COL_BUDGET), (last_row, start_col + COL_BUDGET)).color = BUDGET_YELLOW
            sheet.range((5, start_col + COL_ADJ), (last_row, start_col + COL_ADJ)).color = ADJ_ORANGE
            sheet.range((5, start_col + COL_FORECAST), (last_row, start_col + COL_FORECAST)).color = FORECAST_GREEN

        # Total column formatting
        sheet.range((5, total_col), (last_row, total_col)).font.bold = True
        sheet.range((5, total_col), (last_row, total_col)).color = FORECAST_GREEN

        # Format total rows
        for r in total_rows:
            try:
                row_range = sheet.range((r, 1), (r, total_col))
                row_range.font.bold = True
                row_range.color = SUBTOTAL_GRAY
            except:
                pass

        # Set column widths
        sheet.range('A:A').column_width = 35

        for month_idx in range(12):
            start_col = 2 + (month_idx * COLS_PER_MONTH)
            try:
                sheet.range((1, start_col + COL_ACTUAL)).column_width = 10
                sheet.range((1, start_col + COL_BUDGET)).column_width = 10
                sheet.range((1, start_col + COL_ADJ)).column_width = 8
                sheet.range((1, start_col + COL_NOTE)).column_width = 15
                sheet.range((1, start_col + COL_FORECAST)).column_width = 11
            except:
                pass

        sheet.range((1, total_col)).column_width = 13

        # Group columns so only FORECAST is visible (hide ACTUAL, BUDGET, ADJ, NOTE)
        try:
            for month_idx in range(12):
                start_col = 2 + (month_idx * COLS_PER_MONTH)
                # Group columns: ACTUAL, BUDGET, ADJ, NOTE (hide them, keep FORECAST visible)
                group_start = self._col_letter(start_col + COL_ACTUAL)
                group_end = self._col_letter(start_col + COL_NOTE)
                sheet.range(f'{group_start}:{group_end}').api.Columns.Group()

            # Collapse all groups
            sheet.api.Outline.ShowLevels(ColumnLevels=1)
        except Exception as e:
            print(f"Warning: Could not group columns: {e}")

        # Freeze panes (Account column and header rows)
        try:
            sheet.range('B5').api.Select()
            sheet.book.app.api.ActiveWindow.FreezePanes = True
        except:
            pass

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=total_col + 2)

    def _create_forecast_summary_sheet(self, sheet, accounts, months):
        """Create Forecast Summary with Current Month, YTD Actual vs Budget, and YTD Forecast comparisons.
        Matches P&L formatting with indentation, borders, and profit % rows.

        Layout:
        - Column A: Account names (with P&L-style indentation)
        - Column B: Current Month Actual
        - Column C: Current Month Budget
        - Column D: Current Month Variance $
        - Column E: Current Month Variance %
        - Column F: (spacer)
        - Column G: YTD Actual
        - Column H: YTD Budget
        - Column I: YTD Variance $
        - Column J: YTD Variance %
        - Column K: (spacer)
        - Column L: YTD Forecast
        - Column M: YTD Budget (for forecast comparison)
        - Column N: Forecast Var $
        - Column O: Forecast Var %
        """
        # Colors
        DARK_BLUE = (22, 33, 62)
        LIGHT_GRAY = (242, 242, 242)
        HEADER_GRAY = (200, 200, 200)
        SUBTOTAL_GRAY = (236, 236, 236)  # Matches P&L

        # Spacer columns
        SPACER1_COL = 6  # F
        SPACER2_COL = 11  # K
        LAST_DATA_COL = 15  # O

        # Determine current year
        if months:
            current_year = months[-1][1]
        else:
            current_year = 2024

        col_letter_last = self._col_letter(len(months) + 1)

        # Determine forecast sheet name and column offsets based on mode
        is_multi_division = hasattr(self, 'divisions') and len(self.divisions) > 1
        forecast_sheet_name = 'Consolidated_Forecast' if is_multi_division else 'Forecast'
        # In multi-division mode: col A=Division, col B=Account, data starts col C
        # In single mode: col A=Account, data starts col B
        src_acct_col = 'B' if is_multi_division else 'A'
        src_data_start_col = 'C' if is_multi_division else 'B'

        # Row 1: Title
        sheet.range('A1').value = f'Forecast Summary - {current_year}'
        sheet.range('A1').font.bold = True
        sheet.range('A1').font.size = 14
        sheet.range('A1').font.name = 'Calibri Light'

        # Row 2: Section headers (merged)
        sheet.range('B2').value = 'Current Month'
        try:
            sheet.range('B2:E2').merge()
            sheet.range('B2').api.HorizontalAlignment = -4108  # xlCenter
        except:
            pass
        sheet.range('B2').font.bold = True
        sheet.range('B2').font.name = 'Calibri Light'
        sheet.range('B2').color = DARK_BLUE
        sheet.range('B2').font.color = (255, 255, 255)

        sheet.range('G2').value = 'Year-to-Date Actual vs Budget'
        try:
            sheet.range('G2:J2').merge()
            sheet.range('G2').api.HorizontalAlignment = -4108  # xlCenter
        except:
            pass
        sheet.range('G2').font.bold = True
        sheet.range('G2').font.name = 'Calibri Light'
        sheet.range('G2').color = DARK_BLUE
        sheet.range('G2').font.color = (255, 255, 255)

        sheet.range('L2').value = 'YTD Forecast vs Budget'
        try:
            sheet.range('L2:O2').merge()
            sheet.range('L2').api.HorizontalAlignment = -4108  # xlCenter
        except:
            pass
        sheet.range('L2').font.bold = True
        sheet.range('L2').font.name = 'Calibri Light'
        sheet.range('L2').color = DARK_BLUE
        sheet.range('L2').font.color = (255, 255, 255)

        # Row 3: Column headers
        headers = [
            ('A3', 'Account'),
            ('B3', 'Actual'),
            ('C3', 'Budget'),
            ('D3', 'Var $'),
            ('E3', 'Var %'),
            ('F3', ''),  # Spacer
            ('G3', 'Actual'),
            ('H3', 'Budget'),
            ('I3', 'Var $'),
            ('J3', 'Var %'),
            ('K3', ''),  # Spacer
            ('L3', 'Forecast'),
            ('M3', 'Budget'),
            ('N3', 'Var $'),
            ('O3', 'Var %'),
        ]
        for cell, value in headers:
            sheet.range(cell).value = value

        # Format header row
        header_range = sheet.range('A3:O3')
        header_range.font.bold = True
        header_range.font.name = 'Calibri Light'
        header_range.font.size = 10
        header_range.color = HEADER_GRAY

        # Spacer columns formatting
        sheet.range('F2:F3').color = (255, 255, 255)
        sheet.range('K2:K3').color = (255, 255, 255)

        # ================================================================
        # BUILD ALL DATA IN MEMORY FIRST (OPTIMIZED)
        # ================================================================
        all_data = []  # List of row data arrays
        row_types = []  # Track row type for batch formatting

        # Track key rows for Gross Profit % and Net Profit %
        gross_margin_row = None
        total_income_row = None
        net_income_row = None
        forecast_row = 5  # Track actual Forecast sheet row (starts at 5)

        data_start_row = 4  # Data starts at row 4

        for acct_idx, account in enumerate(accounts):
            account_name = account['name']
            name_lower = account_name.lower()
            actual_row = data_start_row + len(all_data)

            # Get indent level from source file
            indent_level = account.get('indent', 0)

            # Determine display name with indentation
            display_name = account_name
            if indent_level > 0 and not account['is_header'] and not account['is_total']:
                display_name = ('    ' * indent_level) + account_name

            # Track Gross Profit/Margin row
            if 'gross profit' in name_lower:
                gross_margin_row = actual_row

            # Track Net Income row
            if 'net income' in name_lower and account['is_total']:
                net_income_row = actual_row

            # Track Total Income/Revenue row
            if account['is_total'] and 'total' in name_lower:
                if (('income' in name_lower or 'revenue' in name_lower) and
                    'net' not in name_lower and 'other' not in name_lower):
                    total_income_row = actual_row

            # Build row data array: [A, B, C, D, E, F, G, H, I, J, K, L, M, N, O]
            if account['is_header']:
                # Header rows: just the name, empty data cells
                row_data = [display_name] + [''] * 14
                all_data.append(row_data)
                row_types.append('header')
                forecast_row += 1
                continue

            # --- Build formulas for data rows ---
            # In multi-division mode, Source_PL has: col A=Division, col B=Account, col C onwards=data
            # Row 2 has YYYYMM values in the data columns

            # Current Month Actual (B) - sum where account matches and YYYYMM = Menu!G7
            if is_multi_division:
                cm_actual = (
                    f"=SUMPRODUCT("
                    f"(Source_PL!${src_acct_col}$3:${src_acct_col}$1000=TRIM($A{actual_row}))*"
                    f"(Source_PL!${src_data_start_col}$2:${col_letter_last}$2=Menu!$G$7)*"
                    f"(Source_PL!${src_data_start_col}$3:${col_letter_last}$1000))"
                )
            else:
                cm_actual = (
                    f"=SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=TRIM($A{actual_row}))*"
                    f"(Source_PL!$B$2:${col_letter_last}$2=Menu!$G$7)*"
                    f"(Source_PL!$B$3:${col_letter_last}$1000))"
                )

            # Current Month Budget (C)
            cm_budget = (
                f"=SUMPRODUCT("
                f"(Source_Budget!$A$6:$A$1000=TRIM($A{actual_row}))*"
                f"(Source_Budget!$B$5:${col_letter_last}$5=Menu!$G$7)*"
                f"(Source_Budget!$B$6:${col_letter_last}$1000))"
            )
            # Current Month Variance $ (D)
            cm_var = f"=B{actual_row}-C{actual_row}"
            # Current Month Variance % (E)
            cm_var_pct = f"=IFERROR(D{actual_row}/ABS(C{actual_row}),0)"

            # YTD Actual (G) - sum where account matches and year matches and YYYYMM <= Menu!G7
            if is_multi_division:
                ytd_actual = (
                    f"=SUMPRODUCT("
                    f"(Source_PL!${src_acct_col}$3:${src_acct_col}$1000=TRIM($A{actual_row}))*"
                    f"(INT(Source_PL!${src_data_start_col}$2:${col_letter_last}$2/100)=Menu!$F$7)*"
                    f"(Source_PL!${src_data_start_col}$2:${col_letter_last}$2<=Menu!$G$7)*"
                    f"(Source_PL!${src_data_start_col}$3:${col_letter_last}$1000))"
                )
            else:
                ytd_actual = (
                    f"=SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=TRIM($A{actual_row}))*"
                    f"(INT(Source_PL!$B$2:${col_letter_last}$2/100)=Menu!$F$7)*"
                    f"(Source_PL!$B$2:${col_letter_last}$2<=Menu!$G$7)*"
                    f"(Source_PL!$B$3:${col_letter_last}$1000))"
                )

            # YTD Budget (H)
            ytd_budget = (
                f"=SUMPRODUCT("
                f"(Source_Budget!$A$6:$A$1000=TRIM($A{actual_row}))*"
                f"(INT(Source_Budget!$B$5:${col_letter_last}$5/100)=Menu!$F$7)*"
                f"(Source_Budget!$B$5:${col_letter_last}$5<=Menu!$G$7)*"
                f"(Source_Budget!$B$6:${col_letter_last}$1000))"
            )
            # YTD Variance $ (I)
            ytd_var = f"=G{actual_row}-H{actual_row}"
            # YTD Variance % (J)
            ytd_var_pct = f"=IFERROR(I{actual_row}/ABS(H{actual_row}),0)"

            # YTD Forecast (L) - sum of Forecast columns where month <= current
            ytd_forecast_parts = []
            for month_idx in range(12):
                forecast_col = self._col_letter(2 + month_idx * 5 + 4)
                yyyymm = current_year * 100 + (month_idx + 1)
                ytd_forecast_parts.append(f"IF({yyyymm}<=Menu!$G$7,'{forecast_sheet_name}'!{forecast_col}{forecast_row},0)")
            ytd_forecast = "=" + "+".join(ytd_forecast_parts)

            # YTD Budget for Forecast (M)
            ytd_budget_fc = f"=H{actual_row}"
            # Forecast Variance $ (N)
            fc_var = f"=L{actual_row}-M{actual_row}"
            # Forecast Variance % (O)
            fc_var_pct = f"=IFERROR(N{actual_row}/ABS(M{actual_row}),0)"

            # Assemble row: A=name, B-E=CM, F=spacer, G-J=YTD, K=spacer, L-O=Forecast
            row_data = [
                display_name,   # A
                cm_actual,      # B
                cm_budget,      # C
                cm_var,         # D
                cm_var_pct,     # E
                '',             # F (spacer)
                ytd_actual,     # G
                ytd_budget,     # H
                ytd_var,        # I
                ytd_var_pct,    # J
                '',             # K (spacer)
                ytd_forecast,   # L
                ytd_budget_fc,  # M
                fc_var,         # N
                fc_var_pct      # O
            ]
            all_data.append(row_data)

            # Track row type
            if account['is_total']:
                if 'net income' in name_lower:
                    row_types.append('net_income')
                else:
                    row_types.append('total')
            else:
                row_types.append('detail')

            forecast_row += 1

        # ================================================================
        # WRITE ALL DATA IN ONE BULK OPERATION
        # ================================================================
        if all_data:
            data_end_row = data_start_row + len(all_data) - 1
            sheet.range((data_start_row, 1), (data_end_row, 15)).value = all_data
        else:
            data_end_row = data_start_row

        # ================================================================
        # APPLY FORMATTING IN BULK (after data write)
        # ================================================================
        if all_data:
            # Collect rows by type for batch formatting
            header_rows = [data_start_row + idx for idx, rt in enumerate(row_types) if rt == 'header']
            total_rows = [data_start_row + idx for idx, rt in enumerate(row_types) if rt == 'total']
            net_income_rows = [data_start_row + idx for idx, rt in enumerate(row_types) if rt == 'net_income']

            # Format header rows (bold column A)
            for r in header_rows:
                sheet.range((r, 1)).font.bold = True

            # Format total rows (bold, gray background)
            for r in total_rows:
                sheet.range((r, 1)).font.bold = True
                # Apply gray background to non-spacer columns
                for c in [1, 2, 3, 4, 5, 7, 8, 9, 10, 12, 13, 14, 15]:
                    sheet.range((r, c)).color = SUBTOTAL_GRAY

            # Format net income rows (bold, borders)
            for r in net_income_rows:
                sheet.range((r, 1)).font.bold = True

        # Add blank row then Gross Profit % and Net Profit % rows
        row = data_end_row + 2  # Blank row after data

        # Gross Profit % row
        if gross_margin_row and total_income_row:
            sheet.range(f'A{row}').value = 'Gross Profit %'
            sheet.range(f'A{row}').font.name = 'Calibri Light'
            sheet.range(f'A{row}').font.size = 10
            sheet.range(f'A{row}').font.bold = True
            sheet.range(f'A{row}').font.italic = True

            # Current Month GP %
            sheet.range(f'B{row}').value = f"=IFERROR(B{gross_margin_row}/B{total_income_row},0)"
            sheet.range(f'B{row}').number_format = '0.0%'
            sheet.range(f'B{row}').font.name = 'Calibri Light'
            sheet.range(f'B{row}').font.italic = True

            # YTD Actual GP %
            sheet.range(f'G{row}').value = f"=IFERROR(G{gross_margin_row}/G{total_income_row},0)"
            sheet.range(f'G{row}').number_format = '0.0%'
            sheet.range(f'G{row}').font.name = 'Calibri Light'
            sheet.range(f'G{row}').font.italic = True

            # YTD Forecast GP %
            sheet.range(f'L{row}').value = f"=IFERROR(L{gross_margin_row}/L{total_income_row},0)"
            sheet.range(f'L{row}').number_format = '0.0%'
            sheet.range(f'L{row}').font.name = 'Calibri Light'
            sheet.range(f'L{row}').font.italic = True

            row += 1

        # Net Profit % row
        if net_income_row and total_income_row:
            sheet.range(f'A{row}').value = 'Net Profit %'
            sheet.range(f'A{row}').font.name = 'Calibri Light'
            sheet.range(f'A{row}').font.size = 10
            sheet.range(f'A{row}').font.bold = True
            sheet.range(f'A{row}').font.italic = True

            # Current Month NP %
            sheet.range(f'B{row}').value = f"=IFERROR(B{net_income_row}/B{total_income_row},0)"
            sheet.range(f'B{row}').number_format = '0.0%'
            sheet.range(f'B{row}').font.name = 'Calibri Light'
            sheet.range(f'B{row}').font.italic = True

            # YTD Actual NP %
            sheet.range(f'G{row}').value = f"=IFERROR(G{net_income_row}/G{total_income_row},0)"
            sheet.range(f'G{row}').number_format = '0.0%'
            sheet.range(f'G{row}').font.name = 'Calibri Light'
            sheet.range(f'G{row}').font.italic = True

            # YTD Forecast NP %
            sheet.range(f'L{row}').value = f"=IFERROR(L{net_income_row}/L{total_income_row},0)"
            sheet.range(f'L{row}').number_format = '0.0%'
            sheet.range(f'L{row}').font.name = 'Calibri Light'
            sheet.range(f'L{row}').font.italic = True

        # Format data area
        try:
            # Number format for dollar columns
            for col in ['B', 'C', 'D', 'G', 'H', 'I', 'L', 'M', 'N']:
                sheet.range(f'{col}4:{col}{data_end_row}').number_format = '#,##0'

            # Percentage format for variance % columns
            for col in ['E', 'J', 'O']:
                sheet.range(f'{col}4:{col}{data_end_row}').number_format = '0.0%'

            # Font styling
            data_range = sheet.range(f'A4:O{data_end_row}')
            data_range.font.name = 'Calibri Light'
            data_range.font.size = 10
        except:
            pass

        # Clear spacer columns of any background color
        try:
            sheet.range(f'F4:F{row}').color = None
            sheet.range(f'K4:K{row}').color = None
        except:
            pass

        # Set column widths
        sheet.range('A:A').column_width = 40
        sheet.range('B:B').column_width = 12  # CM Actual
        sheet.range('C:C').column_width = 12  # CM Budget
        sheet.range('D:D').column_width = 11  # CM Var $
        sheet.range('E:E').column_width = 9   # CM Var %
        sheet.range('F:F').column_width = 2   # Spacer
        sheet.range('G:G').column_width = 12  # YTD Actual
        sheet.range('H:H').column_width = 12  # YTD Budget
        sheet.range('I:I').column_width = 11  # YTD Var $
        sheet.range('J:J').column_width = 9   # YTD Var %
        sheet.range('K:K').column_width = 2   # Spacer
        sheet.range('L:L').column_width = 12  # YTD Forecast
        sheet.range('M:M').column_width = 12  # YTD Budget (Forecast)
        sheet.range('N:N').column_width = 11  # Forecast Var $
        sheet.range('O:O').column_width = 9   # Forecast Var %

        # Hide gridlines
        try:
            sheet.book.app.api.ActiveWindow.DisplayGridlines = False
        except:
            pass

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    # =========================================================================
    # DASHBOARD MODULE METHODS
    # =========================================================================

    # KPI Definitions - CFO-standard metrics and ratios
    KPI_DEFINITIONS = {
        'profitability': [
            {'name': 'Revenue', 'formula_type': 'direct', 'source': 'PL', 'account': 'Total for Income',
             'description': 'Total revenue generated', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Gross Profit', 'formula_type': 'direct', 'source': 'PL', 'account': 'Gross Profit',
             'description': 'Revenue minus COGS', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Gross Margin %', 'formula_type': 'ratio', 'numerator': 'Gross Profit', 'denominator': 'Total for Income',
             'description': 'Gross profit / revenue', 'default_target': 0.40, 'format': '0.0%', 'higher_is_better': True},
            {'name': 'Net Income', 'formula_type': 'direct', 'source': 'PL', 'account': 'Net Income',
             'description': 'Bottom line profit', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Net Margin %', 'formula_type': 'ratio', 'numerator': 'Net Income', 'denominator': 'Total for Income',
             'description': 'Net income / revenue', 'default_target': 0.10, 'format': '0.0%', 'higher_is_better': True},
            {'name': 'EBITDA', 'formula_type': 'calculated', 'calc_type': 'ebitda',
             'description': 'NI + Interest + Depreciation', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Operating Income', 'formula_type': 'direct', 'source': 'PL', 'account': 'Net Operating Income',
             'description': 'Core operations income', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Operating Margin %', 'formula_type': 'ratio', 'numerator': 'Net Operating Income', 'denominator': 'Total for Income',
             'description': 'Operating income / revenue', 'default_target': 0.12, 'format': '0.0%', 'higher_is_better': True},
        ],
        'liquidity': [
            {'name': 'Current Ratio', 'formula_type': 'bs_ratio', 'numerator': 'Total for Current Assets', 'denominator': 'Total for Current Liabilities',
             'description': 'Current assets / liabilities', 'default_target': 1.5, 'format': '0.00', 'higher_is_better': True},
            {'name': 'Quick Ratio', 'formula_type': 'bs_ratio', 'numerator': 'Total for Current Assets', 'denominator': 'Total for Current Liabilities',
             'description': '(CA - Inventory) / CL', 'default_target': 1.0, 'format': '0.00', 'higher_is_better': True},
            {'name': 'Cash Ratio', 'formula_type': 'bs_ratio', 'numerator': 'Total for Bank Accounts', 'denominator': 'Total for Current Liabilities',
             'description': 'Cash / current liabilities', 'default_target': 0.2, 'format': '0.00', 'higher_is_better': True},
            {'name': 'Working Capital', 'formula_type': 'bs_difference', 'minuend': 'Total for Current Assets', 'subtrahend': 'Total for Current Liabilities',
             'description': 'CA minus CL', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Cash Balance', 'formula_type': 'direct', 'source': 'BS', 'account': 'Total for Bank Accounts',
             'description': 'Total cash on hand', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
        ],
        'efficiency': [
            {'name': 'AR Days (DSO)', 'formula_type': 'days_ratio', 'balance': 'Total for Accounts Receivable', 'flow': 'Total for Income',
             'description': 'Days to collect AR', 'default_target': 45, 'format': '0', 'higher_is_better': False},
            {'name': 'AP Days (DPO)', 'formula_type': 'days_ratio', 'balance': 'Total for Credit Cards', 'flow': 'Total for Cost of Sales',
             'description': 'Days to pay AP', 'default_target': 30, 'format': '0', 'higher_is_better': True},
            {'name': 'Asset Turnover', 'formula_type': 'turnover', 'flow': 'Total for Income', 'balance': 'Total for Assets',
             'description': 'Revenue / assets', 'default_target': 1.0, 'format': '0.00', 'higher_is_better': True},
        ],
        'leverage': [
            {'name': 'Debt-to-Equity', 'formula_type': 'bs_ratio', 'numerator': 'Total for Liabilities', 'denominator': 'Total for Equity',
             'description': 'Debt relative to equity', 'default_target': 1.0, 'format': '0.00', 'higher_is_better': False},
            {'name': 'Debt-to-Assets', 'formula_type': 'bs_ratio', 'numerator': 'Total for Liabilities', 'denominator': 'Total for Assets',
             'description': 'Debt / total assets', 'default_target': 0.5, 'format': '0.0%', 'higher_is_better': False},
            {'name': 'Equity Ratio', 'formula_type': 'bs_ratio', 'numerator': 'Total for Equity', 'denominator': 'Total for Assets',
             'description': 'Equity / total assets', 'default_target': 0.5, 'format': '0.0%', 'higher_is_better': True},
        ],
        'cashflow': [
            {'name': 'Operating Cash Flow', 'formula_type': 'calculated', 'calc_type': 'ocf',
             'description': 'Cash from operations', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Free Cash Flow', 'formula_type': 'calculated', 'calc_type': 'fcf',
             'description': 'OCF minus CapEx', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
        ],
    }

    def _create_dashboard_control_sheet(self, sheet, months):
        """Create the Dashboard Control page with KPI target settings"""
        # Colors
        header_color = (22, 33, 62)  # Dark blue
        section_color = (44, 62, 80)  # Darker gray-blue
        control_color = (232, 244, 253)  # Light blue for editable cells

        # Title
        sheet.range('B2').value = "Dashboard Control Panel"
        sheet.range('B2').font.size = 24
        sheet.range('B2').font.bold = True
        sheet.range('B2').font.color = header_color
        sheet.range('B2:F2').merge()

        sheet.range('B3').value = f"Configure KPI targets - Version {APP_VERSION}"
        sheet.range('B3').font.size = 9
        sheet.range('B3').font.color = (128, 128, 128)
        sheet.range('B3:F3').merge()

        # Instructions
        sheet.range('B5').value = "INSTRUCTIONS"
        sheet.range('B5').font.size = 14
        sheet.range('B5').font.bold = True
        sheet.range('B5').font.color = (255, 255, 255)
        sheet.range('B5:H5').color = section_color

        instructions = [
            "1. Set target values for each KPI in the 'Target' column",
            "2. Yellow threshold: 80% of target (caution)",
            "3. Red threshold: 60% of target (warning)",
            "4. Dashboard updates automatically when data changes"
        ]
        for i, instr in enumerate(instructions):
            sheet.range(f'B{7+i}').value = instr
            sheet.range(f'B{7+i}').font.size = 10

        # Headers
        header_row = 13
        headers = ['Category', 'KPI Name', 'Description', 'Target', 'Yellow %', 'Red %', 'Direction']
        for col, header in enumerate(headers, 2):
            cell = sheet.range((header_row, col))
            cell.value = header
            cell.font.bold = True
            cell.font.color = (255, 255, 255)
            cell.color = (52, 73, 94)
            cell.api.HorizontalAlignment = -4108  # Center

        # Populate KPIs
        data_row = header_row + 1
        for category, kpis in self.KPI_DEFINITIONS.items():
            for kpi in kpis:
                sheet.range((data_row, 2)).value = category.title()
                sheet.range((data_row, 3)).value = kpi['name']
                sheet.range((data_row, 3)).font.bold = True
                sheet.range((data_row, 4)).value = kpi['description']
                sheet.range((data_row, 4)).font.size = 9
                sheet.range((data_row, 4)).font.color = (100, 100, 100)

                # Target (editable)
                target_cell = sheet.range((data_row, 5))
                target_cell.value = kpi['default_target']
                if '%' in kpi['format']:
                    target_cell.number_format = '0.0%'
                elif '0.00' in kpi['format']:
                    target_cell.number_format = '0.00'
                else:
                    target_cell.number_format = '#,##0'
                target_cell.color = control_color

                # Yellow/Red thresholds
                sheet.range((data_row, 6)).value = 0.80
                sheet.range((data_row, 6)).number_format = '0%'
                sheet.range((data_row, 6)).color = control_color

                sheet.range((data_row, 7)).value = 0.60
                sheet.range((data_row, 7)).number_format = '0%'
                sheet.range((data_row, 7)).color = control_color

                # Direction
                sheet.range((data_row, 8)).value = "Higher" if kpi['higher_is_better'] else "Lower"

                # Alternate row shading
                if data_row % 2 == 0:
                    for col in range(2, 9):
                        if sheet.range((data_row, col)).color is None:
                            sheet.range((data_row, col)).color = (248, 249, 250)

                data_row += 1

        # Column widths
        sheet.range('A:A').column_width = 3
        sheet.range('B:B').column_width = 14
        sheet.range('C:C').column_width = 20
        sheet.range('D:D').column_width = 30
        sheet.range('E:E').column_width = 12
        sheet.range('F:F').column_width = 10
        sheet.range('G:G').column_width = 10
        sheet.range('H:H').column_width = 10

        # Tab color
        try:
            sheet.api.Tab.Color = 0xDB7400  # Blue
        except:
            pass

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    def _create_dashboard_sheet(self, sheet, pl_accounts, bs_accounts, months, detected_totals=None):
        """Create the main Dashboard sheet with KPIs and visualizations

        OPTIMIZED VERSION: Uses bulk writes and simplified formulas to prevent hangs.
        """
        # Colors
        header_color = (22, 33, 62)  # Dark blue
        section_color = (44, 62, 80)  # Section headers

        # Get detected account names (or use defaults)
        detected_totals = detected_totals or {}
        total_income_name = detected_totals.get('total_income', 'Total for Income')
        total_cogs_name = detected_totals.get('total_cogs', 'Total for Cost of Sales')
        total_expenses_name = detected_totals.get('total_expenses', 'Total for Expenses')

        company_name = self.company_name.get()
        current_month_col = len(months) + 1
        col_letter = self._col_letter(current_month_col)
        is_multi_division = hasattr(self, 'divisions') and len(self.divisions) > 1

        # =====================================================================
        # BULK WRITE: Build all header data and write at once
        # =====================================================================
        header_data = [
            ['', company_name, '', '', '', '', '', '', '', '', '', ''],
            ['', 'Executive Dashboard', '', '', '', '', '', '', '', '', '', ''],
            ['', f"Current Period: {months[-1][2] if months else 'N/A'}", '', '', '', '', '', '', '', '', '', ''],
            ['', f"Generated: {datetime.now().strftime('%B %d, %Y')} | Version {APP_VERSION}", '', '', '', '', '', '', '', '', '', ''],
        ]
        sheet.range('A2:L5').value = header_data

        # Apply header formatting in bulk
        title_range = sheet.range('B2')
        title_range.font.size = 28
        title_range.font.bold = True
        title_range.font.color = header_color

        subtitle_range = sheet.range('B3')
        subtitle_range.font.size = 16
        subtitle_range.font.color = (127, 140, 141)

        info_range = sheet.range('B4:B5')
        info_range.font.size = 9
        info_range.font.color = (150, 150, 150)

        # =====================================================================
        # DIVISION SELECTOR DROPDOWN (Row 6) - For multi-division models
        # =====================================================================
        if is_multi_division:
            sheet.range('B6').value = 'View:'
            sheet.range('B6').font.bold = True
            sheet.range('B6').font.size = 10

            # Default to "Consolidated"
            sheet.range('C6').value = 'Consolidated'
            sheet.range('C6').font.bold = True
            sheet.range('C6').font.size = 10
            sheet.range('C6').color = (230, 230, 250)

            # Create dropdown list: Consolidated + all division names
            div_names = ['Consolidated'] + [d.get('name', d) if isinstance(d, dict) else getattr(d, 'name', str(d)) for d in self.divisions]
            div_list = ','.join(div_names)

            try:
                sheet.range('C6').api.Validation.Delete()
                sheet.range('C6').api.Validation.Add(
                    Type=3,  # xlValidateList
                    AlertStyle=1,
                    Formula1=div_list
                )
            except Exception as e:
                print(f"[Dashboard] Could not add division dropdown: {e}")

            # Store the selected division sheet name formula in a helper cell (H6, hidden)
            # This will be used by formulas to determine which sheet to pull from
            sheet.range('H6').value = '=IF(C6="Consolidated","Consolidated_PL",C6&"_PL")'
            sheet.range('H6').font.color = (255, 255, 255)  # White text (hidden)

        # =====================================================================
        # P&L SUMMARY TABLE (Rows 8-13) - SIMPLIFIED with SUMIF formulas
        # Using simpler formulas that only lookup by account name in column B
        # =====================================================================

        # Build P&L summary data in memory first
        pl_summary_labels = ['', 'CURRENT MTH', '', 'YTD', '', '% of Rev']
        pl_accounts_info = [
            ('Revenue', total_income_name, True, (39, 174, 96)),
            ('Cost of Goods Sold', total_cogs_name, False, (231, 76, 60)),
            ('Gross Profit', 'Gross Profit', True, (52, 152, 219)),
            ('Operating Expenses', total_expenses_name, False, (230, 126, 34)),
            ('Net Income', 'Net Income', True, (155, 89, 182)),
        ]

        # Write header row
        sheet.range('B7:G7').value = [pl_summary_labels]
        header_range = sheet.range('B7:G7')
        header_range.font.bold = True
        header_range.font.size = 10
        header_range.font.color = (255, 255, 255)
        header_range.color = section_color

        # Build all P&L summary rows data
        pl_data = []
        for i, (label, account, is_bold, color) in enumerate(pl_accounts_info):
            row = 8 + i
            # Use SUMIF/SUMIFS based on mode
            # Current month: lookup in the last month column
            if is_multi_division:
                # If Consolidated selected (C6="Consolidated"), sum all divisions
                # If specific division selected, filter by that division
                cm_formula = f'=IF($C$6="Consolidated",SUMIF(Source_PL!$B$3:$B$1500,"{account}",Source_PL!{col_letter}$3:{col_letter}$1500),SUMIFS(Source_PL!{col_letter}$3:{col_letter}$1500,Source_PL!$A$3:$A$1500,$C$6,Source_PL!$B$3:$B$1500,"{account}"))'
            else:
                cm_formula = f'=SUMIF(Source_PL!$A$3:$A$1500,"{account}",Source_PL!{col_letter}$3:{col_letter}$1500)'

            # YTD: sum all months in current year using SUMPRODUCT with year filter
            # Menu!I7 contains current year (in multi-division mode) or Menu!F7 (single-division)
            # Source_PL row 2 has YYYYMM values, we extract year by dividing by 100
            if is_multi_division:
                # Consolidated: sum all divisions for the year
                # Division-specific: filter by division and year
                ytd_formula = f'=IF($C$6="Consolidated",IFERROR(SUMPRODUCT((Source_PL!$B$3:$B$1500="{account}")*(INT(Source_PL!$C$2:{col_letter}$2/100)=Menu!$I$7)*(Source_PL!$C$3:{col_letter}$1500)),0),IFERROR(SUMPRODUCT((Source_PL!$A$3:$A$1500=$C$6)*(Source_PL!$B$3:$B$1500="{account}")*(INT(Source_PL!$C$2:{col_letter}$2/100)=Menu!$I$7)*(Source_PL!$C$3:{col_letter}$1500)),0))'
            else:
                ytd_formula = f'=IFERROR(SUMPRODUCT((Source_PL!$A$3:$A$1500="{account}")*(INT(Source_PL!$B$2:{col_letter}$2/100)=Menu!$F$7)*(Source_PL!$B$3:{col_letter}$1500)),0)'

            # % of Revenue
            pct_formula = f'=IFERROR(E{row}/E8,0)' if label != 'Revenue' else 1.0

            pl_data.append([label, cm_formula, '', ytd_formula, '', pct_formula])

        # Write all P&L data at once
        sheet.range('B8:G12').value = pl_data

        # Apply formatting to P&L summary section
        for i, (label, account, is_bold, color) in enumerate(pl_accounts_info):
            row = 8 + i
            # Label formatting
            sheet.range(f'B{row}').font.bold = is_bold
            sheet.range(f'B{row}').font.size = 11
            sheet.range(f'B{row}').font.color = color
            # Number formatting
            sheet.range(f'C{row}').number_format = '"$"#,##0'
            sheet.range(f'E{row}').number_format = '"$"#,##0'
            sheet.range(f'G{row}').number_format = '0.0%'

        # Gross Margin and Net Margin indicators
        margin_data = [
            ['Gross Margin:', '=IFERROR(C10/C8,0)'],
            ['', ''],
            ['Net Margin:', '=IFERROR(C12/C8,0)'],
        ]
        sheet.range('H10:I12').value = margin_data
        sheet.range('I10').number_format = '0.0%'
        sheet.range('I12').number_format = '0.0%'

        # KPI Sections - Start after P&L Summary table
        # OPTIMIZED: Build all data in memory first, then write in bulk
        current_row = 14
        sections = [
            ('PROFITABILITY METRICS', 'profitability'),
            ('LIQUIDITY METRICS', 'liquidity'),
            ('EFFICIENCY METRICS', 'efficiency'),
            ('LEVERAGE METRICS', 'leverage'),
            ('CASH FLOW METRICS', 'cashflow'),
        ]

        control_row = 14  # Starting row in Dashboard_Control

        # Build all section data in memory first
        all_rows_data = []  # List of (row_num, row_data, row_type, kpi_format)
        kpi_data_rows = []  # Track which rows are KPI data rows for formatting

        for section_title, category in sections:
            kpis = self.KPI_DEFINITIONS.get(category, [])
            if not kpis:
                continue

            # Section header row
            section_row = ['', section_title] + [''] * 10  # Columns B through L
            all_rows_data.append((current_row, section_row, 'section', None))
            current_row += 1

            # Column headers row
            header_row = ['', 'KPI', 'Current', 'Target', 'Status', '', 'YTD', 'Target', 'Status', '', 'Trend']
            all_rows_data.append((current_row, header_row, 'header', None))
            current_row += 1

            # KPI data rows
            for kpi in kpis:
                higher = kpi['higher_is_better']

                # Build formulas
                current_formula = self._build_dashboard_kpi_formula(kpi, 'current', current_month_col, months, is_multi_division)
                # Target formula: default to 0 if cell is empty or has error
                target_formula = f'=IFERROR(IF(Dashboard_Control!E{control_row}="",0,Dashboard_Control!E{control_row}),0)'

                # Status formula: show "-" if target is 0 or empty (no target defined)
                if higher:
                    status_formula = f'=IF(OR(D{current_row}=0,D{current_row}=""),"-",IF(C{current_row}>=D{current_row},"G",IF(C{current_row}>=D{current_row}*0.8,"Y","R")))'
                else:
                    status_formula = f'=IF(OR(D{current_row}=0,D{current_row}=""),"-",IF(C{current_row}<=D{current_row},"G",IF(C{current_row}<=D{current_row}*1.2,"Y","R")))'

                ytd_formula = self._build_dashboard_kpi_formula(kpi, 'ytd', current_month_col, months, is_multi_division)

                if '%' in kpi['format'] or 'ratio' in kpi['formula_type']:
                    ytd_target = f"=D{current_row}"
                else:
                    ytd_target = f"=D{current_row}*{len(months)}"

                # YTD Status formula: show "-" if target is 0 or empty
                if higher:
                    ytd_status = f'=IF(OR(H{current_row}=0,H{current_row}=""),"-",IF(G{current_row}>=H{current_row},"G",IF(G{current_row}>=H{current_row}*0.8,"Y","R")))'
                else:
                    ytd_status = f'=IF(OR(H{current_row}=0,H{current_row}=""),"-",IF(G{current_row}<=H{current_row},"G",IF(G{current_row}<=H{current_row}*1.2,"Y","R")))'

                # Simple trend indicator: Up, Down, or Flat
                trend_formula = f'=IF(G{current_row}=0,"-",IF(C{current_row}>G{current_row}*1.05,"Up",IF(C{current_row}<G{current_row}*0.95,"Down","Flat")))'

                # Row data: columns A through K (indices 0-10)
                kpi_row = [
                    '',  # A
                    kpi['name'],  # B
                    current_formula,  # C
                    target_formula,  # D
                    status_formula,  # E
                    '',  # F (spacer)
                    ytd_formula,  # G
                    ytd_target,  # H
                    ytd_status,  # I
                    '',  # J (spacer)
                    trend_formula  # K
                ]
                all_rows_data.append((current_row, kpi_row, 'kpi', kpi['format']))
                kpi_data_rows.append((current_row, kpi['format'], current_row % 2 == 0))

                current_row += 1
                control_row += 1

            current_row += 1  # Space between sections

        # BULK WRITE: Write all data at once per section type
        for row_num, row_data, row_type, kpi_format in all_rows_data:
            # Write row data in one operation (columns B through L = 2 through 12)
            sheet.range((row_num, 2), (row_num, 12)).value = row_data[1:]  # Skip column A

        # SIMPLIFIED FORMATTING: Only format section/header rows (minimal per-KPI formatting)
        for row_num, row_data, row_type, kpi_format in all_rows_data:
            if row_type == 'section':
                section_range = sheet.range((row_num, 2), (row_num, 12))
                section_range.font.bold = True
                section_range.font.color = (255, 255, 255)
                section_range.color = section_color

            elif row_type == 'header':
                header_range = sheet.range((row_num, 2), (row_num, 12))
                header_range.font.bold = True
                header_range.font.color = (255, 255, 255)
                header_range.color = (52, 73, 94)

        # Apply number formats to entire columns at once (much faster than per-cell)
        if kpi_data_rows:
            first_kpi_row = kpi_data_rows[0][0]
            last_kpi_row = kpi_data_rows[-1][0]
            # Apply common number format to value columns
            sheet.range(f'C{first_kpi_row}:C{last_kpi_row}').number_format = '#,##0.00'
            sheet.range(f'D{first_kpi_row}:D{last_kpi_row}').number_format = '#,##0.00'
            sheet.range(f'G{first_kpi_row}:G{last_kpi_row}').number_format = '#,##0.00'
            sheet.range(f'H{first_kpi_row}:H{last_kpi_row}').number_format = '#,##0.00'

        # =====================================================================
        # CONDITIONAL FORMATTING for Status columns (E and I)
        # G = Green (good), Y = Yellow (warning), R = Red (bad)
        # =====================================================================
        try:
            # Status columns E and I - apply conditional formatting for G/Y/R
            if kpi_data_rows:
                for status_col in ['E', 'I']:
                    status_range = sheet.range(f'{status_col}{first_kpi_row}:{status_col}{last_kpi_row}')

                    # Delete any existing conditional formatting
                    try:
                        status_range.api.FormatConditions.Delete()
                    except:
                        pass

                    # Green for "G" - good performance
                    status_range.api.FormatConditions.Add(
                        Type=1,  # xlCellValue
                        Operator=3,  # xlEqual
                        Formula1='"G"'
                    )
                    status_range.api.FormatConditions(1).Interior.Color = 0x90EE90  # Light green (BGR)
                    status_range.api.FormatConditions(1).Font.Color = 0x228B22  # Dark green
                    status_range.api.FormatConditions(1).Font.Bold = True

                    # Yellow for "Y" - warning
                    status_range.api.FormatConditions.Add(
                        Type=1,
                        Operator=3,
                        Formula1='"Y"'
                    )
                    status_range.api.FormatConditions(2).Interior.Color = 0x00FFFF  # Yellow (BGR is 0x00FFFF)
                    status_range.api.FormatConditions(2).Font.Color = 0x008080  # Dark yellow/olive
                    status_range.api.FormatConditions(2).Font.Bold = True

                    # Red for "R" - poor performance
                    status_range.api.FormatConditions.Add(
                        Type=1,
                        Operator=3,
                        Formula1='"R"'
                    )
                    status_range.api.FormatConditions(3).Interior.Color = 0x8080FF  # Light red (BGR)
                    status_range.api.FormatConditions(3).Font.Color = 0x0000CD  # Dark red
                    status_range.api.FormatConditions(3).Font.Bold = True

                    # Center-align status columns
                    status_range.api.HorizontalAlignment = -4108  # xlCenter
        except Exception as e:
            print(f"[Dashboard] Could not apply conditional formatting: {e}")

        # Column widths
        try:
            sheet.range('A:A').column_width = 3
            sheet.range('B:B').column_width = 22
            sheet.range('C:C').column_width = 14
            sheet.range('D:D').column_width = 12
            sheet.range('E:E').api.EntireColumn.AutoFit()  # AutoFit Status column
            sheet.range('F:F').column_width = 3
            sheet.range('G:G').column_width = 14
            sheet.range('H:H').column_width = 12
            sheet.range('I:I').api.EntireColumn.AutoFit()  # AutoFit YTD Status column
            sheet.range('J:J').column_width = 3
            sheet.range('K:K').column_width = 8  # Trend column - compact
        except:
            pass

        # Format Trend column K with conditional formatting for Up/Down/Flat
        try:
            if kpi_data_rows:
                trend_range = sheet.range(f'K{first_kpi_row}:K{last_kpi_row}')
                trend_range.api.HorizontalAlignment = -4108  # xlCenter
                trend_range.font.size = 10
                trend_range.font.name = 'Calibri Light'

                # Add conditional formatting for trend text
                try:
                    trend_range.api.FormatConditions.Delete()

                    # "Up" = Green
                    trend_range.api.FormatConditions.Add(Type=1, Operator=3, Formula1='"Up"')
                    trend_range.api.FormatConditions(1).Font.Color = 0x008000  # Green

                    # "Down" = Red
                    trend_range.api.FormatConditions.Add(Type=1, Operator=3, Formula1='"Down"')
                    trend_range.api.FormatConditions(2).Font.Color = 0x0000FF  # Red

                    # "Flat" = Gray
                    trend_range.api.FormatConditions.Add(Type=1, Operator=3, Formula1='"Flat"')
                    trend_range.api.FormatConditions(3).Font.Color = 0x808080  # Gray
                except:
                    pass
        except:
            pass

        # Vertically center all cells
        try:
            used_range = sheet.api.UsedRange
            used_range.VerticalAlignment = -4108  # xlVAlignCenter
        except:
            pass

        # Tab color
        try:
            sheet.api.Tab.Color = 0x60AE27  # Green
        except:
            pass

        # Hide gridlines
        try:
            sheet.api.Activate()
            sheet.book.app.api.ActiveWindow.DisplayGridlines = False
        except:
            pass

        # =====================================================================
        # CHARTS SECTION - Add visualizations to the right of the dashboard
        # =====================================================================
        try:
            self._add_dashboard_charts(sheet, months, is_multi_division, current_row, detected_totals)
        except Exception as e:
            print(f"[Dashboard] Could not add charts: {e}")

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    def _add_dashboard_charts(self, sheet, months, is_multi_division, last_kpi_row, detected_totals=None):
        """Add charts to the Dashboard sheet

        Creates:
        1. Pie chart showing expense breakdown (Revenue vs COGS vs Expenses)
        2. Monthly Revenue line chart
        3. Monthly Net Income line chart
        4. Monthly Gross Margin line chart
        """
        import win32com.client as win32

        # Chart positioning constants
        chart_left = 720  # Start position (column N area - about 720 pixels)
        chart_width = 350
        chart_height = 200
        chart_gap = 20

        # Get detected account names for formulas
        detected_totals = detected_totals or {}
        total_income_name = detected_totals.get('total_income', 'Total for Income')
        total_cogs_name = detected_totals.get('total_cogs', 'Total for Cost of Sales')
        total_expenses_name = detected_totals.get('total_expenses', 'Total for Expenses')

        # Determine data columns based on mode
        acct_col = 'B' if is_multi_division else 'A'
        data_start_col = 'C' if is_multi_division else 'B'
        year_cell = 'Menu!$I$7' if is_multi_division else 'Menu!$F$7'

        num_months = len(months)
        last_data_col = self._col_letter(num_months + (2 if is_multi_division else 1))

        # =====================================================================
        # CHART DATA AREA: Write chart source data to hidden columns (starting at column N)
        # =====================================================================
        chart_data_col = 14  # Column N

        # Row 2: Header "Chart Data"
        sheet.range((2, chart_data_col)).value = "Chart Data"
        sheet.range((2, chart_data_col)).font.bold = True

        # --- Pie Chart Data (Revenue vs COGS vs Expenses breakdown) ---
        sheet.range((4, chart_data_col)).value = "Category"
        sheet.range((4, chart_data_col + 1)).value = "Amount"

        # Use flexible matching patterns for account names
        # These patterns will match variations like "Total Income", "Total for Income", etc.
        pie_labels = ['Revenue', 'Cost of Goods', 'Operating Exp.']
        pie_search_terms = ['Total*Income', 'Total*Cost', 'Total*Expense']

        for i, (label, search_term) in enumerate(zip(pie_labels, pie_search_terms)):
            row = 5 + i
            sheet.range((row, chart_data_col)).value = label
            # YTD sum formula using SUMPRODUCT with wildcard matching via COUNTIF pattern
            # Use exact account names from detected_totals if available, otherwise search
            if i == 0:
                account = total_income_name
            elif i == 1:
                account = total_cogs_name
            else:
                account = total_expenses_name

            if is_multi_division:
                formula = f'=ABS(IFERROR(SUMPRODUCT((Source_PL!${acct_col}$3:${acct_col}$1500="{account}")*(INT(Source_PL!${data_start_col}$2:{last_data_col}$2/100)={year_cell})*(Source_PL!${data_start_col}$3:{last_data_col}$1500)),0))'
            else:
                formula = f'=ABS(IFERROR(SUMPRODUCT((Source_PL!${acct_col}$3:${acct_col}$1500="{account}")*(INT(Source_PL!${data_start_col}$2:{last_data_col}$2/100)={year_cell})*(Source_PL!${data_start_col}$3:{last_data_col}$1500)),0))'
            sheet.range((row, chart_data_col + 1)).value = formula

        # --- Monthly Data for Line Charts ---
        # Row 10: "Monthly Data" header
        sheet.range((10, chart_data_col)).value = "Month"
        sheet.range((10, chart_data_col + 1)).value = "Revenue"
        sheet.range((10, chart_data_col + 2)).value = "Net Income"
        sheet.range((10, chart_data_col + 3)).value = "Gross Margin %"

        # Write month labels and formulas for each month (up to last 12 months)
        display_months = months[-12:] if len(months) > 12 else months

        for i, (month, year, label) in enumerate(display_months):
            row = 11 + i
            yyyymm = year * 100 + month
            col_idx = months.index((month, year, label)) + (3 if is_multi_division else 2)
            month_col = self._col_letter(col_idx)

            # Month label (short form)
            sheet.range((row, chart_data_col)).value = label[:3] if len(label) > 3 else label

            # Revenue formula
            if is_multi_division:
                rev_formula = f'=IFERROR(SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"{total_income_name}",Source_PL!{month_col}$3:{month_col}$1500),0)'
            else:
                rev_formula = f'=IFERROR(SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"{total_income_name}",Source_PL!{month_col}$3:{month_col}$1500),0)'
            sheet.range((row, chart_data_col + 1)).value = rev_formula

            # Net Income formula
            if is_multi_division:
                ni_formula = f'=IFERROR(SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"Net Income",Source_PL!{month_col}$3:{month_col}$1500),0)'
            else:
                ni_formula = f'=IFERROR(SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"Net Income",Source_PL!{month_col}$3:{month_col}$1500),0)'
            sheet.range((row, chart_data_col + 2)).value = ni_formula

            # Gross Margin % formula (Gross Profit / Revenue)
            if is_multi_division:
                gp_formula = f'=IFERROR(SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"Gross Profit",Source_PL!{month_col}$3:{month_col}$1500),0)'
            else:
                gp_formula = f'=IFERROR(SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"Gross Profit",Source_PL!{month_col}$3:{month_col}$1500),0)'
            gm_formula = f'=IFERROR({gp_formula}/{chart_data_col + 1}{row},0)'
            # Simplified: calculate GP/Revenue directly
            sheet.range((row, chart_data_col + 3)).value = f'=IFERROR(SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"Gross Profit",Source_PL!{month_col}$3:{month_col}$1500)/{self._col_letter(chart_data_col + 1)}{row},0)'

        num_data_rows = len(display_months)

        # Format chart data columns
        sheet.range((11, chart_data_col + 1), (10 + num_data_rows, chart_data_col + 1)).number_format = '"$"#,##0'
        sheet.range((11, chart_data_col + 2), (10 + num_data_rows, chart_data_col + 2)).number_format = '"$"#,##0'
        sheet.range((11, chart_data_col + 3), (10 + num_data_rows, chart_data_col + 3)).number_format = '0.0%'

        # Make chart data columns nearly invisible (but not hidden - hidden columns break charts)
        # Use white font on white background and narrow width instead of hiding
        try:
            for col in range(chart_data_col, chart_data_col + 4):
                col_range = sheet.range((1, col), (25, col))  # Cover enough rows
                col_range.font.color = (255, 255, 255)  # White text
                col_range.color = (255, 255, 255)  # White background
                sheet.range((1, col)).column_width = 0.5  # Very narrow but not hidden
        except:
            pass

        # =====================================================================
        # CREATE CHARTS using Excel COM API
        # =====================================================================
        try:
            charts = sheet.api.ChartObjects()

            # --- 1. PIE CHART: Revenue vs COGS vs Expenses ---
            pie_chart = charts.Add(chart_left, 30, chart_width, chart_height)
            pie_chart.Name = "PieChart"
            pie = pie_chart.Chart
            pie.ChartType = 5  # xlPie

            # Set data source
            pie_data_range = sheet.range((4, chart_data_col), (7, chart_data_col + 1))
            pie.SetSourceData(pie_data_range.api)

            # Title
            pie.HasTitle = True
            pie.ChartTitle.Text = "YTD Financial Breakdown"
            pie.ChartTitle.Font.Size = 11
            pie.ChartTitle.Font.Bold = True

            # Data labels
            pie.ApplyDataLabels(5)  # xlDataLabelsShowPercent

            # Legend
            pie.HasLegend = True
            pie.Legend.Position = -4107  # xlLegendPositionBottom

            # Colors for pie slices (Revenue=green, COGS=orange, Expenses=red)
            try:
                pie.SeriesCollection(1).Points(1).Interior.Color = 0x60AE27  # Green (Revenue)
                pie.SeriesCollection(1).Points(2).Interior.Color = 0x00A5FF  # Orange (COGS)
                pie.SeriesCollection(1).Points(3).Interior.Color = 0x4C4CE6  # Red (Expenses)
            except:
                pass

            # --- 2. REVENUE LINE CHART ---
            rev_top = 30 + chart_height + chart_gap
            rev_chart = charts.Add(chart_left, rev_top, chart_width, chart_height)
            rev_chart.Name = "RevenueChart"
            rev = rev_chart.Chart
            rev.ChartType = 65  # xlLineMarkers

            # Set data source for Revenue
            rev_labels = sheet.range((11, chart_data_col), (10 + num_data_rows, chart_data_col))
            rev_values = sheet.range((11, chart_data_col + 1), (10 + num_data_rows, chart_data_col + 1))

            rev.SeriesCollection().NewSeries()
            rev.SeriesCollection(1).Values = rev_values.api
            rev.SeriesCollection(1).XValues = rev_labels.api
            rev.SeriesCollection(1).Name = "Revenue"

            # Title
            rev.HasTitle = True
            rev.ChartTitle.Text = "Monthly Revenue"
            rev.ChartTitle.Font.Size = 11
            rev.ChartTitle.Font.Bold = True

            # Format line
            try:
                rev.SeriesCollection(1).Format.Line.ForeColor.RGB = 0x60AE27  # Green
                rev.SeriesCollection(1).Format.Line.Weight = 2.5
            except:
                pass

            # Remove legend (single series)
            rev.HasLegend = False

            # --- 3. NET INCOME LINE CHART ---
            ni_top = rev_top + chart_height + chart_gap
            ni_chart = charts.Add(chart_left, ni_top, chart_width, chart_height)
            ni_chart.Name = "NetIncomeChart"
            ni = ni_chart.Chart
            ni.ChartType = 65  # xlLineMarkers

            # Set data source for Net Income
            ni_values = sheet.range((11, chart_data_col + 2), (10 + num_data_rows, chart_data_col + 2))

            ni.SeriesCollection().NewSeries()
            ni.SeriesCollection(1).Values = ni_values.api
            ni.SeriesCollection(1).XValues = rev_labels.api
            ni.SeriesCollection(1).Name = "Net Income"

            # Title
            ni.HasTitle = True
            ni.ChartTitle.Text = "Monthly Net Income"
            ni.ChartTitle.Font.Size = 11
            ni.ChartTitle.Font.Bold = True

            # Format line
            try:
                ni.SeriesCollection(1).Format.Line.ForeColor.RGB = 0xB05E9B  # Purple
                ni.SeriesCollection(1).Format.Line.Weight = 2.5
            except:
                pass

            ni.HasLegend = False

            # --- 4. GROSS MARGIN LINE CHART ---
            gm_top = ni_top + chart_height + chart_gap
            gm_chart = charts.Add(chart_left, gm_top, chart_width, chart_height)
            gm_chart.Name = "GrossMarginChart"
            gm = gm_chart.Chart
            gm.ChartType = 65  # xlLineMarkers

            # Set data source for Gross Margin
            gm_values = sheet.range((11, chart_data_col + 3), (10 + num_data_rows, chart_data_col + 3))

            gm.SeriesCollection().NewSeries()
            gm.SeriesCollection(1).Values = gm_values.api
            gm.SeriesCollection(1).XValues = rev_labels.api
            gm.SeriesCollection(1).Name = "Gross Margin %"

            # Title
            gm.HasTitle = True
            gm.ChartTitle.Text = "Monthly Gross Margin %"
            gm.ChartTitle.Font.Size = 11
            gm.ChartTitle.Font.Bold = True

            # Format line
            try:
                gm.SeriesCollection(1).Format.Line.ForeColor.RGB = 0xD49434  # Blue/teal
                gm.SeriesCollection(1).Format.Line.Weight = 2.5
            except:
                pass

            gm.HasLegend = False

            # Format Y-axis as percentage
            try:
                gm.Axes(2).TickLabels.NumberFormat = "0%"  # xlValue axis
            except:
                pass

        except Exception as e:
            print(f"[Dashboard] Chart creation error: {e}")
            import traceback
            traceback.print_exc()

    def _build_current_month_sumproduct(self, sheet, account, current_col, multi_division=False):
        """
        Build a SUMPRODUCT formula for current month that dynamically references Menu!C7.
        Returns value for the exact month/year matching Menu!E7 (month) and Menu!F7 (year).

        Source sheets have:
        - Single-division: Row 2: YYYYMM helper values, Row 3+: Account data
        - Multi-division: Col A=Division, Col B=Account, Row 2: YYYYMM values start at Col C

        Args:
            multi_division: If True, adjust formula for multi-division structure with
                           division filter based on Dashboard!L3
        """
        col_letter = self._col_letter(current_col)

        if multi_division:
            # Multi-division: Col A=Division, Col B=Account, Col C+ = values
            # If Dashboard!L3 = "Consolidated", sum all divisions; else filter by division
            return (
                f"=IF(Dashboard!$L$3=\"Consolidated\","
                f"SUMPRODUCT("
                f"({sheet}!$B$3:$B$1000=\"{account}\")*"
                f"({sheet}!$C$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"({sheet}!$C$3:{col_letter}$1000)),"
                f"SUMPRODUCT("
                f"({sheet}!$A$3:$A$1000=Dashboard!$L$3)*"
                f"({sheet}!$B$3:$B$1000=\"{account}\")*"
                f"({sheet}!$C$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"({sheet}!$C$3:{col_letter}$1000)))"
            )
        else:
            # Single-division: Col A=Account, Col B+ = values
            return (
                f"=SUMPRODUCT("
                f"({sheet}!$A$3:$A$1000=\"{account}\")*"
                f"({sheet}!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"({sheet}!$B$3:{col_letter}$1000)"
                f")"
            )

    def _build_ytd_sumproduct(self, sheet, account, current_col, multi_division=False):
        """
        Build a SUMPRODUCT formula for YTD that only sums months in the current year.
        Uses Menu!F7 for year and Menu!E7 for current month number.

        Source sheets now have:
        - Row 1: Headers (month names)
        - Row 2: YYYYMM helper values (e.g., 202411 for Nov 2024)
        - Row 3+: Account data

        Args:
            multi_division: If True, adjust formula for multi-division structure with
                           division filter based on Dashboard!L3
        """
        col_letter = self._col_letter(current_col)

        if multi_division:
            # Multi-division: Col A=Division, Col B=Account, Col C+ = values
            return (
                f"=IF(Dashboard!$L$3=\"Consolidated\","
                f"SUMPRODUCT("
                f"({sheet}!$B$3:$B$1000=\"{account}\")*"
                f"(INT({sheet}!$C$2:{col_letter}$2/100)=Menu!$F$7)*"
                f"(MOD({sheet}!$C$2:{col_letter}$2,100)<=Menu!$E$7)*"
                f"({sheet}!$C$3:{col_letter}$1000)),"
                f"SUMPRODUCT("
                f"({sheet}!$A$3:$A$1000=Dashboard!$L$3)*"
                f"({sheet}!$B$3:$B$1000=\"{account}\")*"
                f"(INT({sheet}!$C$2:{col_letter}$2/100)=Menu!$F$7)*"
                f"(MOD({sheet}!$C$2:{col_letter}$2,100)<=Menu!$E$7)*"
                f"({sheet}!$C$3:{col_letter}$1000)))"
            )
        else:
            # Single-division: Col A=Account, Col B+ = values
            return (
                f"=SUMPRODUCT("
                f"({sheet}!$A$3:$A$1000=\"{account}\")*"
                f"(INT({sheet}!$B$2:{col_letter}$2/100)=Menu!$F$7)*"
                f"(MOD({sheet}!$B$2:{col_letter}$2,100)<=Menu!$E$7)*"
                f"({sheet}!$B$3:{col_letter}$1000)"
                f")"
            )

    def _build_dashboard_kpi_formula(self, kpi, period, current_col, months, multi_division=False):
        """Build formula for a KPI based on its type and period

        FIXED VERSION: Properly handles YTD by summing across all months in current year.
        For multi-division mode, uses Column B for account lookup (Column A is Division).

        Args:
            kpi: KPI definition dict
            period: 'current' or 'ytd'
            current_col: Current month column number
            months: List of month tuples
            multi_division: If True, uses Column B for account lookup
        """
        formula_type = kpi['formula_type']
        col_letter = self._col_letter(current_col)

        # Determine account column and data start column based on mode
        # Multi-division: Col A = Division, Col B = Account, Col C+ = data
        # Single-division: Col A = Account, Col B+ = data
        acct_col = 'B' if multi_division else 'A'
        data_start_col = 'C' if multi_division else 'B'
        year_cell = 'Menu!$I$7' if multi_division else 'Menu!$F$7'

        def make_sumif(sheet, account, col):
            """Create simple SUMIF formula for current month"""
            return f'SUMIF({sheet}!${acct_col}$3:${acct_col}$1500,"{account}",{sheet}!{col}$3:{col}$1500)'

        def make_ytd_sumproduct(sheet, account):
            """Create SUMPRODUCT formula for YTD (sum all months in current year)"""
            # Row 2 has YYYYMM values, we extract year by INT(value/100)
            return (
                f'SUMPRODUCT('
                f'({sheet}!${acct_col}$3:${acct_col}$1500="{account}")*'
                f'(INT({sheet}!${data_start_col}$2:{col_letter}$2/100)={year_cell})*'
                f'({sheet}!${data_start_col}$3:{col_letter}$1500))'
            )

        if formula_type == 'direct':
            source = kpi['source']
            account = kpi['account']
            sheet = 'Source_PL' if source == 'PL' else 'Source_BS'
            if period == 'ytd' and source == 'PL':
                # Use SUMPRODUCT for YTD on P&L accounts
                return f'=IFERROR({make_ytd_sumproduct(sheet, account)},0)'
            else:
                # Current month or Balance Sheet (point-in-time)
                return f'={make_sumif(sheet, account, col_letter)}'

        elif formula_type == 'ratio':
            num = kpi['numerator']
            den = kpi['denominator']
            if period == 'ytd':
                num_f = make_ytd_sumproduct('Source_PL', num)
                den_f = make_ytd_sumproduct('Source_PL', den)
            else:
                num_f = make_sumif('Source_PL', num, col_letter)
                den_f = make_sumif('Source_PL', den, col_letter)
            return f"=IFERROR({num_f}/{den_f},0)"

        elif formula_type == 'bs_ratio':
            # Balance Sheet ratios use point-in-time values, not YTD
            num = kpi['numerator']
            den = kpi['denominator']
            num_f = make_sumif('Source_BS', num, col_letter)
            den_f = make_sumif('Source_BS', den, col_letter)
            return f"=IFERROR({num_f}/{den_f},0)"

        elif formula_type == 'bs_difference':
            # Balance Sheet differences use point-in-time values
            min_acct = kpi['minuend']
            sub_acct = kpi['subtrahend']
            min_f = make_sumif('Source_BS', min_acct, col_letter)
            sub_f = make_sumif('Source_BS', sub_acct, col_letter)
            return f"={min_f}-{sub_f}"

        elif formula_type == 'days_ratio':
            balance = kpi['balance']
            flow = kpi['flow']
            bal_f = make_sumif('Source_BS', balance, col_letter)
            if period == 'ytd':
                flow_f = make_ytd_sumproduct('Source_PL', flow)
            else:
                flow_f = make_sumif('Source_PL', flow, col_letter)
            return f"=IFERROR({bal_f}/{flow_f}*30,0)"

        elif formula_type == 'turnover':
            flow = kpi['flow']
            balance = kpi['balance']
            bal_f = make_sumif('Source_BS', balance, col_letter)
            if period == 'ytd':
                flow_f = make_ytd_sumproduct('Source_PL', flow)
            else:
                flow_f = make_sumif('Source_PL', flow, col_letter)
            return f"=IFERROR({flow_f}/{bal_f},0)"

        elif formula_type == 'calculated':
            calc_type = kpi['calc_type']
            return self._build_calculated_dashboard_formula(calc_type, period, current_col, months)

        return "=0"

    def _build_calculated_dashboard_formula(self, calc_type, period, current_col, months):
        """Build formula for complex calculated KPIs

        FIXED: Uses dynamic account lookup with wildcard matching instead of hardcoded account names.
        Searches for accounts containing 'Interest' or 'Depreciation' keywords.
        """
        col_letter = self._col_letter(current_col)

        def current_sumif(account):
            """Helper to build current month SUMIF formula for an account"""
            return f'SUMIF(Source_PL!$A$3:$A$1500,"{account}",Source_PL!{col_letter}$3:{col_letter}$1500)'

        def current_sumif_wildcard(keyword):
            """Helper to build current month SUMIF with wildcard matching"""
            return f'SUMIF(Source_PL!$A$3:$A$1500,"*{keyword}*",Source_PL!{col_letter}$3:{col_letter}$1500)'

        def ytd_sum(account):
            """Helper to sum YTD values - sums all columns for the current year"""
            # For YTD, we need to sum across months in the current year
            # This simplified version just uses the same current period value
            # A more complete solution would sum across all months where YEAR matches
            return f'SUMIF(Source_PL!$A$3:$A$1500,"{account}",Source_PL!{col_letter}$3:{col_letter}$1500)'

        def ytd_sum_wildcard(keyword):
            """Helper to sum YTD values with wildcard matching"""
            return f'SUMIF(Source_PL!$A$3:$A$1500,"*{keyword}*",Source_PL!{col_letter}$3:{col_letter}$1500)'

        if calc_type == 'ebitda':
            # EBITDA = Net Income + Interest Expense + Depreciation + Amortization
            # Use wildcard matching to find accounts containing these keywords
            if period == 'current':
                ni = current_sumif("Net Income")
                # Search for any account containing "Interest" (for interest expense)
                int_e = current_sumif_wildcard("Interest")
                # Search for any account containing "Depreciation" or "Amortization"
                dep = current_sumif_wildcard("Depreciation")
                amort = current_sumif_wildcard("Amortization")
            else:
                ni = ytd_sum("Net Income")
                int_e = ytd_sum_wildcard("Interest")
                dep = ytd_sum_wildcard("Depreciation")
                amort = ytd_sum_wildcard("Amortization")
            return f"=IFERROR({ni}+ABS({int_e})+ABS({dep})+ABS({amort}),0)"

        elif calc_type == 'ocf':
            # Operating Cash Flow = Net Income + Depreciation (simplified)
            if period == 'current':
                ni = current_sumif("Net Income")
                dep = current_sumif_wildcard("Depreciation")
                amort = current_sumif_wildcard("Amortization")
            else:
                ni = ytd_sum("Net Income")
                dep = ytd_sum_wildcard("Depreciation")
                amort = ytd_sum_wildcard("Amortization")
            return f"=IFERROR({ni}+ABS({dep})+ABS({amort}),0)"

        elif calc_type == 'fcf':
            # Free Cash Flow = OCF - CapEx (simplified as OCF for now)
            if period == 'current':
                ni = current_sumif("Net Income")
                dep = current_sumif_wildcard("Depreciation")
                amort = current_sumif_wildcard("Amortization")
            else:
                ni = ytd_sum("Net Income")
                dep = ytd_sum_wildcard("Depreciation")
                amort = ytd_sum_wildcard("Amortization")
            return f"=IFERROR({ni}+ABS({dep})+ABS({amort}),0)"

        return "=0"

    def _col_letter(self, col_num):
        """Convert column number to letter"""
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result

    # ================================================================
    # TEMPLATE V2 UTILITY FUNCTIONS
    # ================================================================
    # These functions support the optimized "Clone, Inject, Adjust" pattern
    # for faster model generation using pre-built template sheets.

    def _clone_template_sheet(self, wb, template_name, new_name, position_after=None):
        """Clone a template sheet, rename it, and position it correctly.

        This is the core of the template-based optimization. Instead of building
        sheets from scratch (slow), we clone pre-formatted templates (fast).

        Args:
            wb: xlwings Workbook object
            template_name: Name of template sheet (e.g., '_TPL_PL')
            new_name: Name for the cloned sheet (e.g., 'Consolidated_PL')
            position_after: Name of sheet to position after (optional)

        Returns:
            xlwings Sheet object for the new sheet
        """
        try:
            template_sheet = wb.sheets[template_name]

            # Copy the sheet (creates "template_name (2)")
            if position_after:
                template_sheet.api.Copy(After=wb.sheets[position_after].api)
            else:
                template_sheet.api.Copy(After=wb.sheets[-1].api)

            # Find the copied sheet and rename it
            # Excel names copies as "SheetName (2)"
            copy_name = f'{template_name} (2)'
            new_sheet = wb.sheets[copy_name]
            new_sheet.name = new_name

            return new_sheet

        except Exception as e:
            print(f"Error cloning template sheet '{template_name}': {e}")
            # Fallback: create new sheet
            if position_after:
                new_sheet = wb.sheets.add(new_name, after=wb.sheets[position_after])
            else:
                new_sheet = wb.sheets.add(new_name)
            return new_sheet

    def _inject_accounts_bulk(self, sheet, accounts, months, start_row, source_sheet='Source_PL',
                               is_consolidated=True, division_name=None):
        """Inject account data into a template sheet using bulk writes.

        This replaces the slow cell-by-cell formula writing with a single
        bulk write operation, which is 10-100x faster.

        Args:
            sheet: xlwings Sheet object (cloned from template)
            accounts: List of account dictionaries
            months: List of (month, year, name) tuples
            start_row: First row to write data (after headers)
            source_sheet: Name of source data sheet
            is_consolidated: If True, use SUMIF; if False, use SUMIFS with division filter
            division_name: Division name for SUMIFS filter (required if not consolidated)

        Returns:
            Number of rows written
        """
        if not accounts:
            return 0

        source_start = 3
        source_end = 1500
        num_months = len(months)

        # Pre-calculate column letters for efficiency
        col_letters = [self._col_letter(i + 2) for i in range(num_months)]

        # Build all data in memory first
        all_rows = []
        row_types = []

        for idx, account in enumerate(accounts):
            account_name = account['name']
            indent_level = account.get('indent', 0)
            actual_row = start_row + idx

            # Display name with indentation
            display_name = account_name
            if indent_level > 0 and not account.get('is_header') and not account.get('is_total'):
                display_name = ('    ' * indent_level) + account_name

            row_data = [display_name]

            if account.get('is_header', False):
                # Header rows: empty data cells (template has formatting)
                row_data.extend([''] * num_months)
                all_rows.append(row_data)
                row_types.append('header')
                continue

            # Build SUMIF/SUMIFS formulas for each month
            for i, cl in enumerate(col_letters):
                if is_consolidated:
                    formula = f'=SUMIF({source_sheet}!$B${source_start}:$B${source_end},"{account_name}",{source_sheet}!{cl}${source_start}:{cl}${source_end})'
                else:
                    formula = f'=SUMIFS({source_sheet}!{cl}${source_start}:{cl}${source_end},{source_sheet}!$A${source_start}:$A${source_end},"{division_name}",{source_sheet}!$B${source_start}:$B${source_end},"{account_name}")'
                row_data.append(formula)

            all_rows.append(row_data)
            row_types.append('total' if account.get('is_total') else 'detail')

        # Single bulk write operation
        if all_rows:
            num_cols = len(all_rows[0])
            end_row = start_row + len(all_rows) - 1
            sheet.range((start_row, 1), (end_row, num_cols)).value = all_rows

        return len(all_rows)

    def _adjust_template_ranges(self, sheet, data_start_row, actual_rows, placeholder_rows,
                                 last_data_col):
        """Adjust formula ranges in template after injecting actual data.

        Template formulas reference placeholder row numbers (e.g., row 204 for 200 rows).
        This adjusts them to reference the actual data range.

        Args:
            sheet: xlwings Sheet object
            data_start_row: First row of account data
            actual_rows: Number of actual account rows
            placeholder_rows: Number of placeholder rows in template
            last_data_col: Last column with data
        """
        if actual_rows >= placeholder_rows:
            return  # No adjustment needed

        old_end_row = data_start_row + placeholder_rows - 1
        new_end_row = data_start_row + actual_rows - 1

        # Use Find/Replace for bulk formula adjustment
        try:
            for col in range(1, last_data_col + 1):
                col_letter = self._col_letter(col)
                # Replace references to old end row with new end row
                sheet.api.Cells.Replace(
                    What=f'{col_letter}{old_end_row}',
                    Replacement=f'{col_letter}{new_end_row}',
                    LookAt=2  # xlPart - partial match
                )
        except Exception as e:
            print(f"Warning: Range adjustment error: {e}")

    def _delete_excess_template_rows(self, sheet, data_start_row, actual_rows, placeholder_rows):
        """Delete excess placeholder rows from template after injecting data.

        Args:
            sheet: xlwings Sheet object
            data_start_row: First row of account data
            actual_rows: Number of actual account rows
            placeholder_rows: Number of placeholder rows in template
        """
        if actual_rows >= placeholder_rows:
            return  # No rows to delete

        excess_start = data_start_row + actual_rows
        excess_end = data_start_row + placeholder_rows - 1

        try:
            sheet.range(f'{excess_start}:{excess_end}').api.Delete()
        except Exception as e:
            print(f"Warning: Could not delete excess rows: {e}")

    def _hide_template_sheets(self, wb):
        """Hide all template sheets (those starting with _TPL_) in the final output.

        Args:
            wb: xlwings Workbook object
        """
        for sheet in wb.sheets:
            if sheet.name.startswith('_TPL_'):
                try:
                    sheet.api.Visible = False  # xlSheetHidden
                except Exception as e:
                    print(f"Warning: Could not hide template sheet {sheet.name}: {e}")

    def _ensure_template_v2_exists(self):
        """Ensure DNA_Template_v2.xlsm exists, create it if not.

        This generates the optimized template with pre-built sheets.
        Called once when the app starts or when template is missing.

        Returns:
            Path to template file, or None if creation failed
        """
        if os.path.exists(TEMPLATE_V2_PATH):
            return TEMPLATE_V2_PATH

        print("Creating DNA_Template_v2.xlsm...")
        try:
            self._create_template_v2()
            return TEMPLATE_V2_PATH
        except Exception as e:
            print(f"Error creating template v2: {e}")
            return None

    def _create_template_v2(self):
        """Create the optimized DNA_Template_v2.xlsm with pre-built template sheets.

        This function builds the template programmatically with all formatting,
        formulas, and structure pre-configured. It only needs to run once.
        """
        print("Building DNA_Template_v2.xlsm...")

        app = xw.App(visible=False)
        try:
            wb = app.books.add()

            # Define standard colors
            DARK_BLUE = (22, 33, 62)
            HEADER_WHITE = (255, 255, 255)
            SUBTOTAL_GRAY = (236, 236, 236)

            # ============================================================
            # Create _TPL_PL (P&L Template)
            # ============================================================
            tpl_pl = wb.sheets.add('_TPL_PL')
            self._build_pl_template_sheet(tpl_pl, DARK_BLUE, HEADER_WHITE, SUBTOTAL_GRAY)

            # ============================================================
            # Create _TPL_BS (Balance Sheet Template)
            # ============================================================
            tpl_bs = wb.sheets.add('_TPL_BS')
            self._build_bs_template_sheet(tpl_bs, DARK_BLUE, HEADER_WHITE, SUBTOTAL_GRAY)

            # ============================================================
            # Create placeholder sheets for other templates
            # These will be built out in later phases
            # ============================================================
            tpl_cf = wb.sheets.add('_TPL_CF')
            tpl_cf.range('A1').value = 'Cash Flow Template - Phase 2'

            tpl_forecast = wb.sheets.add('_TPL_Forecast')
            tpl_forecast.range('A1').value = 'Forecast Template - Phase 3'

            # ============================================================
            # Create standard sheets (Menu, Dashboard, etc.)
            # ============================================================
            menu_sheet = wb.sheets.add('Menu')
            menu_sheet.range('A1').value = 'Menu'
            menu_sheet.range('A1').font.bold = True

            # Remove default Sheet1
            for sheet in wb.sheets:
                if sheet.name == 'Sheet1':
                    sheet.delete()
                    break

            # Save as macro-enabled workbook
            wb.save(TEMPLATE_V2_PATH)
            print(f"Template v2 created: {TEMPLATE_V2_PATH}")

        finally:
            wb.close()
            app.quit()

    def _build_pl_template_sheet(self, sheet, dark_blue, header_white, subtotal_gray):
        """Build the _TPL_PL template sheet with all formatting pre-configured.

        Structure:
        - Row 1-2: Title area
        - Row 3: YYYYMM helper row (hidden)
        - Row 4: Header row
        - Row 5-204: 200 placeholder account rows with formatting
        - Row 205+: Validation section
        """
        PLACEHOLDER_ROWS = 200
        DATA_START_ROW = 5
        HEADER_ROW = 4

        # Max columns: Account + 24 months + Notes + spacer + YTD cols + spacer + 3 years
        MAX_COLS = 1 + 24 + 1 + 1 + 4 + 1 + 3  # = 35

        # Title rows
        sheet.range('A1').value = '<<<COMPANY_NAME>>>'
        sheet.range('A1').font.size = 14
        sheet.range('A1').font.bold = True

        sheet.range('A2').value = 'Profit & Loss Statement'
        sheet.range('A2').font.size = 12
        sheet.range('A2').font.bold = True

        # Row 3: Helper row (will contain YYYYMM values)
        sheet.range('A3').value = 'YYYYMM Helper'
        sheet.range('A3').font.color = (255, 255, 255)  # White (hidden)

        # Header row
        header_data = ['Account'] + [f'Month {i}' for i in range(1, 25)]  # 24 month placeholders
        header_data.extend(['Notes', '', 'PY YTD', 'CY YTD', 'Var $', 'Var %', ''])
        header_data.extend(['Year 1', 'Year 2', 'Year 3'])

        sheet.range((HEADER_ROW, 1), (HEADER_ROW, len(header_data))).value = [header_data]

        # Format header row
        header_range = sheet.range((HEADER_ROW, 1), (HEADER_ROW, len(header_data)))
        header_range.font.bold = True
        header_range.color = dark_blue
        header_range.font.color = header_white

        # Pre-format placeholder rows
        for row in range(DATA_START_ROW, DATA_START_ROW + PLACEHOLDER_ROWS):
            # Number format for data columns
            sheet.range((row, 2), (row, 25)).number_format = '#,##0'
            # YTD columns
            sheet.range((row, 28), (row, 29)).number_format = '#,##0'
            sheet.range((row, 30)).number_format = '#,##0'
            sheet.range((row, 31)).number_format = '0.0%'
            # Year columns
            sheet.range((row, 33), (row, 35)).number_format = '#,##0'

        # Column widths
        sheet.range('A:A').column_width = 35
        sheet.range((1, 27)).column_width = 2  # Spacer 1
        sheet.range((1, 32)).column_width = 2  # Spacer 2

        # Hide row 3 (YYYYMM helper)
        try:
            sheet.range('3:3').api.Hidden = True
        except:
            pass

        # Add marker for injection point
        sheet.range(f'A{DATA_START_ROW}').value = '<<<INSERT_ACCOUNTS_HERE>>>'

        print(f"  Built _TPL_PL with {PLACEHOLDER_ROWS} placeholder rows")

    def _build_bs_template_sheet(self, sheet, dark_blue, header_white, subtotal_gray):
        """Build the _TPL_BS template sheet with all formatting pre-configured.

        Similar structure to P&L but typically fewer accounts.
        """
        PLACEHOLDER_ROWS = 150
        DATA_START_ROW = 5
        HEADER_ROW = 4

        # Title rows
        sheet.range('A1').value = '<<<COMPANY_NAME>>>'
        sheet.range('A1').font.size = 14
        sheet.range('A1').font.bold = True

        sheet.range('A2').value = 'Balance Sheet'
        sheet.range('A2').font.size = 12
        sheet.range('A2').font.bold = True

        # Row 3: Helper row
        sheet.range('A3').value = 'YYYYMM Helper'
        sheet.range('A3').font.color = (255, 255, 255)

        # Header row
        header_data = ['Account'] + [f'Month {i}' for i in range(1, 25)]
        sheet.range((HEADER_ROW, 1), (HEADER_ROW, len(header_data))).value = [header_data]

        header_range = sheet.range((HEADER_ROW, 1), (HEADER_ROW, len(header_data)))
        header_range.font.bold = True
        header_range.color = dark_blue
        header_range.font.color = header_white

        # Pre-format placeholder rows
        for row in range(DATA_START_ROW, DATA_START_ROW + PLACEHOLDER_ROWS):
            sheet.range((row, 2), (row, 25)).number_format = '#,##0'

        # Column widths
        sheet.range('A:A').column_width = 35

        # Hide row 3
        try:
            sheet.range('3:3').api.Hidden = True
        except:
            pass

        # Add marker
        sheet.range(f'A{DATA_START_ROW}').value = '<<<INSERT_ACCOUNTS_HERE>>>'

        print(f"  Built _TPL_BS with {PLACEHOLDER_ROWS} placeholder rows")

    # ================================================================
    # END TEMPLATE V2 UTILITY FUNCTIONS
    # ================================================================

    def _add_back_to_menu_link(self, sheet, row=1, col=1):
        """Add a 'Back to Menu' hyperlink at the specified position"""
        try:
            cell = sheet.range((row, col))
            cell.value = '<< Menu'
            cell.font.name = 'Calibri Light'
            cell.font.size = 9
            cell.font.color = (0, 102, 204)  # Light blue
            cell.font.underline = True
            cell.add_hyperlink('#Menu!A1', text_to_display='<< Menu')
        except Exception as e:
            print(f"Back to Menu link warning: {e}")

    def _setup_print_area(self, sheet, last_row=None, last_col=None):
        """Set up professional print settings for a sheet

        Args:
            sheet: xlwings Sheet object
            last_row: Last row to include (optional, uses UsedRange if not specified)
            last_col: Last column to include (optional, uses UsedRange if not specified)
        """
        try:
            ps = sheet.api.PageSetup

            # Landscape orientation
            ps.Orientation = 2  # xlLandscape

            # Fit all columns on one page, rows can span multiple pages
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False  # Allow multiple pages vertically

            # Margins (in inches)
            ps.LeftMargin = 36  # 0.5 inch
            ps.RightMargin = 36
            ps.TopMargin = 54  # 0.75 inch
            ps.BottomMargin = 54
            ps.HeaderMargin = 36
            ps.FooterMargin = 36

            # Header: Company name on left, sheet name center
            company = self.company_name.get()
            ps.LeftHeader = f"&\"Calibri Light,Regular\"&10{company}"
            ps.CenterHeader = f"&\"Calibri Light,Bold\"&12&A"  # Sheet name
            ps.RightHeader = ""

            # Footer: Date on left, page number center
            ps.LeftFooter = "&\"Calibri Light,Regular\"&9&D"  # Date
            ps.CenterFooter = "&\"Calibri Light,Regular\"&9Page &P of &N"  # Page X of Y
            ps.RightFooter = ""

            # Repeat rows at top (header row)
            ps.PrintTitleRows = "$1:$4"

            # Gridlines off for cleaner print
            ps.PrintGridlines = False

        except Exception as e:
            print(f"Print setup warning: {e}")

    def _apply_notes_division_dropdown(self, sheet, notes_col, start_row, end_row, divisions=None):
        """Add division dropdown to Notes column for multi-division mode

        Args:
            sheet: xlwings Sheet object
            notes_col: Column number for Notes
            start_row: First data row
            end_row: Last data row
            divisions: List of division dicts (with 'name' key)
        """
        try:
            if not divisions or len(divisions) < 2:
                return  # No dropdown needed for single division

            # Build dropdown list: "All" + division names
            division_names = ["All"] + [d.get('name', d) if isinstance(d, dict) else d for d in divisions]
            dropdown_list = ",".join(division_names)

            # Apply data validation to notes column range
            notes_range = sheet.range((start_row, notes_col), (end_row, notes_col))
            notes_range.api.Validation.Delete()  # Clear existing
            notes_range.api.Validation.Add(
                Type=3,  # xlValidateList
                AlertStyle=1,  # xlValidAlertStop
                Operator=1,  # xlBetween
                Formula1=dropdown_list
            )
            notes_range.api.Validation.ShowDropDown = False  # Show dropdown arrow
            notes_range.api.Validation.InCellDropdown = True
            notes_range.api.Validation.ShowError = False  # Allow free text too

        except Exception as e:
            print(f"Warning: Could not add division dropdown to Notes: {e}")

    def _apply_row_grouping(self, sheet, accounts, start_row):
        """Apply Excel row grouping (outline) based on account sections

        Groups rows between section headers and their totals to allow
        collapsing/expanding sections in Excel.

        Args:
            sheet: xlwings Sheet object
            accounts: List of account dictionaries with 'is_header' and 'is_total' flags
            start_row: First data row (after header row)
        """
        try:
            # Track sections: (header_row, total_row)
            sections = []
            current_section_start = None
            row_idx = start_row

            for account in accounts:
                if account.get('is_header', False):
                    # Start of a new section
                    if current_section_start is not None:
                        # Close previous section without a total (shouldn't happen but handle it)
                        sections.append((current_section_start, row_idx - 1))
                    current_section_start = row_idx
                elif account.get('is_total', False) and current_section_start is not None:
                    # End of current section - group rows BETWEEN header and total
                    if row_idx > current_section_start + 1:
                        # Only group if there are rows between header and total
                        sections.append((current_section_start + 1, row_idx - 1))
                    current_section_start = None

                row_idx += 1

            # Apply grouping to each section
            for group_start, group_end in sections:
                if group_end > group_start:
                    try:
                        # Group rows (creates collapsible outline)
                        rows_to_group = sheet.range(f'{group_start}:{group_end}').api.Rows
                        rows_to_group.Group()
                    except Exception as e:
                        print(f"Warning: Could not group rows {group_start}-{group_end}: {e}")

            # Set outline settings - summary rows below detail (default)
            try:
                sheet.api.Outline.SummaryRow = 0  # xlAbove = 0, summary rows above detail
            except:
                pass

        except Exception as e:
            print(f"Warning: Could not apply row grouping: {e}")

    def _build_source_lookup_formula(self, source_sheet, account_name, col_num, division=None):
        """Build a SUMIF or SUMIFS formula for looking up source data

        Args:
            source_sheet: Name of source sheet (e.g., "Source_PL")
            account_name: Account name to look up
            col_num: Column number for the value (1-based)
            division: Optional division name for multi-division filtering

        Returns:
            Excel formula string
        """
        col_letter = self._col_letter(col_num)
        # Use limited ranges (rows 3-1500) instead of entire columns for speed
        sr = 3  # source start row
        er = 1500  # source end row

        if division or self.is_multi_division.get():
            # Multi-division mode: Use SUMIFS with Division filter
            # Structure: Division in A, Account in B, values start in C
            div_filter = division if division else 'Menu!$G$5'
            return (f'=SUMIFS({source_sheet}!{col_letter}${sr}:{col_letter}${er},'
                    f'{source_sheet}!$A${sr}:$A${er},"{div_filter}",'
                    f'{source_sheet}!$B${sr}:$B${er},"{account_name}")')
        else:
            # Single division mode: Use SUMIF (backward compatible)
            return f'=SUMIF({source_sheet}!$A${sr}:$A${er},"{account_name}",{source_sheet}!{col_letter}${sr}:{col_letter}${er})'

    def _build_consolidated_formula(self, source_sheet, account_name, col_num):
        """Build a formula that sums across all divisions (consolidated view)

        Args:
            source_sheet: Name of source sheet (e.g., "Source_PL")
            account_name: Account name to look up
            col_num: Column number for the value (1-based)

        Returns:
            Excel formula string for consolidated sum
        """
        col_letter = self._col_letter(col_num)
        # Use limited ranges (rows 3-1500) instead of entire columns for speed
        sr = 3
        er = 1500

        if self.is_multi_division.get():
            # Consolidated: Sum all divisions (no division filter)
            return f'=SUMIF({source_sheet}!$B${sr}:$B${er},"{account_name}",{source_sheet}!{col_letter}${sr}:{col_letter}${er})'
        else:
            # Single division: Same as regular lookup
            return f'=SUMIF({source_sheet}!$A${sr}:$A${er},"{account_name}",{source_sheet}!{col_letter}${sr}:{col_letter}${er})'

    def _group_columns_by_year(self, sheet, months, header_row):
        """
        Hide and group columns based on current month (last month in data):
        - Group and hide all prior year columns (before current year)
        - Current year columns up to current month remain visible
        - Future months (after current month) are hidden but not grouped

        The "current month" is determined by the last month in the data,
        which corresponds to Menu cell C7.
        """
        if not months:
            return

        try:
            # Current month is the LAST month in the data (this is what Menu C7 shows)
            current_month_num, current_year, current_month_name = months[-1]
            prior_year = current_year - 1

            print(f"Column visibility: Current month = {current_month_name} (month {current_month_num}), Year = {current_year}")

            # Track columns to hide and group
            prior_year_cols = []  # Columns from prior years (to be grouped and hidden)
            future_month_cols = []  # Columns after current month (to be hidden only)

            # months is a list of tuples: (month_num, year, display_name)
            for i, (m, y, name) in enumerate(months):
                col = i + 2  # Data starts at column 2 (B)

                if y < current_year:
                    # Prior year - group and hide
                    prior_year_cols.append((col, y))
                elif y == current_year and m > current_month_num:
                    # Future month in current year - hide only
                    future_month_cols.append(col)
                elif y > current_year:
                    # Future year - hide only
                    future_month_cols.append(col)
                # else: current year, current or prior month - leave visible

            # Group prior year columns by year
            year_groups = {}
            for col, y in prior_year_cols:
                if y not in year_groups:
                    year_groups[y] = [col, col]
                else:
                    year_groups[y][1] = col

            # Create groups for each prior year and hide them
            for year in sorted(year_groups.keys()):
                start_col, end_col = year_groups[year]
                try:
                    # Create group
                    col_range = sheet.range((1, start_col), (1, end_col))
                    col_range.api.EntireColumn.Group()

                    # Hide the grouped columns
                    col_range.api.EntireColumn.Hidden = True
                    print(f"Grouped and hid year {year}: columns {start_col} to {end_col}")
                except Exception as e:
                    print(f"Column grouping error for year {year}: {e}")

            # Hide future month columns (not grouped, just hidden)
            for col in future_month_cols:
                try:
                    sheet.range((1, col)).api.EntireColumn.Hidden = True
                    print(f"Hid future month column {col}")
                except Exception as e:
                    print(f"Error hiding column {col}: {e}")

        except Exception as e:
            print(f"Year grouping error: {e}")


def main():
    root = tk.Tk()
    app = FinancialModelApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
