"""
CFO Financial Model Generator
Flask web application that generates macro-enabled Excel financial models
"""

import os
import io
import re
import zipfile
from datetime import datetime
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import pandas as pd
from excel_generator import FinancialModelGenerator

app = Flask(__name__)
app.secret_key = 'dna-model-secret-key-change-in-production'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

# Path to VBA code file
VBA_CODE_PATH = os.path.join(os.path.dirname(__file__), 'vba_code.bas')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_months():
    return [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ]

def get_years():
    current_year = datetime.now().year
    return list(range(current_year - 5, current_year + 2))

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Validate files
        if 'pl_file' not in request.files or 'bs_file' not in request.files:
            flash('Please upload both P&L and Balance Sheet files.', 'error')
            return redirect(request.url)

        pl_file = request.files['pl_file']
        bs_file = request.files['bs_file']

        if pl_file.filename == '' or bs_file.filename == '':
            flash('Please select both files.', 'error')
            return redirect(request.url)

        if not (allowed_file(pl_file.filename) and allowed_file(bs_file.filename)):
            flash('Invalid file type. Please upload Excel (.xlsx, .xls) or CSV files.', 'error')
            return redirect(request.url)

        # Get configuration
        company_name = request.form.get('company_name', 'Company').strip()
        if not company_name:
            company_name = 'Company'

        fiscal_year_start = int(request.form.get('fiscal_year_start', 1))
        first_display_month = int(request.form.get('first_display_month', 1))
        first_display_year = int(request.form.get('first_display_year', datetime.now().year))

        try:
            # Read uploaded files
            pl_data = read_financial_file(pl_file)
            bs_data = read_financial_file(bs_file)

            # Generate Financial model
            generator = FinancialModelGenerator(
                company_name=company_name,
                fiscal_year_start=fiscal_year_start,
                first_display_month=first_display_month,
                first_display_year=first_display_year,
                pl_data=pl_data,
                bs_data=bs_data
            )

            excel_buffer = generator.generate()

            # Create safe filename
            safe_company_name = re.sub(r'[^\w\s-]', '', company_name).strip().replace(' ', '_')

            # Create a ZIP file containing both the Excel and VBA files
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                # Add Excel file
                excel_buffer.seek(0)
                zf.writestr(f"{safe_company_name}_Financial_Model.xlsx", excel_buffer.read())

                # Add VBA code file
                if os.path.exists(VBA_CODE_PATH):
                    with open(VBA_CODE_PATH, 'r') as vba_file:
                        zf.writestr('vba_code.bas', vba_file.read())

                # Add README
                readme_content = f"""
{company_name} - Financial Model
====================================

This package contains:
1. {safe_company_name}_Financial_Model.xlsx - Your financial model
2. vba_code.bas - VBA macros for full functionality

SETUP INSTRUCTIONS:
-------------------
1. Open {safe_company_name}_Financial_Model.xlsx in Excel
2. Save As "Excel Macro-Enabled Workbook (*.xlsm)"
3. Press Alt+F11 to open VBA Editor
4. File > Import File > Select vba_code.bas
5. Close VBA Editor and save
6. Enable macros when prompted

The model includes:
- P&L Statement with VLOOKUP formulas
- Balance Sheet with VLOOKUP formulas
- Cash Flow Statement structure
- Variance notes functionality
- Diagnostics and validation

To upload new monthly data:
- Press Alt+F8
- Run "UploadPLFile" or "UploadBSFile"

Generated: {datetime.now().strftime('%B %d, %Y at %H:%M')}
"""
                zf.writestr('README.txt', readme_content)

            zip_buffer.seek(0)

            return send_file(
                zip_buffer,
                mimetype='application/zip',
                as_attachment=True,
                download_name=f"{safe_company_name}_Financial_Model.zip"
            )

        except Exception as e:
            flash(f'Error generating model: {str(e)}', 'error')
            import traceback
            traceback.print_exc()
            return redirect(request.url)

    return render_template('index.html',
                         months=get_months(),
                         years=get_years(),
                         current_year=datetime.now().year)

def read_financial_file(file):
    """Read a financial file (Excel or CSV) and return parsed data"""
    filename = secure_filename(file.filename)
    ext = filename.rsplit('.', 1)[1].lower()

    if ext == 'csv':
        df = pd.read_csv(file, header=None)
    else:
        df = pd.read_excel(file, header=None)

    return df

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True, port=5000)
