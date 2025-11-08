import os
import sys
import pandas as pd
from flask import Flask, render_template, request, send_file, redirect, url_for
import zipfile
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from num2words import num2words
from io import BytesIO
def generate_pdf_filename(data):
    """Generate unique filename using vendor name and code."""
    vendor_name = data['vendor_name'].replace(' ', '_')
    vendor_code = data['vendor_code']
    return f"{vendor_name}_{vendor_code}.pdf"

# If running from PyInstaller bundle
if getattr(sys, 'frozen', False):
    os.environ['PATH'] = sys._MEIPASS + os.pathsep + os.environ['PATH']


base_path = os.environ.get("FLASK_APP_BASE", os.path.abspath("."))


app = Flask(__name__, template_folder=os.path.join(base_path, "templates"))





UPLOAD_FOLDER = '/tmp/uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def load_excel(file_path):
    company_info = pd.read_excel(file_path, header=None, nrows=4)
    df = pd.read_excel(file_path, dtype=str, header=3)
    df.columns = df.columns.map(str)
    return company_info, df

def find_column(columns, search):
    search = search.lower().replace(" ", "")
    for col in columns:
        col_clean = col.lower().replace(" ", "")
        if search in col_clean:
            return col
    return None

def format_joining_date(date_str):
    try:
        dt = pd.to_datetime(date_str)
        return dt.strftime('%d %B %Y')
    except Exception:
        return date_str or ""

def net_payment_words_inr(amount):
    try:
        n = float(str(amount).replace(",", "")) if amount else 0
        words = num2words(n, lang='en_IN').replace(' and', '').replace('-', ' ')
        words = words.replace(',', '')
        words = words.title()
        if words == "Zero":
            return "Rupees Zero Only"
        return f"Rupees {words} Only"
    except Exception:
        return "Rupees Zero Only"

def standardize(row, columns, month_str, month_col, net_payment_col, total_payable_col, company_name, company_address):
    def get_val(key):
        col = find_column(columns, key)
        return row.get(col, "") if col else ""

    total_fee = row.get(month_col, "")
    net_payment = row.get(net_payment_col, "")
    total_payable = row.get(total_payable_col, "")
    contract_start_date = format_joining_date(get_val("Contract Start Date"))
    contract_end_date = format_joining_date(get_val("Contract End Date"))
    return {
        "company_name": company_name,
        "company_address": company_address,
        "month": month_str,
        "vendor_name": get_val("Vendor's Name"),
        "contract_start_date": contract_start_date,
        "contract_end_date": contract_end_date,
        "location": get_val("Location"),
        "pay_days": get_val("Pay Days"),
        "LOP_days": get_val("LOP days") or 0,
        "vendor_code": get_val("Vendor's Code"),
        "bank_name": get_val("Consultant’s  Bank Name"),
        "bank_account": get_val("Consultant’s  Bank A/c No."),
        "pan_no": get_val("PAN No."),
        "monthly_fee": get_val("Monthly Fee"),
        "total_fee": total_fee,
        "bonus": get_val("Incentive/Bonus") or 0,
        "travel_reimbursement": get_val("Travel Reimbursement") or 0,
        "new_area_allowance": get_val("New Area Allowance") or 0,
        "tds_10": get_val("TDS@10%"),
        "other_deduction": get_val("Other Deduction") or 0,
        "financial_pendency": get_val("Financial Pendency") or 0,
        "advance_recovery": get_val("Advance Recovery") or 0,
        "total_gross": get_val("Total Gross"),
        "total_deductions": get_val("Total Deduction"),
        "net_payment": net_payment,
        "net_payment_words": net_payment_words_inr(net_payment),
        "total_payable": total_payable
    }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    uploaded_file = request.files['excel_file']
    if uploaded_file.filename == '':
        return 'No file selected.'

    file_path = os.path.join(UPLOAD_FOLDER, uploaded_file.filename)
    uploaded_file.save(file_path)

    company_info, df = load_excel(file_path)
    columns = list(df.columns)

    company_name = str(company_info.iloc[0, 0]) if not company_info.empty else ""
    company_address = "<br>".join([
        str(company_info.iloc[i, 0]) for i in range(1, company_info.shape[0])
        if pd.notna(company_info.iloc[i, 0])
        and str(company_info.iloc[i, 0]).strip()
        and "consultants pay-out sheet" not in str(company_info.iloc[i, 0]).lower()
        and "s. no." not in str(company_info.iloc[i, 0]).lower()
    ])

    month_col = find_column(columns, "Total Fee in")
    net_payment_col = find_column(columns, "Net Payment for")
    total_payable_col = find_column(columns, "Total Payable in")

    if not month_col or not net_payment_col or not total_payable_col:
        return 'Required dynamic columns not found in Excel.'

    month_str = month_col.split("in")[-1].strip().replace("'", "")
    if len(month_str) > 2 and not month_str[-3] == "'":
        month_str = month_str[:-2] + "'" + month_str[-2:]

    # Create OUTPUT_FOLDER dynamically based on month
    OUTPUT_FOLDER = f"/tmp/generated_pdfs_{month_str.replace(' ', '_')}"
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # Clean folder
    for f in os.listdir(OUTPUT_FOLDER):
        os.remove(os.path.join(OUTPUT_FOLDER, f))

    env = Environment(loader=FileSystemLoader(os.path.join(base_path, "templates")))

    template = env.get_template('payslip_template.html')

    vendor_counter = 1
    for index, row in df.iterrows():
        filled_data = standardize(
            row, columns, month_str, month_col, net_payment_col,
            total_payable_col, company_name, company_address
        )

        vendor_name = str(filled_data['vendor_name']).strip()
        vendor_name_lower = vendor_name.lower()
        if not vendor_name or vendor_name_lower in ["total", "nan"]:
            continue

        # ✅ Numbered log in terminal
        print(f"{vendor_counter}: {vendor_name.upper()}")
        vendor_counter += 1

        html_out = template.render(data=filled_data)
        file_name = f"{vendor_name.replace(' ', '_')}_{filled_data['vendor_code']}_{month_str}.pdf"
        output_path = os.path.join(OUTPUT_FOLDER, file_name)
        HTML(string=html_out).write_pdf(output_path)

    # Create ZIP file
    zip_path = f"/tmp/Payslips_{month_str.replace(' ', '_')}.zip"
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for filename in os.listdir(OUTPUT_FOLDER):
            file_path = os.path.join(OUTPUT_FOLDER, filename)
            zipf.write(file_path, arcname=filename)

    return redirect(url_for('success'))

@app.route('/success')
def success():
    return render_template('success.html')

@app.route('/download')
def download_zip():
    import glob
    zips = sorted(glob.glob("/tmp/Payslips_*.zip"), key=os.path.getmtime, reverse=True)
    if not zips:
        return "No zip file found."
    return send_file(zips[0], as_attachment=True)

if __name__ == "__main__":
    import sys
    # Disable auto-reloader when packaged into .exe
    use_reloader = "--no-reload" not in sys.argv
    app.run(debug=False, use_reloader=use_reloader)









