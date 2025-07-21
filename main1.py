from flask import Flask, render_template, request, send_file
import os
import re
import pdfplumber
import arabic_reshaper
import xlsxwriter
from bidi.algorithm import get_display

app = Flask(__name__)

# Configure upload folder
app.config['UPLOAD_FOLDER'] = 'process'

# Keep track of uploaded file paths
django = []  # placeholder to avoid name clash
uploaded_file_paths = []

def clear_process_folder():
    """
    Delete all files in the process folder before processing new uploads.
    """
    folder = app.config['UPLOAD_FOLDER']
    if os.path.isdir(folder):
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            if os.path.isfile(file_path):
                try:
                    os.remove(file_path)
                except Exception as e:
                    app.logger.error(f"Error deleting file {file_path}: {e}")

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # Before handling new uploads, clear out any existing files
    clear_process_folder()

    global uploaded_file_paths
    uploaded_file_paths = []

    uploaded_files = request.files.getlist('files[]')
    for uploaded_file in uploaded_files:
        if uploaded_file and uploaded_file.filename:
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], uploaded_file.filename)
            uploaded_file.save(file_path)
            uploaded_file_paths.append(file_path)

    # Process the newly uploaded files and generate sheet.xlsx
    process_files()

    # Serve the generated Excel and then clear the folder again
    response = send_file(os.path.join(app.config['UPLOAD_FOLDER'], 'sheet.xlsx'), as_attachment=True)
    clear_process_folder()
    return response


def extract_visa_info_from_text(text, file_name):
    # Extract the name directly from the file name
    name = os.path.splitext(os.path.basename(file_name))[0]

    # Nationality / Country
    country_match = re.search(r'Nationality\s+([A-Za-z ]+)', text)
    country = country_match.group(1).strip() if country_match else None

    # Passport Number (alphanumeric)
    passport_match = re.search(r'(?:Passport No\.|PassportNo\.|رقم الجواز)\s*([A-Z0-9]+)', text)
    passport_number = passport_match.group(1).strip() if passport_match else None

    # Valid From / Until
    valid_from_match = re.search(r'Valid From\s+(\d{2}/\d{2}/\d{4})', text)
    valid_from = valid_from_match.group(1).strip() if valid_from_match else None
    valid_until_match = re.search(r'Valid Until\s+(\d{2}/\d{2}/\d{4})', text)
    valid_until = valid_until_match.group(1).strip() if valid_until_match else None

    # Duration of Stay (handles Arabic numerals too)
    duration_of_stay = None
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "Duration of Stay" in line:
            duration_line = line + (lines[i+1] if i+1 < len(lines) else "")
            duration_match = re.search(r'(\d+|[٠-٩]+)', duration_line)
            if duration_match:
                raw = duration_match.group(1)
                arabic_to_english = {
                    '٠':'0','١':'1','٢':'2','٣':'3','٤':'4',
                    '٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'
                }
                duration_of_stay = ''.join(arabic_to_english.get(c, c) for c in raw)
            break

    # Entry Type
    entry_type_match = re.search(r'Entry Type\s+(Single|Multiple)', text)
    entry_type = entry_type_match.group(1).strip() if entry_type_match else None

    # Visa No
    visa_no_match = re.search(r'Visa No\.\s+(\d+)', text)
    visa_no = visa_no_match.group(1).strip() if visa_no_match else None

    return {
        "Name": name,
        "Country": country,
        "Passport Number": passport_number,
        "Visa No": visa_no,
        "Valid From": valid_from,
        "Valid Until": valid_until,
        "Duration of Stay": duration_of_stay,
        "Entry Type": entry_type
    }


def process_files():
    # Create an Excel workbook and worksheet
    out_path = os.path.join(app.config['UPLOAD_FOLDER'], 'sheet.xlsx')
    workbook = xlsxwriter.Workbook(out_path)
    worksheet = workbook.add_worksheet()

    headers = [
        "Name", "Country", "Passport Number", "Visa No",
        "Valid From", "Valid Until", "Duration of Stay", "Entry Type"
    ]
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    # Fill rows
    for row_num, file_path in enumerate(uploaded_file_paths, start=1):
        with pdfplumber.open(file_path) as pdf:
            first_page = pdf.pages[0]
            extracted_text = first_page.extract_text() or ""

        # Handle Arabic shaping and BIDI
        reshaped = arabic_reshaper.reshape(extracted_text)
        bidi_text = get_display(reshaped)

        visa_info = extract_visa_info_from_text(bidi_text, file_path)

        for col_num, key in enumerate(headers):
            worksheet.write(row_num, col_num, visa_info.get(key))

    workbook.close()


if __name__ == '__main__':
    # Ensure process folder exists
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True, host='0.0.0.0', port=8080)
