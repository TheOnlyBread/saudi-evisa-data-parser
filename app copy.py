from flask import Flask, render_template_string, request, send_file
from pyngrok import ngrok
import os
import re
import pdfplumber
import arabic_reshaper
import xlsxwriter
from bidi.algorithm import get_display

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

with open("upload.html", "r", encoding="utf-8") as f:
    upload_html = f.read()

@app.route("/")
def index():
    return render_template_string(upload_html)

@app.route("/upload", methods=["POST"])
def upload_file():
    uploaded_files = request.files.getlist("files[]")
    saved_paths = []
    for uploaded_file in uploaded_files:
        if uploaded_file and uploaded_file.filename:
            path = os.path.join(UPLOAD_FOLDER, uploaded_file.filename)
            uploaded_file.save(path)
            saved_paths.append(path)

    # Generate Excel
    out_path = os.path.join(UPLOAD_FOLDER, 'sheet.xlsx')
    workbook = xlsxwriter.Workbook(out_path)
    worksheet = workbook.add_worksheet()
    headers = ["Name", "Country", "Passport Number", "Visa No", "Valid From", "Valid Until", "Duration of Stay", "Entry Type"]
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    for row_num, file_path in enumerate(saved_paths, 1):
        with pdfplumber.open(file_path) as pdf:
            text = pdf.pages[0].extract_text() or ""
        reshaped = arabic_reshaper.reshape(text)
        bidi_text = get_display(reshaped)
        name = os.path.splitext(os.path.basename(file_path))[0]
        country = re.search(r'Nationality\s+([A-Za-z ]+)', bidi_text)
        passport = re.search(r'(?:Passport No\.|PassportNo\.|رقم الجواز)\s*([A-Z0-9]+)', bidi_text)
        valid_from = re.search(r'Valid From\s+(\d{2}/\d{2}/\d{4})', bidi_text)
        valid_until = re.search(r'Valid Until\s+(\d{2}/\d{2}/\d{4})', bidi_text)
        duration = None
        for line in bidi_text.splitlines():
            if "Duration of Stay" in line:
                match = re.search(r'(\d+|[٠-٩]+)', line)
                if match:
                    raw = match.group(1)
                    arabic_digits = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'}
                    duration = ''.join(arabic_digits.get(c, c) for c in raw)
                break
        entry_type = re.search(r'Entry Type\s+(Single|Multiple)', bidi_text)
        visa_no = re.search(r'Visa No\.\s+(\d+)', bidi_text)

        row = [
            name,
            country.group(1).strip() if country else None,
            passport.group(1).strip() if passport else None,
            visa_no.group(1).strip() if visa_no else None,
            valid_from.group(1).strip() if valid_from else None,
            valid_until.group(1).strip() if valid_until else None,
            duration,
            entry_type.group(1).strip() if entry_type else None
        ]
        for col_num, item in enumerate(row):
            worksheet.write(row_num, col_num, item)

    workbook.close()
    return send_file(out_path, as_attachment=True)

# Start server via ngrok
url = ngrok.connect(5000)
print(f"Public URL: {url}")
app.run(port=5000)
