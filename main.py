from flask import Flask, render_template, request, send_file
import os
import pdfplumber
import arabic_reshaper
import xlsxwriter
from bidi.algorithm import get_display

app = Flask(__name__)

uploaded_file_paths = []

app.config['UPLOAD_FOLDER'] = 'process'  # Set the upload folder
print('***************************')
print(os.listdir())


@app.route('/')
def index():
  return render_template('upload.html')


@app.route('/upload', methods=['POST'])
def upload_file():
  global uploaded_file_paths
  uploaded_file_paths = []
  uploaded_files = request.files.getlist('files[]')
  for uploaded_file in uploaded_files:
    if uploaded_file.filename != '':
      uploaded_file.save(
          os.path.join(app.config['UPLOAD_FOLDER'], uploaded_file.filename))
      uploaded_file_paths.append(
          os.path.join(app.config['UPLOAD_FOLDER'], uploaded_file.filename))
  process_files()
  return send_file(os.path.join(app.config['UPLOAD_FOLDER'], 'sheet.xlsx'),
                   as_attachment=True), delete()


@app.route('/process')
def process_files():
  # Add your processing logic here
  char_to_replace = {
      'PM':
      '',
      'Visa No. ':
      '',
      'Valid from ':
      '',
      'ﺭﻗﻢ ﺍﻟﺘﺄﺷﻴﺮﺓ  ':
      '',
      'ﺻﺎﻟﺤﺔ ﺍﻋﺘﺒﺎﺭﺍ ﻣﻦ':
      '',
      'Valid until ':
      '',
      'ﺻﺎﻟﺤﺔ ﻟﻐﺎﻳﺔ':
      '',
      'ﻣﺪﺓ ﺍﻹﻗﺎﻣﺔ ﻳﻮﻡ Duration of Stay Days ':
      '',
      'Passport No. ':
      '',
      'ﺭﻗﻢ ﺟﻮﺍﺯ ﺍﻟﺴﻔﺮ ':
      '',
      'ﻣﺼﺪﺭ ﺍﻟﺘﺄﺷﻴﺮﺓ ﺍﻟﺴﻔﺎﺭﺓ ﺍﻟﺴﻌﻮﺩﻳﺔ ﺍﻟﺮﻗﻤﻴﺔ - Place of issue Saudi Digital Embassy':
      '',
      'Name':
      '',
      'ﺍﻻﺳﻢ':
      '',
      'ﺍﻟﺠﻨﺴﻴﺔ':
      '',
      'ﻧﻮﻉ ﺍﻟﺘﺄﺷﻴﺮﺓ ﺯﻳﺎﺭﺓ ﺣﻜﻮﻣﻴﺔ - Type Of Visa Gov. Visit':
      '',
      ' Nationality ':
      '',
      ' Entry Type ':
      '',
      'ﻋﺪﺩ ﻣﺮﺎﺗ ﺍﻟﺪﺧﻮﻝ ':
      '',
      'ﺍﻟﻐﺮﺽ ﻣﻮﺳﻢ ﺍﻟﺮﻳﺎﺽ - Purpose Riyadh Season':
      '',
      '.Visa No':
      '',
      'ﺭﻗﻢ ﺍﻟﺴﺠﻞ':
      '',
      'https://visa.mofa.gov.sa/Home/PrintEventVisa Page 1 of 2':
      '',
      '.Application No':
      '',
      'ﺭﻗﻢ ﺍﻟﻄﻠﺐ  ':
      '',
      'Nationality ':
      '',
      'Duration of Stay Days':
      '',
      'ﻣﺪﺓ ﺍﻹﻗﺎﻣﺔ ٩٠ ﻳﻮﻡ':
      '',
      'ﻣﺪﺓ ﺍﻹﻗﺎﻣﺔ':
      '',
      'Days':
      '',
      'Entry Type ':
      '',
      '-':
      '',
      'PassportNo.':
      '',
      'VisaNo.':
      '',
      'Validfrom':
      '',
      'Validuntil':
      '',
      'Duration of Stay':
      '',
      '٣٠':
      '',
      'ﻳﻮﻡ':
      '',
      'ﺭﻗﻢﺍﻟﺘﺄﺷﻴﺮﺓ':
      '',
      'ﺭﻗﻢ ﺍﻟﺠﻮﺍﺯ':
      '',
      'ﺭﻗﻢ ﺍﻟﺘﺄﺷﻴﺮﺓ ': '',
      'ﺻﺎﻟﺤﺔ ﺍﻋﺘﺒﺎﺭ ﺍ ﻣﻦ Valid From ': '',
      ' Valid Until ': ''
  }
  count = len(uploaded_file_paths)

  workbook = xlsxwriter.Workbook(
      os.path.join(app.config['UPLOAD_FOLDER']) + '/sheet.xlsx')
  worksheet = workbook.add_worksheet()
  n = 1
  while n <= count:
    for filename in uploaded_file_paths:
      x1 = 1
      x3 = 2
      x5 = 3
      x7 = 4
      x8 = 5
      x12 = 8
      x14 = 11
      file = filename
      pdf = pdfplumber.open(file)
      page = pdf.pages[0]
      text = page.extract_text()
      reshaped_text = arabic_reshaper.reshape(text)
      bidi_text = get_display(reshaped_text)
      for key, value in char_to_replace.items():
        bidi_text = bidi_text.replace(key, value)

      Visano = bidi_text.splitlines()[1]
      print("Visa no:", Visano)
      if '6' not in Visano:
        x1 = x1 - 1
        x3 = x3 - 1
        x5 = x5 - 1
        x7 = x7 - 1
        x8 = x8 - 1
        x12 = x12 - 2
        x14 = x14 - 2
      else:
        print('all good')

      print(bidi_text)
      Visano = bidi_text.splitlines()[x1]
      print("Visa no:", Visano)
      start = bidi_text.splitlines()[x3]
      end = bidi_text.splitlines()[x5]
      passport = bidi_text.splitlines()[x8]
      duration = bidi_text.splitlines()[x7]
      country = bidi_text.splitlines()[x12]
      if country == "":
        country = bidi_text.splitlines()[x12 - 1]
      vtype = bidi_text.splitlines()[x14]
      if vtype == "":
        vtype = bidi_text.splitlines()[x14 - 1]
      country = country.encode('ascii', 'ignore').decode()
      vtype = vtype.encode('ascii', 'ignore').decode()
      #vtype = vtype.encode('ascii', 'ignore', **('errors',)).decode()
      name = filename.split('/')[len(filename.split('/')) - 1]
      name = name.replace('.pdf', '')
      vtype = vtype.replace(' ', '')
      Visano = Visano.replace(' ', '')
      country = country.lstrip()
      duration = duration.replace(' ', '')
      worksheet.write('A' + str(n), name)
      worksheet.write('B' + str(n), country)
      worksheet.write('C' + str(n), passport)
      worksheet.write('D' + str(n), Visano)
      worksheet.write('E' + str(n), start)
      worksheet.write('F' + str(n), end)
      worksheet.write('G' + str(n), duration)
      worksheet.write('H' + str(n), vtype)
      n = n + 1

  workbook.close()


  # Delete uploaded files
def delete():
  for filename in os.listdir(app.config['UPLOAD_FOLDER']):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.isfile(file_path):
      os.remove(file_path)


if __name__ == '__main__':
  os.makedirs(app.config['UPLOAD_FOLDER'],
              exist_ok=True)  # Create 'process' folder if it doesn't exist
  app.run(debug=True, host='0.0.0.0', port=8080)
