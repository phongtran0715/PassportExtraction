import pytesseract
from passporteye import read_mrz
import os
import sys
import pandas as pd
import xlsxwriter
from datetime import datetime
from dateutil.relativedelta import relativedelta
import re
from pdf2image import convert_from_path
import shutil

DEFAULT_VALUE = "---"
types = []
names = []
surnames = []
sexs = []
date_of_births = []
countrys = []
numbers = []
nationalitys = []
expiration_dates = []
passport_files = []

files = []
failed_file=0

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def standardized_name(name):
    name = re.sub(r"[ ]{1,}[K ]{2,}", "---", name)
    name = re.sub(r"[---]{2,}([A-Z])\w+|[---]{2,}", "", name)
    return name

def standardized_date(str_date, isDob):
    result=str_date
    try:
        dt = datetime.strptime(str_date, '%y%m%d')
        if isDob == True:
            if dt > datetime.now():
                dt -= relativedelta(years=100)
        result = "%d-%02d-%02d" % (dt.year, dt.month, dt.day)
    except:
        result=str_date
    return result

def update_data(mrz_data, file):
    if(mrz_data != None):
        types.append(mrz_data["type"])
        
        name_val = standardized_name(mrz_data["names"])
        names.append(name_val)

        surname_val = standardized_name(mrz_data["surname"])
        surnames.append(surname_val)
        
        sexs.append(mrz_data["sex"])
        date_of_births.append(standardized_date(mrz_data["date_of_birth"], True))
        countrys.append(mrz_data["country"])
        
        number_val = mrz_data["number"]
        number_val = re.sub('[!@#$<>]', '', number_val)
        numbers.append(number_val)
        
        nationalitys.append(mrz_data["nationality"])
        expiration_dates.append(standardized_date(mrz_data["expiration_date"], False))
        passport_files.append(file)
    else:
        types.append(DEFAULT_VALUE)
        names.append(DEFAULT_VALUE)
        surnames.append(DEFAULT_VALUE)
        sexs.append(DEFAULT_VALUE)
        date_of_births.append(DEFAULT_VALUE)
        countrys.append(DEFAULT_VALUE)
        numbers.append(DEFAULT_VALUE)
        nationalitys.append(DEFAULT_VALUE)
        expiration_dates.append(DEFAULT_VALUE)
        passport_files.append(file)

print("Start process passport")
ocr_path=resource_path("Tesseract-OCR")+"\\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = ocr_path
# r=root, d=directories, f = files
for r, d, f in os.walk("images"):
    for file in f:
        if file.endswith('.jpg') or file.endswith('.png') or file.endswith('.pdf'):
            files.append(os.path.join(r,file))

pdf_dir="pdf"
for i in range (len(files)):
    try:
        f = files[i]
        # Process image
        print("Process file ({0} / {1}): {2}".format(i, len(files), f))
        if(f.endswith('.pdf')):
            if not os.path.exists(pdf_dir):
                os.makedirs(pdf_dir)
            print("Process PDF file")
            pages = convert_from_path(f, dpi=200)
            page = pages[0]
            new_file=os.path.splitext(f.split('\\')[1])[0]+".jpg"
            pdf_convert_file=pdf_dir+"\\"+new_file
            print("converted file : %s" % pdf_convert_file)
            page.save(pdf_convert_file, 'JPEG')
            mrz = read_mrz(pdf_convert_file)
            mrz_data = mrz.to_dict()
            update_data(mrz_data, f)
        else:
            mrz = read_mrz(f)
            mrz_data = mrz.to_dict()
            update_data(mrz_data, f)
    except:
        failed_file+=1
        update_data(None, f)
        print("Can not process file : {0} - {1}\n".format(f, sys.exc_info()[0]))

if os.path.exists(pdf_dir):
    shutil.rmtree(pdf_dir)

# Create a Pandas dataframe from some data.

df = pd.DataFrame({
    'Name': names,
    'SurName' : surnames,
    'Sex' : sexs,
    'Date of birth' : date_of_births,
    'Country' : countrys,
    'Number' : numbers,
    'Nationality' : nationalitys,
    'Expiration Date' : expiration_dates,
    'File Name' : passport_files})

df = df[['Name','SurName','Sex','Date of birth', 'Country', 'Number', 'Nationality', 'Expiration Date', 'File Name']]
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('passport_detail.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')
# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Add a header format.
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})

# Write the column headers with the defined format.
for col_num, value in enumerate(df.columns.values):
    worksheet.write(0, col_num + 1, value, header_format)

# Set the column width and format.
worksheet.set_column('B:B', 20)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 10)
worksheet.set_column('E:E', 20)
worksheet.set_column('F:F', 10)
worksheet.set_column('G:G', 20)
worksheet.set_column('H:H', 10)
worksheet.set_column('I:I', 20)
worksheet.set_column('J:J', 100)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
print("=========================")
print("Total file processed :\t %d" % len(files))
print("Total failed file :\t %d" % failed_file)
print("Output file  :\t %s\n" % (os.getcwd() + r"\passport_detail.xlsx"))

print("Bye")
input("Press Enter to continue...")
