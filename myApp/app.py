import os
from werkzeug.utils import secure_filename # type: ignore
from flask import Flask, request, render_template # type: ignore
import pdfplumber # type: ignore
import pandas as pd # type: ignore
import pytesseract # type: ignore
from PIL import Image # type: ignore
import cv2 # type: ignore
import numpy as np # type: ignore

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

UPLOAD_FOLDER = './static/uploads'
ALLOWED_EXTENSIONS = set(['pdf', 'png', 'jpg', 'jpeg'])

app = Flask(__name__)
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = "secret key"

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_image(image_path):
    image = Image.open(image_path)
    gray = cv2.cvtColor(np.array(image), cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
    processed_image_path = os.path.join(app.config['UPLOAD_FOLDER'], 'processed_image.png')
    cv2.imwrite(processed_image_path, thresh)
    custom_config = r'--oem 3 --psm 6'
    text = pytesseract.image_to_string(thresh, config=custom_config)

    return text

def extract_tables_from_pdf(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                table = pd.DataFrame(table)
                table.columns = table.iloc[0]
                table.drop(0, inplace=True)
                tables.append(table)

    return tables

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/extract_table', methods=['POST'])
def extract_table():
    file = request.files['file']
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            text = extract_text_from_image(file_path)
            return render_template('home.html', org_img_name=filename, tables=[text], ntables=1)
        elif filename.lower().endswith('.pdf'):
            tables = extract_tables_from_pdf(file_path)
            with pd.ExcelWriter('extracted_tables.xlsx') as writer:
                for i, df in enumerate(tables):
                    df.to_excel(writer, sheet_name=str(i+1), index=False)
            tables_html = [x.to_html(index=False) for x in tables]
            return render_template('home.html', org_img_name=filename, tables=tables_html, ntables=len(tables))

    return render_template('home.html')

if __name__ == '__main__':
    app.run(debug=True)