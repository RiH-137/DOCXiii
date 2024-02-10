from flask import Flask, render_template, request, send_file

from pdf2docx import Converter
from docx2pdf import convert

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    if request.method == 'POST':
        file = request.files['file']
        conversion_type = request.form['conversion_type']

        if conversion_type == 'pdf_to_docx':
            try:
                new_filename = f"{file.filename.rsplit('.', 1)[0]}.docx"
                converter = Converter(file)
                converter.convert(new_filename)
                converter.close()
                return send_file(new_filename, as_attachment=True)
            except Exception as e:
                error_message = f"Error converting PDF to DOCX: {e}"
                return render_template('index.html', error=error_message)

        elif conversion_type == 'docx_to_pdf':
            try:
                new_filename = f"{file.filename.rsplit('.', 1)[0]}.pdf"
                convert(file.filename, new_filename)
                return send_file(new_filename, as_attachment=True)
            except Exception as e:
                error_message = f"Error converting DOCX to PDF: {e}"
                return render_template('index.html', error=error_message)

        else:
            return render_template('index.html', error="Invalid conversion type")

if __name__ == '__main__':
    app.run(debug=True)
