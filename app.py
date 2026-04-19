import os
import uuid
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename
from pdf2docx import Converter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB limit

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), 'outputs')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Only PDF files are allowed'}), 400

    # Enforce size limit explicitly (werkzeug raises RequestEntityTooLarge on
    # oversized requests, but checking content_length gives a nicer message).
    content_length = request.content_length
    if content_length and content_length > app.config['MAX_CONTENT_LENGTH']:
        return jsonify({'error': 'File exceeds the 50 MB size limit'}), 413

    original_name = secure_filename(file.filename)
    base_name = os.path.splitext(original_name)[0]
    unique_id = uuid.uuid4().hex

    pdf_path = os.path.join(UPLOAD_FOLDER, f'{unique_id}.pdf')
    docx_path = os.path.join(OUTPUT_FOLDER, f'{unique_id}.docx')

    try:
        file.save(pdf_path)

        cv = Converter(pdf_path)
        cv.convert(docx_path)
        cv.close()

        download_name = f'{base_name}.docx'
        return send_file(
            docx_path,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )
    except Exception:
        return jsonify({'error': 'Conversion failed. Please ensure the file is a valid PDF.'}), 500
    finally:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        if os.path.exists(docx_path):
            os.remove(docx_path)


if __name__ == '__main__':
    app.run(debug=False)
