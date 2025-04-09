from flask import Flask, request, send_file, jsonify
from werkzeug.utils import secure_filename
import os
import PyPDF2
from opencc import OpenCC
from docx import Document

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# 確保文件夾存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# 中文簡繁轉換
cc = OpenCC('s2t')  # 簡體轉繁體

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    
    # 提取 PDF 內容
    try:
        pdf_reader = PyPDF2.PdfReader(filepath)
        text = ''
        for page in pdf_reader.pages:
            text += page.extract_text()
        
        # 轉換為繁體中文
        traditional_text = cc.convert(text)
        
        # 生成 Word 檔案
        output_filename = os.path.splitext(filename)[0] + '.docx'
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        document = Document()
        document.add_paragraph(traditional_text)
        document.save(output_path)
        
        return jsonify({'download_link': f'/download/{output_filename}'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if os.path.exists(output_path):
        return send_file(output_path, as_attachment=True)
    else:
        return jsonify({'error': 'File not found'}), 404

if __name__ == '__main__':
    app.run(debug=True)