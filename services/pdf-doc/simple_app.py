import os
import tempfile
import logging
from flask import Flask, request, render_template, send_from_directory, jsonify
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

# Local imports
from smart_converter import smart_pdf_to_docx

# .env 파일 로드
load_dotenv()

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

# 폴더 생성
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'pdf'}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '파일이 없습니다.'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'success': False, 'error': '파일이 선택되지 않았습니다.'}), 400
        
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(pdf_path)
        
        logging.info(f"변환 요청 받음: {filename}")
        
        # 지능형 변환기 사용
        output_path = smart_pdf_to_docx(pdf_path)
        
        if output_path and os.path.exists(output_path):
            logging.info(f"변환 성공: {output_path}")
            return jsonify({
                'success': True, 
                'download_url': f'/download/{os.path.basename(output_path)}'
            })
        else:
            logging.error("변환 실패 또는 결과 파일 없음")
            return jsonify({'success': False, 'error': 'PDF를 DOCX로 변환하는 데 실패했습니다.'}), 500
            
    return jsonify({'success': False, 'error': '허용되지 않는 파일 형식입니다.'}), 400

@app.route('/download/<filename>')
def download(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)