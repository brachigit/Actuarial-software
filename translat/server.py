from flask import Flask, request, jsonify, send_file
import os
from extract_pdf import extract_pdf
from datetime import datetime

app = Flask(__name__)

# הוספת CORS headers ידנית
@app.after_request
def after_request(response):
    response.headers.add('Access-Control-Allow-Origin', '*')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization')
    response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
    return response

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/upload', methods=['POST'])
def upload_pdf():
    try:
        print("Received upload request")
        
        if 'pdf' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['pdf']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        print(f"Uploading file: {file.filename}")
        
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        
        print(f"File saved to: {filepath}")
        
        return jsonify({'message': 'PDF uploaded', 'filename': file.filename})
    
    except Exception as e:
        print(f"Error in upload: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/translate', methods=['POST'])
def translate_pdf():
    try:
        filename = request.json['filename']
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        
        if not os.path.exists(filepath):
            return jsonify({'error': 'File not found'}), 404
        
        print(f"Starting translation for: {filename}")
        
        # קריאה לפונקציית extract_pdf
        excel_path = extract_pdf(filepath)
        
        print(f"Excel file created: {excel_path}")
        
        return jsonify({
            'message': f'Translation complete for {filename}',
            'excel_file': os.path.basename(excel_path),
            'download_url': f'/download/{os.path.basename(excel_path)}'
        })
    
    except Exception as e:
        print(f"Error in translate: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/files', methods=['GET'])
def list_files():
    try:
        outputs_dir = 'outputs'
        if not os.path.exists(outputs_dir):
            return jsonify({'files': []})
        
        files = []
        for filename in os.listdir(outputs_dir):
            if filename.endswith('.xlsx'):
                file_path = os.path.join(outputs_dir, filename)
                file_stat = os.stat(file_path)
                files.append({
                    'name': filename,
                    'size': file_stat.st_size,
                    'created': datetime.fromtimestamp(file_stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                    'download_url': f'/download/{filename}'
                })
        
        # מיון לפי תאריך יצירה (החדשים ראשונים)
        files.sort(key=lambda x: x['created'], reverse=True)
        
        return jsonify({'files': files})
    
    except Exception as e:
        print(f"Error listing files: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        outputs_dir = 'outputs'
        file_path = os.path.join(outputs_dir, filename)
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(file_path, as_attachment=True)
    
    except Exception as e:
        print(f"Error in download: {str(e)}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    print("Starting server...")
    print(f"Upload folder: {os.path.abspath(UPLOAD_FOLDER)}")
    app.run(debug=True, host='0.0.0.0', port=5000)
