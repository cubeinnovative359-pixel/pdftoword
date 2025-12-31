"""
PDF to Word Converter - Vercel Serverless Function
Modified for serverless environment with temporary file handling
"""

from flask import Flask, request, jsonify, send_file
from pdf2docx import Converter
import tempfile
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configuration
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB (Vercel has 4.5MB limit per request, but keeping 10MB for flexibility)
ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/api/convert', methods=['POST', 'OPTIONS'])
def convert_pdf():
    """
    Convert PDF to Word document
    Returns the converted file directly as a download
    """
    # Handle CORS preflight
    if request.method == 'OPTIONS':
        response = jsonify({'status': 'ok'})
        response.headers.add('Access-Control-Allow-Origin', '*')
        response.headers.add('Access-Control-Allow-Headers', 'Content-Type')
        response.headers.add('Access-Control-Allow-Methods', 'POST')
        return response
    
    try:
        # Validate file presence
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': 'Only PDF files are allowed'}), 400
        
        # Check file size
        file.seek(0, os.SEEK_END)
        file_size = file.tell()
        file.seek(0)
        
        if file_size > MAX_FILE_SIZE:
            return jsonify({'success': False, 'error': 'File size exceeds 10MB limit'}), 400
        
        # Get original filename
        original_filename = secure_filename(file.filename)
        base_name = os.path.splitext(original_filename)[0]
        
        # Create temporary files (automatically cleaned up)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as pdf_temp:
            file.save(pdf_temp.name)
            pdf_path = pdf_temp.name
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as docx_temp:
            docx_path = docx_temp.name
        
        try:
            # Convert PDF to Word
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
            
            # Send file as response
            response = send_file(
                docx_path,
                as_attachment=True,
                download_name=f"{base_name}.docx",
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
            
            # Add CORS headers
            response.headers.add('Access-Control-Allow-Origin', '*')
            
            return response
            
        finally:
            # Cleanup temporary files
            try:
                if os.path.exists(pdf_path):
                    os.unlink(pdf_path)
                if os.path.exists(docx_path):
                    os.unlink(docx_path)
            except Exception as cleanup_error:
                print(f"Cleanup error: {str(cleanup_error)}")
        
    except Exception as e:
        print(f"Conversion error: {str(e)}")
        return jsonify({
            'success': False, 
            'error': f'Conversion failed: {str(e)}'
        }), 500

# Vercel serverless function handler
def handler(event, context):
    """Vercel serverless function entry point"""
    with app.request_context(event):
        return app.full_dispatch_request()

# For local testing
if __name__ == '__main__':
    app.run(debug=True, port=5000)
