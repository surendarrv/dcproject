from flask import Flask, request, render_template, jsonify, send_file
import os
import tempfile
from converter import ExcelToTextConverter

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Check file type
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({'error': 'Please upload an Excel file (.xlsx or .xls)'}), 400
        
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
            file.save(temp_file.name)
            temp_path = temp_file.name
        
        try:
            # Convert Excel to text
            converter = ExcelToTextConverter()
            converted_text = converter.convert_excel_to_text(temp_path)
            
            # Clean up temp file
            os.unlink(temp_path)
            
            return jsonify({
                'success': True,
                'preview': converted_text,
                'filename': file.filename,
                'message': '🎉 MAPPING sheet found successfully!'
            })
            
        except Exception as e:
            # Clean up temp file
            if os.path.exists(temp_path):
                os.unlink(temp_path)
            
            error_msg = str(e)
            if "MAPPING sheet not found" in error_msg:
                return jsonify({
                    'error': '❌ MAPPING sheet not found in Excel file. Please ensure your Excel file contains a sheet named "MAPPING".'
                }), 400
            else:
                return jsonify({
                    'error': f'❌ Error processing file: {error_msg}'
                }), 400
    
    except Exception as e:
        return jsonify({'error': f'❌ Upload failed: {str(e)}'}), 500

@app.route('/download', methods=['POST'])
def download_file():
    try:
        data = request.get_json()
        content = data.get('content', '')
        filename = data.get('filename', 'output.txt')
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt') as temp_file:
            temp_file.write(content)
            temp_path = temp_file.name
        
        return send_file(
            temp_path,
            as_attachment=True,
            download_name=filename.replace('.xlsx', '.txt').replace('.xls', '.txt'),
            mimetype='text/plain'
        )
    
    except Exception as e:
        return jsonify({'error': f'Download failed: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)
