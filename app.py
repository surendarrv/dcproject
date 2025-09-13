from flask import Flask, request, render_template, jsonify, send_file
import os
import tempfile
import re
from converter import ExcelToTextConverter

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Checpy if file was uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Get selected mode
        mode = request.form.get('mode', 'DEMO/IRAS/DIST-normal')
        
        # Check file type
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({'error': 'Please upload an Excel file (.xlsx or .xls)'}), 400
        
        # Check if mode is supported
        if mode not in ['DEMO/IRAS/DIST-normal', 'DIST_WITH_VARIABLE', 'IRAS_WITH_VARIABLE', 'XYZA', 'DIST']:
            return jsonify({'error': 'Invalid mode selected'}), 400
        
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
            file.save(temp_file.name)
            temp_path = temp_file.name
        
        try:
            # Handle different modes
            if mode == 'DEMO/IRAS/DIST-normal':
                # First, detect if there's a prefix in the column header
                converter = ExcelToTextConverter()
                prefix_info = converter.detect_prefix_in_file(temp_path)
                
                if prefix_info['has_prefix']:
                    # Return prefix detection info for user confirmation
                    # Store temp path in a way that can be retrieved later
                    return jsonify({
                        'success': False,
                        'prefix_detected': True,
                        'prefix': prefix_info['prefix'],
                        'column_name': prefix_info['column_name'],
                        'temp_path': temp_path,  # Pass the temp path for later use
                        'filename': file.filename,
                        'message': f"🔍 Detected prefix '{prefix_info['prefix']}' in column header: '{prefix_info['column_name']}'"
                    })
                else:
                    # No prefix detected, proceed with normal conversion
                    converted_text = converter.convert_excel_to_text(temp_path)
                    
                    # Clean up temp file
                    os.unlink(temp_path)
                    
                    return jsonify({
                        'success': True,
                        'preview': converted_text,
                        'filename': file.filename,
                        'message': '🎉 MAPPING sheet found successfully!'
                    })
            elif mode == 'DIST':
                # DIST mode - show coming soon message
                # Clean up temp file
                os.unlink(temp_path)
                
                return jsonify({
                    'success': False,
                    'error': '🚧 DIST mode is coming soon! Please use "DEMO", "IRAS", or "XYZA" mode for now.'
                })
            elif mode == 'XYZA':
                # XYZA mode - show coming soon message
                # Clean up temp file
                os.unlink(temp_path)
                
                return jsonify({
                    'success': False,
                    'error': '🚧 XYZA mode is coming soon! Please use "DEMO", "DIST", or "IRAS" mode for now.'
                })
            elif mode == 'IRAS_WITH_VARIABLE':
                # IRAS mode - same logic as DEMO but without detailed table data extraction
                converter = ExcelToTextConverter()
                prefix_info = converter.detect_prefix_in_file(temp_path)
                
                if prefix_info['has_prefix']:
                    # Return prefix detection info for user confirmation
                    return jsonify({
                        'success': False,
                        'prefix_detected': True,
                        'prefix': prefix_info['prefix'],
                        'column_name': prefix_info['column_name'],
                        'temp_path': temp_path,
                        'filename': file.filename,
                        'message': f"🔍 Detected prefix '{prefix_info['prefix']}' in column header: '{prefix_info['column_name']}'"
                    })
                else:
                    # No prefix detected, proceed with IRAS conversion (mapping fields only)
                    converted_text = converter.convert_excel_to_text(temp_path, mode='IRAS')
                    
                    # Clean up temp file
                    os.unlink(temp_path)
                    
                    return jsonify({
                        'success': True,
                        'preview': converted_text,
                        'filename': file.filename,
                        'message': '🎉 IRAS mode - MAPPING sheet found successfully!'
                    })
            else:
                # For DIST_WITH_VARIABLE mode
                # Clean up temp file
                os.unlink(temp_path)
                
                return jsonify({
                    'success': False,
                    'error': f'🚧 {mode.replace("_", " ")} mode is coming soon! Please use "DEMO", "IRAS", or "XYZA" mode for now.'
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

@app.route('/convert_with_prefix', methods=['POST'])
def convert_with_prefix():
    try:
        data = request.get_json()
        temp_path = data.get('temp_path')
        apply_prefix = data.get('apply_prefix', True)
        
        if not temp_path or not os.path.exists(temp_path):
            return jsonify({'error': 'File not found'}), 400
        
        try:
            converter = ExcelToTextConverter()
            # Check if this is DOCA mode by looking at the request data
            mode = data.get('mode', 'DEMO/IRAS/DIST-normal')
            converted_text = converter.convert_excel_to_text(temp_path, apply_prefix, mode=mode)
            
            # Clean up temp file
            os.unlink(temp_path)
            
            return jsonify({
                'success': True,
                'preview': converted_text,
                'message': f'🎉 Conversion completed {"with" if apply_prefix else "without"} prefix!'
            })
            
        except Exception as e:
            # Clean up temp file
            if os.path.exists(temp_path):
                os.unlink(temp_path)
            
            return jsonify({
                'error': f'❌ Error processing file: {str(e)}'
            }), 400
    
    except Exception as e:
        return jsonify({'error': f'❌ Conversion failed: {str(e)}'}), 500

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

@app.route('/cobol-generator')
def cobol_generator():
    return render_template('cobol_generator.html')

@app.route('/generate-cobol', methods=['POST'])
def generate_cobol():
    try:
        data = request.get_json()
        content = data.get('content', '')
        
        if not content:
            return jsonify({'error': 'No content provided'}), 400
        
        # Generate COBOL code from the control card content
        cobol_code = generate_cobol_from_control_card(content)
        
        return jsonify({
            'success': True,
            'cobol_code': cobol_code
        })
    
    except Exception as e:
        return jsonify({'error': f'COBOL generation failed: {str(e)}'}), 500

def generate_cobol_from_control_card(content):
    """Generate COBOL code from control card content"""
    
    lines = content.split('\n')
    cobol_lines = []
    
    # COBOL program header
    cobol_lines.extend([
        "       IDENTIFICATION DIVISION.",
        "       PROGRAM-ID. CONTROL-CARD-PROCESSOR.",
        "       AUTHOR. AI-GENERATED.",
        "       DATE-WRITTEN. " + str(__import__('datetime').datetime.now().strftime('%Y-%m-%d')) + ".",
        "",
        "       ENVIRONMENT DIVISION.",
        "       CONFIGURATION SECTION.",
        "       SOURCE-COMPUTER. IBM-370.",
        "       OBJECT-COMPUTER. IBM-370.",
        "",
        "       DATA DIVISION.",
        "       WORKING-STORAGE SECTION.",
        "       01  WS-CONTROL-CARD.",
        "           05  WS-FIELD-NAME           PIC X(30).",
        "           05  WS-FIELD-POSITION       PIC 9(4).",
        "           05  WS-FIELD-LENGTH         PIC 9(3).",
        "           05  WS-FIELD-TYPE           PIC X(10).",
        "           05  WS-TABLE-REF            PIC X(1).",
        "",
        "       01  WS-OUTPUT-RECORD.",
        "           05  WS-RECORD-TYPE          PIC X(4).",
        "           05  WS-FIELD-DATA           PIC X(50).",
        "",
        "       01  WS-EOF-FLAG                PIC X(1) VALUE 'N'.",
        "           88  WS-EOF                  VALUE 'Y'.",
        "",
        "       PROCEDURE DIVISION.",
        "       0000-MAIN-PROCESSING.",
        "           PERFORM 1000-INITIALIZE",
        "           PERFORM 2000-PROCESS-CONTROL-CARDS",
        "           PERFORM 3000-CLOSE-FILES",
        "           STOP RUN.",
        "",
        "       1000-INITIALIZE.",
        "           DISPLAY 'CONTROL CARD PROCESSOR STARTED'",
        "           MOVE 'N' TO WS-EOF-FLAG.",
        "",
        "       2000-PROCESS-CONTROL-CARDS.",
        "           PERFORM UNTIL WS-EOF",
        "               READ CONTROL-CARD-FILE",
        "                   AT END MOVE 'Y' TO WS-EOF-FLAG",
        "                   NOT AT END PERFORM 2100-PROCESS-RECORD",
        "               END-READ",
        "           END-PERFORM.",
        "",
        "       2100-PROCESS-RECORD.",
        "           EVALUATE WS-RECORD-TYPE",
        "               WHEN '0000' PERFORM 2200-PROCESS-FIELD-MAPPING",
        "               WHEN 'TABL' PERFORM 2300-PROCESS-TABLE-REFERENCE",
        "               WHEN OTHER  DISPLAY 'UNKNOWN RECORD TYPE: ' WS-RECORD-TYPE",
        "           END-EVALUATE.",
        "",
        "       2200-PROCESS-FIELD-MAPPING.",
        "           DISPLAY 'PROCESSING FIELD MAPPING: ' WS-FIELD-DATA",
        "           PERFORM 2400-VALIDATE-FIELD",
        "           PERFORM 2500-WRITE-OUTPUT-RECORD.",
        "",
        "       2300-PROCESS-TABLE-REFERENCE.",
        "           DISPLAY 'PROCESSING TABLE REFERENCE: ' WS-FIELD-DATA",
        "           PERFORM 2600-LOOKUP-TABLE-DATA.",
        "",
        "       2400-VALIDATE-FIELD.",
        "           IF WS-FIELD-LENGTH > 0",
        "               DISPLAY 'FIELD LENGTH VALID: ' WS-FIELD-LENGTH",
        "           ELSE",
        "               DISPLAY 'ERROR: INVALID FIELD LENGTH'",
        "           END-IF.",
        "",
        "       2500-WRITE-OUTPUT-RECORD.",
        "           WRITE OUTPUT-RECORD FROM WS-OUTPUT-RECORD.",
        "",
        "       2600-LOOKUP-TABLE-DATA.",
        "           EVALUATE WS-TABLE-REF",
        "               WHEN 'A' PERFORM 2700-PROCESS-TABLE-A",
        "               WHEN 'B' PERFORM 2800-PROCESS-TABLE-B",
        "               WHEN 'C' PERFORM 2900-PROCESS-TABLE-C",
        "               WHEN 'D' PERFORM 3000-PROCESS-TABLE-D",
        "               WHEN OTHER DISPLAY 'UNKNOWN TABLE REFERENCE: ' WS-TABLE-REF",
        "           END-EVALUATE.",
        "",
        "       2700-PROCESS-TABLE-A.",
        "           DISPLAY 'PROCESSING TABLE A - DOLLAR RANGES'",
        "           PERFORM 3100-LOAD-TABLE-A-DATA.",
        "",
        "       2800-PROCESS-TABLE-B.",
        "           DISPLAY 'PROCESSING TABLE B - OCCUPATION CODES'",
        "           PERFORM 3200-LOAD-TABLE-B-DATA.",
        "",
        "       2900-PROCESS-TABLE-C.",
        "           DISPLAY 'PROCESSING TABLE C - ASSOCIATED FIELDS'",
        "           PERFORM 3300-LOAD-TABLE-C-DATA.",
        "",
        "       3000-PROCESS-TABLE-D.",
        "           DISPLAY 'PROCESSING TABLE D - TIME NEEDS'",
        "           PERFORM 3400-LOAD-TABLE-D-DATA.",
        "",
        "       3100-LOAD-TABLE-A-DATA.",
        "           DISPLAY 'LOADING TABLE A DATA FOR DOLLAR RANGE PROCESSING'",
        "           MOVE 'TABLE-A-LOADED' TO WS-FIELD-DATA.",
        "",
        "       3200-LOAD-TABLE-B-DATA.",
        "           DISPLAY 'LOADING TABLE B DATA FOR OCCUPATION PROCESSING'",
        "           MOVE 'TABLE-B-LOADED' TO WS-FIELD-DATA.",
        "",
        "       3300-LOAD-TABLE-C-DATA.",
        "           DISPLAY 'LOADING TABLE C DATA FOR ASSOCIATED FIELDS'",
        "           MOVE 'TABLE-C-LOADED' TO WS-FIELD-DATA.",
        "",
        "       3400-LOAD-TABLE-D-DATA.",
        "           DISPLAY 'LOADING TABLE D DATA FOR TIME NEEDS'",
        "           MOVE 'TABLE-D-LOADED' TO WS-FIELD-DATA.",
        "",
        "       3000-CLOSE-FILES.",
        "           DISPLAY 'CONTROL CARD PROCESSOR COMPLETED'",
        "           CLOSE CONTROL-CARD-FILE",
        "           CLOSE OUTPUT-FILE.",
        "",
        "       END PROGRAM CONTROL-CARD-PROCESSOR."
    ])
    
    # Add specific field mappings from the control card
    field_count = 0
    table_count = 0
    
    for line in lines:
        line = line.strip()
        if line.startswith('0000 '):
            field_count += 1
            # Extract field information
            parts = line.split()
            if len(parts) >= 3:
                field_name = parts[1] if len(parts) > 1 else 'UNKNOWN'
                field_pos = parts[2] if len(parts) > 2 else '0000'
                
                cobol_lines.extend([
                    "",
                    f"       * FIELD MAPPING #{field_count}",
                    f"       * Field: {field_name}",
                    f"       * Position: {field_pos}",
                    f"       01  WS-FIELD-{field_count:03d}.",
                    f"           05  WS-{field_name.replace('-', '-')}",
                    f"                PIC X(20) VALUE '{field_name}'.",
                ])
        
        elif line.startswith('# table '):
            table_count += 1
            table_ref = line.split()[-1] if line.split() else 'X'
            cobol_lines.extend([
                "",
                f"       * TABLE REFERENCE #{table_count}",
                f"       * Table: {table_ref}",
                f"       01  WS-TABLE-{table_ref}.",
                f"           05  WS-TABLE-{table_ref}-DATA",
                f"                PIC X(50) VALUE 'TABLE-{table_ref}-DATA'.",
            ])
    
    # Add summary information
    cobol_lines.extend([
        "",
        "       * SUMMARY INFORMATION",
        f"       * Total Fields Processed: {field_count}",
        f"       * Total Table References: {table_count}",
        "       * Generated by AI Control Card Processor",
        ""
    ])
    
    return '\n'.join(cobol_lines)

if __name__ == '__main__':
    print("Starting Flask application...")
    print("Application will be available at: http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
