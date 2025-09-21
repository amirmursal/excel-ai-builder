#!/usr/bin/env python3
"""
Web-based AI Excel Automation Interface
"""

from flask import Flask, render_template_string, request, jsonify, send_file
import pandas as pd
import os
from datetime import datetime
import json

# Import our automation class
from ai_excel_automation import AIExcelAutomation

app = Flask(__name__)

# Global automation instance
automation = None

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>AI Excel Automation</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .header { text-align: center; color: #333; margin-bottom: 30px; }
        .section { margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }
        .form-group { margin: 10px 0; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input[type="text"], textarea { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
        button { background: #007bff; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin: 5px; }
        button:hover { background: #0056b3; }
        .output { background: #f8f9fa; padding: 15px; border-radius: 4px; white-space: pre-wrap; font-family: monospace; max-height: 400px; overflow-y: auto; }
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .info { background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb; }
        .sheet-info { background: #e2e3e5; padding: 10px; border-radius: 4px; margin: 10px 0; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🤖 AI Excel Automation</h1>
            <p>Upload an Excel file and give natural language instructions to automate data processing</p>
        </div>

        <div class="section">
            <h3>📊 Current File Status</h3>
            <div id="file-status">
                {% if automation %}
                    <div class="success">
                        ✅ File loaded: {{ automation.excel_file_path }}
                        <br>📋 Sheets: {{ automation.data.keys() | list | length }}
                        <br>📊 Current sheet: {{ automation.current_sheet }}
                    </div>
                {% else %}
                    <div class="info">
                        ℹ️ No file loaded. Please upload an Excel file.
                    </div>
                {% endif %}
            </div>
        </div>

        <div class="section">
            <h3>📁 Upload Excel File</h3>
            <form action="/upload" method="post" enctype="multipart/form-data">
                <div class="form-group">
                    <label for="file">Select Excel File:</label>
                    <input type="file" id="file" name="file" accept=".xlsx,.xls" required>
                </div>
                <button type="submit">Upload File</button>
            </form>
        </div>

        {% if automation %}
        <div class="section">
            <h3>📋 Available Sheets</h3>
            <div class="sheet-info">
                {% for sheet_name, df in automation.data.items() %}
                    <div style="margin: 5px 0;">
                        <strong>{{ sheet_name }}</strong> 
                        {% if sheet_name == automation.current_sheet %}(current){% endif %}
                        - {{ df.shape[0] }} rows × {{ df.shape[1] }} columns
                        <button onclick="switchSheet('{{ sheet_name }}')" style="margin-left: 10px; padding: 5px 10px; font-size: 12px;">Switch</button>
                    </div>
                {% endfor %}
            </div>
        </div>

        <div class="section">
            <h3>🤖 Give Instructions</h3>
            <form action="/execute" method="post">
                <div class="form-group">
                    <label for="instruction">Enter your instruction:</label>
                    <input type="text" id="instruction" name="instruction" 
                           placeholder="e.g., 'show first 10 rows', 'count insurance types', 'filter by date range'" 
                           style="width: 100%;">
                </div>
                <button type="submit">Execute Instruction</button>
            </form>
            
            <div style="margin-top: 15px;">
                <h4>💡 Example Instructions:</h4>
                <ul>
                    <li><code>show first 10 rows</code> - Display first 10 rows</li>
                    <li><code>count insurance types</code> - Count different insurance types</li>
                    <li><code>show data info</code> - Display data information</li>
                    <li><code>filter by office name</code> - Filter data by office</li>
                    <li><code>generate summary report</code> - Create a summary report</li>
                </ul>
            </div>
        </div>

        <div class="section">
            <h3>📤 Export Data</h3>
            <form action="/export" method="post">
                <div class="form-group">
                    <label for="filename">Export filename (optional):</label>
                    <input type="text" id="filename" name="filename" 
                           placeholder="processed_data.xlsx">
                </div>
                <button type="submit">Export Current Data</button>
            </form>
        </div>
        {% endif %}

        <div class="section">
            <h3>📝 Output</h3>
            <div class="output" id="output">
                {% if output %}
                    {{ output }}
                {% else %}
                    Ready to process your instructions...
                {% endif %}
            </div>
        </div>
    </div>

    <script>
        function switchSheet(sheetName) {
            fetch('/switch_sheet', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({sheet: sheetName})
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    location.reload();
                } else {
                    alert('Error: ' + data.error);
                }
            });
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    global automation
    return render_template_string(HTML_TEMPLATE, automation=automation, output="")

@app.route('/upload', methods=['POST'])
def upload_file():
    global automation
    
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    try:
        # Save uploaded file
        filename = file.filename
        file.save(filename)
        
        # Initialize automation
        automation = AIExcelAutomation(filename)
        
        return jsonify({
            'success': True, 
            'message': f'File uploaded successfully! Loaded {len(automation.data)} sheets.',
            'sheets': list(automation.data.keys()),
            'current_sheet': automation.current_sheet
        })
        
    except Exception as e:
        return jsonify({'error': f'Error uploading file: {str(e)}'}), 500

@app.route('/execute', methods=['POST'])
def execute_instruction():
    global automation
    
    if not automation:
        return jsonify({'error': 'No file loaded'}), 400
    
    instruction = request.form.get('instruction', '').strip()
    if not instruction:
        return jsonify({'error': 'No instruction provided'}), 400
    
    try:
        # Capture output
        import io
        import sys
        from contextlib import redirect_stdout
        
        output_buffer = io.StringIO()
        
        with redirect_stdout(output_buffer):
            automation.execute_instruction(instruction)
        
        output = output_buffer.getvalue()
        
        return render_template_string(HTML_TEMPLATE, 
                                   automation=automation, 
                                   output=output)
        
    except Exception as e:
        return render_template_string(HTML_TEMPLATE, 
                                   automation=automation, 
                                   output=f"Error: {str(e)}")

@app.route('/switch_sheet', methods=['POST'])
def switch_sheet():
    global automation
    
    if not automation:
        return jsonify({'error': 'No file loaded'}), 400
    
    data = request.get_json()
    sheet_name = data.get('sheet')
    
    if not sheet_name:
        return jsonify({'error': 'No sheet name provided'}), 400
    
    try:
        automation.switch_sheet(sheet_name)
        return jsonify({'success': True, 'current_sheet': automation.current_sheet})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/export', methods=['POST'])
def export_data():
    global automation
    
    if not automation:
        return jsonify({'error': 'No file loaded'}), 400
    
    filename = request.form.get('filename', '').strip()
    if not filename:
        filename = None
    
    try:
        output_file = automation.export_data(filename)
        if output_file:
            return send_file(output_file, as_attachment=True)
        else:
            return jsonify({'error': 'Failed to export data'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    print("🚀 Starting AI Excel Automation Web Interface...")
    print("📱 Open your browser and go to: http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
