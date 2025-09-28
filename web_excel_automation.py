#!/usr/bin/env python3
"""
Web-based Excel Automation App
Upload Excel files and give natural language instructions
"""

from flask import Flask, render_template_string, request, jsonify, send_file, redirect, url_for
import pandas as pd
import os
import json
from datetime import datetime
import re
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Global variables to store session data
current_data = {}
current_sheet = None
current_filename = None
conversation_history = []

# HTML Template
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AI Excel Automation</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        .container { 
            max-width: 1200px; 
            margin: 0 auto; 
            background: white; 
            border-radius: 15px; 
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            overflow: hidden;
        }
        .header { 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white; 
            padding: 30px; 
            text-align: center; 
        }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p { font-size: 1.2em; opacity: 0.9; }
        .nav-buttons {
            display: flex;
            gap: 15px;
            margin-bottom: 20px;
            padding: 15px;
            background: #f8f9fa;
            border-radius: 8px;
            border: 1px solid #e9ecef;
        }
        .nav-buttons a {
            display: inline-block;
            padding: 10px 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            text-decoration: none;
            border-radius: 6px;
            font-weight: 600;
            transition: all 0.3s ease;
            box-shadow: 0 2px 8px rgba(102, 126, 234, 0.3);
        }
        .nav-buttons a:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
        }
        .nav-buttons a:active {
            transform: translateY(0);
        }
        .content { padding: 30px; }
        .section { 
            margin: 25px 0; 
            padding: 25px; 
            border: 1px solid #e0e0e0; 
            border-radius: 10px; 
            background: #fafafa;
        }
        .section h3 { 
            color: #333; 
            margin-bottom: 20px; 
            font-size: 1.4em;
            border-bottom: 2px solid #667eea;
            padding-bottom: 10px;
        }
        .form-group { margin: 15px 0; }
        label { 
            display: block; 
            margin-bottom: 8px; 
            font-weight: 600; 
            color: #555;
        }
        input[type="file"], input[type="text"], textarea { 
            width: 100%; 
            padding: 12px; 
            border: 2px solid #ddd; 
            border-radius: 8px; 
            font-size: 16px;
            transition: border-color 0.3s;
        }
        input[type="file"]:focus, input[type="text"]:focus, textarea:focus { 
            outline: none; 
            border-color: #667eea; 
        }
        button { 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white; 
            padding: 12px 25px; 
            border: none; 
            border-radius: 8px; 
            cursor: pointer; 
            margin: 5px; 
            font-size: 16px;
            font-weight: 600;
            transition: transform 0.2s;
        }
        button:hover { 
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        }
        .output { 
            background: #1e1e1e; 
            color: #f8f8f2; 
            padding: 20px; 
            border-radius: 8px; 
            white-space: pre-wrap; 
            font-family: 'Courier New', monospace; 
            max-height: 400px; 
            overflow-y: auto;
            border: 1px solid #333;
        }
        .status { 
            padding: 15px; 
            margin: 15px 0; 
            border-radius: 8px; 
            font-weight: 600;
        }
        .success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .info { background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb; }
        .sheet-info { 
            background: #f8f9fa; 
            padding: 15px; 
            border-radius: 8px; 
            margin: 15px 0; 
            border-left: 4px solid #667eea;
        }
        .sheet-item { 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            padding: 10px; 
            margin: 5px 0; 
            background: white; 
            border-radius: 5px; 
            border: 1px solid #e0e0e0;
        }
        .sheet-item.current { 
            background: #e3f2fd; 
            border-color: #2196f3; 
        }
        .examples { 
            background: #f0f8ff; 
            padding: 15px; 
            border-radius: 8px; 
            margin-top: 15px;
        }
        .examples h4 { color: #1976d2; margin-bottom: 10px; }
        .examples ul { margin-left: 20px; }
        .examples li { margin: 5px 0; }
        .examples code { 
            background: #e3f2fd; 
            padding: 2px 6px; 
            border-radius: 3px; 
            font-family: 'Courier New', monospace;
        }
        .loading { 
            display: none; 
            text-align: center; 
            padding: 20px; 
        }
        .spinner { 
            border: 4px solid #f3f3f3; 
            border-top: 4px solid #667eea; 
            border-radius: 50%; 
            width: 40px; 
            height: 40px; 
            animation: spin 1s linear infinite; 
            margin: 0 auto;
        }
        @keyframes spin { 
            0% { transform: rotate(0deg); } 
            100% { transform: rotate(360deg); } 
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🤖 AI Excel Automation</h1>
            <p>Upload Excel files and give natural language instructions to automate your data processing</p>
        </div>

        <div class="content">
            <div class="nav-buttons">
                <a href="#" onclick="openComparisonTool()">📊 Comparison Tool</a>
            </div>

            <!-- File Status -->
            <div class="section">
                <h3>📊 Current File Status</h3>
                <div id="file-status">
                    {% if current_filename %}
                        <div class="success">
                            ✅ File loaded: {{ current_filename }}
                            <br>📋 Sheets: {{ current_data.keys() | list | length }}
                            <br>📊 Current sheet: {{ current_sheet }}
                        </div>
                    {% else %}
                        <div class="info">
                            ℹ️ No file loaded. Please upload an Excel file to get started.
                        </div>
                    {% endif %}
                </div>
            </div>

            <!-- File Upload -->
            <div class="section">
                <h3>📁 Upload Excel File</h3>
                <form action="/upload" method="post" enctype="multipart/form-data" id="upload-form">
                    <div class="form-group">
                        <label for="file">Select Excel File (.xlsx, .xls):</label>
                        <input type="file" id="file" name="file" accept=".xlsx,.xls" required>
                    </div>
                    <button type="submit" id="upload-btn">📤 Upload File</button>
                    <button type="button" onclick="resetApp()" style="background: #dc3545;">🔄 Reset App</button>
                </form>
                <div class="loading" id="upload-loading">
                    <div class="spinner"></div>
                    <p>Processing file...</p>
                </div>
            </div>

            {% if current_filename %}
            <!-- Sheet Selection -->
            <div class="section">
                <h3>📋 Available Sheets</h3>
                <div class="sheet-info">
                    {% for sheet_name, df in current_data.items() %}
                        <div class="sheet-item {% if sheet_name == current_sheet %}current{% endif %}">
                            <div>
                                <strong>{{ sheet_name }}</strong>
                                {% if sheet_name == current_sheet %}(current){% endif %}
                                - {{ df.shape[0] }} rows × {{ df.shape[1] }} columns
                            </div>
                            <button onclick="switchSheet('{{ sheet_name }}')" class="switch-btn">Switch</button>
                        </div>
                    {% endfor %}
                </div>
            </div>

            <!-- Instruction Input -->
            <div class="section">
                <h3>🤖 Give Instructions</h3>
                <form action="/execute" method="post" id="instruction-form">
                    <div class="form-group">
                        <label for="instruction">Enter your instruction:</label>
                        <input type="text" id="instruction" name="instruction" 
                               placeholder="e.g., 'show first 10 rows', 'copy Insurance column to Insurance New', 'count appointments by office'"
                               style="width: 100%;">
                    </div>
                    <button type="submit" id="execute-btn">🚀 Execute Instruction</button>
                </form>
                
                <div class="examples">
                    <h4>💡 Example Instructions:</h4>
                    <ul>
                        <li><code>show first 10 rows</code> - Display first 10 rows</li>
                        <li><code>show data info</code> - Display data information</li>
                        <li><code>count insurance types</code> - Count different insurance types</li>
                        <li><code>copy Insurance column to Insurance New</code> - Copy column data</li>
                        <li><code>filter by office name</code> - Filter data by office</li>
                        <li><code>generate summary report</code> - Create a summary report</li>
                        <li><code>export current data</code> - Export processed data</li>
                    </ul>
                </div>
                
                <div class="loading" id="execute-loading">
                    <div class="spinner"></div>
                    <p>Processing instruction...</p>
                </div>
            </div>

            <!-- Export Data -->
            <div class="section">
                <h3>📤 Export Data</h3>
                <form action="/export" method="post">
                    <div class="form-group">
                        <label for="filename">Export filename (optional):</label>
                        <input type="text" id="filename" name="filename" 
                               placeholder="processed_data.xlsx">
                    </div>
                    <button type="submit">💾 Export Current Data</button>
                </form>
            </div>
            {% endif %}

            <!-- Output Display -->
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
    </div>

    <script>
        // Form submission with loading states
        document.getElementById('upload-form').addEventListener('submit', function() {
            document.getElementById('upload-loading').style.display = 'block';
            document.getElementById('upload-btn').disabled = true;
        });

        document.getElementById('instruction-form').addEventListener('submit', function(e) {
            // Check if file is loaded
            if (!document.querySelector('.sheet-info')) {
                e.preventDefault();
                alert('Please upload an Excel file first before giving instructions.');
                return;
            }
            document.getElementById('execute-loading').style.display = 'block';
            document.getElementById('execute-btn').disabled = true;
        });

        // Switch sheet function
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
            })
            .catch(error => {
                alert('Error switching sheet: ' + error);
            });
        }

        // Auto-scroll output to bottom
        function scrollOutput() {
            const output = document.getElementById('output');
            output.scrollTop = output.scrollHeight;
        }

        // Scroll output on page load
        window.onload = function() {
            scrollOutput();
        }

        // Reset app function
        function openComparisonTool() {
            // Check if we're on Railway (production) or localhost
            const currentHost = window.location.hostname;
            let comparisonUrl;
            
            if (currentHost.includes('railway.app') || currentHost.includes('up.railway.app')) {
                // Production: Use the comparison tool URL
                comparisonUrl = 'https://your-comparison-app.railway.app/comparison';
            } else {
                // Local development
                comparisonUrl = 'http://localhost:5002/comparison';
            }
            
            window.open(comparisonUrl, '_blank');
        }

        function resetApp() {
            if (confirm('Are you sure you want to reset the app? This will clear all uploaded data.')) {
                fetch('/reset', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'}
                })
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        // Clear the output area
                        document.getElementById('output').innerHTML = '';
                        // Reload the page to reset the UI
                        location.reload();
                    } else {
                        alert('Error resetting app: ' + data.error);
                    }
                })
                .catch(error => {
                    alert('Error resetting app: ' + error);
                });
            }
        }
    </script>
</body>
</html>
"""

def process_instruction(instruction, df):
    """Process instruction and return code to execute"""
    instruction = instruction.lower().strip()
    
    if "show" in instruction or "display" in instruction:
        if "first" in instruction:
            num = extract_number(instruction) or 10
            return f"print(df.head({num}))"
        elif "last" in instruction:
            num = extract_number(instruction) or 10
            return f"print(df.tail({num}))"
        else:
            return "print(df.head(10))"
    
    elif "info" in instruction or "describe" in instruction:
        return """
print("=== DATA INFO ===")
print(f"Shape: {df.shape}")
print(f"Columns: {list(df.columns)}")
print("\\nData types:")
print(df.dtypes)
print("\\nMissing values:")
print(df.isnull().sum())
print("\\nBasic statistics:")
print(df.describe())
"""
    
    elif "copy" in instruction and "column" in instruction:
        if "insurance" in instruction and "insurance new" in instruction:
            return """
df['Insurance New'] = df['Insurance']
print(f"✅ Copied Insurance column to Insurance New")
print(f"Insurance New now has {df['Insurance New'].notna().sum()} non-null values")
"""
        else:
            return "print('Available columns for copying:', list(df.columns))"
    
    elif "reformat" in instruction and "insurance" in instruction:
        return """
import re

# State abbreviations mapping
STATE_ABBREVIATIONS = {
    'AL': 'Alabama', 'AK': 'Alaska', 'AR': 'Arkansas', 'AZ': 'Arizona',
    'CA': 'California', 'CO': 'Colorado', 'CT': 'Connecticut', 'DE': 'Delaware',
    'DC': 'District of Columbia', 'FL': 'Florida', 'GA': 'Georgia', 'HI': 'Hawaii',
    'ID': 'Idaho', 'IL': 'Illinois', 'IN': 'Indiana', 'IA': 'Iowa',
    'KS': 'Kansas', 'KY': 'Kentucky', 'LA': 'Louisiana', 'ME': 'Maine',
    'MD': 'Maryland', 'MA': 'Massachusetts', 'MI': 'Michigan', 'MN': 'Minnesota',
    'MS': 'Mississippi', 'MO': 'Missouri', 'MT': 'Montana', 'NE': 'Nebraska',
    'NV': 'Nevada', 'NH': 'New Hampshire', 'NJ': 'New Jersey', 'NM': 'New Mexico',
    'NY': 'New York', 'NC': 'North Carolina', 'ND': 'North Dakota', 'OH': 'Ohio',
    'OK': 'Oklahoma', 'OR': 'Oregon', 'PA': 'Pennsylvania', 'RI': 'Rhode Island',
    'SC': 'South Carolina', 'SD': 'South Dakota', 'TN': 'Tennessee', 'TX': 'Texas',
    'UT': 'Utah', 'VT': 'Vermont', 'VA': 'Virginia', 'WA': 'Washington',
    'WV': 'West Virginia', 'WI': 'Wisconsin', 'WY': 'Wyoming'
}

def expand_state_abbreviations(text):
    \"\"\"Expand state abbreviations to full state names\"\"\"
    if pd.isna(text):
        return text
    
    text_str = str(text)
    
    # Look for state abbreviations (2 letters, possibly with spaces around them)
    for abbr, full_name in STATE_ABBREVIATIONS.items():
        # Match abbreviation at word boundaries
        pattern = r'\\b' + abbr + r'\\b'
        text_str = re.sub(pattern, full_name, text_str, flags=re.IGNORECASE)
    
    return text_str

# Reformat Insurance column to match the expected format
def format_insurance_name(insurance_text):
    if pd.isna(insurance_text):
        return insurance_text
    
    insurance_str = str(insurance_text).strip()
    
    # Handle special cases first
    if insurance_str.upper() == 'NO INSURANCE':
        return 'No Insurance'
    elif insurance_str.upper() == 'PATIENT NOT FOUND':
        return 'PATIENT NOT FOUND'
    elif insurance_str.upper() == 'DUPLICATE':
        return 'DUPLICATE'
    
    # Extract company name before "Ph#"
    if "Ph#" in insurance_str:
        company_name = insurance_str.split("Ph#")[0].strip()
    else:
        company_name = insurance_str
    
    # Remove "Primary" and "Secondary" text
    company_name = re.sub(r'\s*\(Primary\)', '', company_name, flags=re.IGNORECASE)
    company_name = re.sub(r'\s*\(Secondary\)', '', company_name, flags=re.IGNORECASE)
    company_name = re.sub(r'\s*Primary', '', company_name, flags=re.IGNORECASE)
    company_name = re.sub(r'\s*Secondary', '', company_name, flags=re.IGNORECASE)
    
    # Handle Delta Dental variations
    if re.search(r'delta\s+dental', company_name, re.IGNORECASE):
        # Extract state from Delta Dental
        delta_match = re.search(r'delta\s+dental\s+(?:of\s+)?(.+)', company_name, re.IGNORECASE)
        if delta_match:
            state = delta_match.group(1).strip()
            # Expand state abbreviations
            state = expand_state_abbreviations(state)
            return f"DD {state}"
        else:
            return "DD"
    
    # Handle BCBS variations
    if re.search(r'bcbs|blue\s+cross|blue\s+shield', company_name, re.IGNORECASE):
        # Extract state from BCBS
        bcbs_match = re.search(r'(?:bcbs|blue\s+cross\s+blue\s+shield|blue\s+cross|blue\s+shield)\s+(?:of\s+)?(.+)', company_name, re.IGNORECASE)
        if bcbs_match:
            state = bcbs_match.group(1).strip()
            # Expand state abbreviations
            state = expand_state_abbreviations(state)
            return f"BCBS {state}"
        else:
            return "BCBS"
    
    # Handle other specific companies
    if re.search(r'metlife|met\s+life', company_name, re.IGNORECASE):
        return "Metlife"
    elif re.search(r'cigna', company_name, re.IGNORECASE):
        return "Cigna"
    elif re.search(r'aarp', company_name, re.IGNORECASE):
        return "AARP"
    elif re.search(r'uhc|united\s*healthcare|united\s*health\s*care', company_name, re.IGNORECASE):
        return "UHC"
    elif re.search(r'teamcare', company_name, re.IGNORECASE):
        return "Teamcare"
    elif re.search(r'humana', company_name, re.IGNORECASE):
        return "Humana"
    elif re.search(r'aetna', company_name, re.IGNORECASE):
        return "Aetna"
    elif re.search(r'guardian', company_name, re.IGNORECASE):
        return "Guardian"
    elif re.search(r'anthem', company_name, re.IGNORECASE):
        return "Anthem"
    elif re.search(r'g\s*e\s*h\s*a', company_name, re.IGNORECASE):
        return "GEHA"
    elif re.search(r'principal', company_name, re.IGNORECASE):
        return "Principal"
    elif re.search(r'ameritas', company_name, re.IGNORECASE):
        return "Ameritas"
    elif re.search(r'physicians\s+mutual', company_name, re.IGNORECASE):
        return "Physicians Mutual"
    elif re.search(r'mutual\s+of\s+omaha', company_name, re.IGNORECASE):
        return "Mutual Omaha"
    elif re.search(r'sunlife|sun\s+life', company_name, re.IGNORECASE):
        return "Sunlife"
    elif re.search(r'liberty\s+dental', company_name, re.IGNORECASE):
        return "Liberty Dental Plan"
    elif re.search(r'careington', company_name, re.IGNORECASE):
        return "Careington Benefit Solutions"
    elif re.search(r'automated\s+benefit', company_name, re.IGNORECASE):
        return "Automated Benefit Services Inc"
    elif re.search(r'network\s+health', company_name, re.IGNORECASE):
        return "Network Health Wisconsin"
    elif re.search(r'regence', company_name, re.IGNORECASE):
        return "REGENCE BCBS"
    elif re.search(r'united\s+concordia', company_name, re.IGNORECASE):
        return "United Concordia"
    elif re.search(r'medical\s+mutual', company_name, re.IGNORECASE):
        return "Medical Mutual"
    elif re.search(r'blue\s+care\s+dental', company_name, re.IGNORECASE):
        return "Blue Care Dental"
    elif re.search(r'dominion\s+dental', company_name, re.IGNORECASE):
        return "Dominion Dental"
    elif re.search(r'carefirst', company_name, re.IGNORECASE):
        return "CareFirst BCBS"
    elif re.search(r'health\s+partners', company_name, re.IGNORECASE):
        return "Health Partners"
    elif re.search(r'keenan', company_name, re.IGNORECASE):
        return "Keenan"
    elif re.search(r'wilson\s+mcshane', company_name, re.IGNORECASE):
        return "Wilson McShane- Delta Dental"
    elif re.search(r'standard\s+(?:life\s+)?insurance', company_name, re.IGNORECASE):
        return "Standard Life Insurance"
    elif re.search(r'plan\s+for\s+health', company_name, re.IGNORECASE):
        return "Plan for Health"
    elif re.search(r'kansas\s+city', company_name, re.IGNORECASE):
        return "Kansas City"
    elif re.search(r'the\s+guardian', company_name, re.IGNORECASE):
        return "The Guardian"
    elif re.search(r'community\s+dental', company_name, re.IGNORECASE):
        return "Community Dental Associates"
    elif re.search(r'northeast\s+delta\s+dental', company_name, re.IGNORECASE):
        return "Northeast Delta Dental"
    elif re.search(r'say\s+cheese\s+dental', company_name, re.IGNORECASE):
        return "SAY CHEESE DENTAL NETWORK"
    elif re.search(r'dentaquest', company_name, re.IGNORECASE):
        return "Dentaquest"
    elif re.search(r'umr', company_name, re.IGNORECASE):
        return "UMR"
    elif re.search(r'mhbp', company_name, re.IGNORECASE):
        return "MHBP"
    elif re.search(r'united\s+states\s+army', company_name, re.IGNORECASE):
        return "United States Army"
    elif re.search(r'conversion\s+default', company_name, re.IGNORECASE):
        return "CONVERSION DEFAULT - Do NOT Delete! Change Pt Ins!"
    elif re.search(r'equitable', company_name, re.IGNORECASE):
        return "Equitable"
    elif re.search(r'manhattan\s+life', company_name, re.IGNORECASE):
        return "Manhattan Life"
    
    # If no specific pattern matches, return the cleaned company name
    return company_name.strip()

# Apply the reformatting
df['Insurance New'] = df['Insurance'].apply(format_insurance_name)

print("✅ Insurance column reformatted to match expected format!")
print("Sample of original vs reformatted:")
sample_df = df[['Insurance', 'Insurance New']].head(15)
print(sample_df.to_string(index=False))
print(f"\\nTotal reformatted entries: {df['Insurance New'].notna().sum()}")
print("\\nUnique reformatted values:")
print(df['Insurance New'].value_counts().head(25))
"""
    
    elif "count" in instruction:
        if "insurance" in instruction:
            return "print('Insurance counts:')\nprint(df['Insurance'].value_counts())"
        elif "office" in instruction:
            return "print('Office counts:')\nprint(df['Office Name'].value_counts())"
        elif "provider" in instruction:
            return "print('Provider counts:')\nprint(df['Provider Name'].value_counts())"
        else:
            return "print(f'Total records: {len(df)}')"
    
    elif "filter" in instruction:
        if "office" in instruction:
            return "print('Available offices:')\nprint(df['Office Name'].value_counts().head(10))"
        elif "insurance" in instruction:
            return "print('Available insurance types:')\nprint(df['Insurance'].value_counts().head(10))"
        else:
            return "print('Available columns for filtering:', list(df.columns))"
    
    elif "summary" in instruction or "report" in instruction:
        return """
print("=== SUMMARY REPORT ===")
print(f"Total records: {len(df)}")
if 'Appoinment Date' in df.columns:
    print(f"Date range: {df['Appoinment Date'].min()} to {df['Appoinment Date'].max()}")
if 'Office Name' in df.columns:
    print(f"Unique offices: {df['Office Name'].nunique()}")
if 'Provider Name' in df.columns:
    print(f"Unique providers: {df['Provider Name'].nunique()}")
if 'Patient ID' in df.columns:
    print(f"Unique patients: {df['Patient ID'].nunique()}")
print("\\nTop 5 Insurance types:")
if 'Insurance' in df.columns:
    print(df['Insurance'].value_counts().head())
print("\\nTop 5 Offices:")
if 'Office Name' in df.columns:
    print(df['Office Name'].value_counts().head())
"""
    
    else:
        # Try to handle complex instructions with better error handling
        return f"""
try:
    print("Processing complex instruction: {instruction}")
    print("Available columns:", list(df.columns))
    
    # For complex instructions, show sample data first
    print("\\nSample data from Insurance column:")
    if 'Insurance' in df.columns:
        print(df['Insurance'].head(10).to_string())
    else:
        print("Insurance column not found. Available columns:", list(df.columns))
    
    print("\\nFor complex data transformations, try these specific commands:")
    print("- 'reformat insurance column' - Clean up insurance names")
    print("- 'show first 10 rows' - Display sample data")
    print("- 'count insurance types' - Count unique insurance types")
    print("- 'copy Insurance to Insurance New' - Copy column data")
    
except Exception as e:
    print(f"Error processing instruction: {{e}}")
    print("Please try a simpler instruction or use one of the suggested commands above.")
"""

def extract_number(text):
    """Extract number from text"""
    numbers = re.findall(r'\d+', text)
    return int(numbers[0]) if numbers else None

@app.route('/')
def index():
    global current_data, current_sheet, current_filename, conversation_history
    
    return render_template_string(HTML_TEMPLATE, 
                                current_data=current_data, 
                                current_sheet=current_sheet, 
                                current_filename=current_filename,
                                output="",
                                comparison_url=os.environ.get('COMPARISON_URL', 'http://localhost:5002/comparison'))

@app.route('/upload', methods=['POST'])
def upload_file():
    global current_data, current_sheet, current_filename, conversation_history
    
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    try:
        # Save uploaded file
        filename = secure_filename(file.filename)
        file.save(filename)
        
        # Load Excel file
        current_data = pd.read_excel(filename, sheet_name=None)
        current_filename = filename
        conversation_history = []
        
        # Set the main sheet as current
        if 'Consolidated' in current_data:
            current_sheet = 'Consolidated'
        else:
            current_sheet = list(current_data.keys())[0]
        
        return redirect(url_for('index'))
        
    except Exception as e:
        return jsonify({'error': f'Error uploading file: {str(e)}'}), 500

@app.route('/execute', methods=['POST'])
def execute_instruction():
    global current_data, current_sheet, conversation_history
    
    if not current_data:
        return jsonify({'error': 'No file loaded'}), 400
    
    instruction = request.form.get('instruction', '').strip()
    if not instruction:
        return jsonify({'error': 'No instruction provided'}), 400
    
    try:
        # Add to conversation history
        conversation_history.append({
            'timestamp': datetime.now(),
            'instruction': instruction,
            'sheet': current_sheet
        })
        
        # Get current dataframe
        df = current_data[current_sheet].copy()
        
        # Generate code
        code = process_instruction(instruction, df)
        
        # Execute code
        import io
        import sys
        from contextlib import redirect_stdout
        
        output_buffer = io.StringIO()
        
        with redirect_stdout(output_buffer):
            exec(code, {'df': df, 'pd': pd, 'print': print})
        
        output = output_buffer.getvalue()
        
        # Update data if modified
        if 'df' in locals():
            current_data[current_sheet] = df
        
        return render_template_string(HTML_TEMPLATE, 
                                   current_data=current_data, 
                                   current_sheet=current_sheet, 
                                   current_filename=current_filename,
                                   output=output)
        
    except Exception as e:
        error_output = f"Error executing instruction: {str(e)}"
        return render_template_string(HTML_TEMPLATE, 
                                   current_data=current_data, 
                                   current_sheet=current_sheet, 
                                   current_filename=current_filename,
                                   output=error_output)

@app.route('/switch_sheet', methods=['POST'])
def switch_sheet():
    global current_data, current_sheet
    
    if not current_data:
        return jsonify({'error': 'No file loaded'}), 400
    
    data = request.get_json()
    sheet_name = data.get('sheet')
    
    if not sheet_name:
        return jsonify({'error': 'No sheet name provided'}), 400
    
    if sheet_name in current_data:
        current_sheet = sheet_name
        return jsonify({'success': True, 'current_sheet': current_sheet})
    else:
        return jsonify({'error': f'Sheet "{sheet_name}" not found'}), 400

@app.route('/export', methods=['POST'])
def export_data():
    global current_data, current_filename
    
    if not current_data:
        return jsonify({'error': 'No file loaded'}), 400
    
    filename = request.form.get('filename', '').strip()
    if not filename:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"processed_data_{timestamp}.xlsx"
    
    try:
        # Create a temporary file in memory instead of saving to disk
        import tempfile
        import os
        
        # Create temporary file
        temp_fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
        
        try:
            with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
                for sheet_name, df in current_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            return send_file(temp_path, as_attachment=True, download_name=filename)
            
        finally:
            # Clean up temporary file
            os.close(temp_fd)
            if os.path.exists(temp_path):
                os.unlink(temp_path)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/reset', methods=['POST'])
def reset_app():
    global current_data, current_sheet, current_filename, conversation_history
    
    try:
        # Reset all data
        current_data = {}
        current_sheet = None
        current_filename = None
        conversation_history = []
        
        return jsonify({
            'success': True, 
            'message': 'App reset successfully. Please upload a new Excel file to continue.'
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/comparison')
def redirect_to_comparison():
    """Redirect to comparison tool"""
    from flask import redirect
    return redirect(os.environ.get('COMPARISON_URL', 'http://localhost:5002/comparison'))

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5001))
    debug = os.environ.get('FLASK_ENV') == 'development'
    
    print("🚀 Starting AI Excel Automation Web App...")
    print(f"📱 Open your browser and go to: http://localhost:{port}")
    app.run(debug=debug, host='0.0.0.0', port=port)
