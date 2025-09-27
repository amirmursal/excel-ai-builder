#!/usr/bin/env python3
"""
AI-Powered Excel Automation Program
This program takes an Excel file as input and performs automation based on natural language instructions.
"""

import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import requests
import json
import os
from datetime import datetime
import re

class AIExcelAutomation:
    def __init__(self, excel_file_path, api_key=None):
        self.excel_file_path = excel_file_path
        self.api_key = api_key
        self.data = {}
        self.current_sheet = None
        self.conversation_history = []
        
        # Load the Excel file
        self.load_excel_file()
        
    def load_excel_file(self):
        """Load all sheets from the Excel file"""
        try:
            # Load all sheets
            self.data = pd.read_excel(self.excel_file_path, sheet_name=None)
            print(f"✅ Loaded Excel file with {len(self.data)} sheets:")
            for sheet_name in self.data.keys():
                print(f"   - {sheet_name}: {self.data[sheet_name].shape}")
            
            # Set the main sheet (Consolidated) as current by default
            if 'Consolidated' in self.data:
                self.current_sheet = 'Consolidated'
            else:
                self.current_sheet = list(self.data.keys())[0]
                
            print(f"📊 Current sheet: {self.current_sheet}")
            
        except Exception as e:
            print(f"❌ Error loading Excel file: {e}")
            raise
    
    def ask_ai(self, instruction):
        """Ask AI for code to execute the instruction"""
        if not self.api_key:
            return self.get_basic_instruction_code(instruction)
        
        try:
            # Prepare the context
            current_df = self.data[self.current_sheet]
            context = f"""
            You are an Excel automation expert. You have access to a pandas DataFrame called 'current_df' with the following structure:
            
            Sheet: {self.current_sheet}
            Shape: {current_df.shape}
            Columns: {list(current_df.columns)}
            Data types: {current_df.dtypes.to_dict()}
            
            Sample data:
            {current_df.head(3).to_string()}
            
            Available variables:
            - current_df: The current DataFrame
            - self.data: Dictionary of all sheets
            - self.current_sheet: Current sheet name
            
            User instruction: {instruction}
            
            Generate Python code that accomplishes this task. Only output executable Python code, no explanations.
            """
            
            # Call AI API (using a simple approach - you can replace with your preferred AI service)
            response = self.call_ai_api(context)
            return response
            
        except Exception as e:
            print(f"⚠️ AI API error: {e}")
            return self.get_basic_instruction_code(instruction)
    
    def call_ai_api(self, prompt):
        """Call AI API (placeholder - replace with your preferred AI service)"""
        # This is a placeholder. Replace with your AI API call
        # For example, OpenAI, Anthropic, or local model
        return self.get_basic_instruction_code(prompt.split("User instruction: ")[-1])
    
    def get_basic_instruction_code(self, instruction):
        """Generate basic code for common instructions without AI"""
        instruction = instruction.lower().strip()
        
        if "show" in instruction or "display" in instruction:
            if "first" in instruction:
                num = self.extract_number(instruction) or 5
                return f"print(current_df.head({num}))"
            elif "last" in instruction:
                num = self.extract_number(instruction) or 5
                return f"print(current_df.tail({num}))"
            else:
                return "print(current_df.head(10))"
        
        elif "info" in instruction or "describe" in instruction:
            return """
print("=== DATA INFO ===")
print(f"Shape: {current_df.shape}")
print(f"Columns: {list(current_df.columns)}")
print("\\nData types:")
print(current_df.dtypes)
print("\\nMissing values:")
print(current_df.isnull().sum())
print("\\nBasic statistics:")
print(current_df.describe())
"""
        
        elif "filter" in instruction:
            if "insurance" in instruction:
                if "no insurance" in instruction:
                    return "filtered_df = current_df[current_df['Insurance'] == 'No Insurance']\nprint(f'Filtered {len(filtered_df)} rows')\nprint(filtered_df.head())"
                else:
                    return "print('Available insurance types:')\nprint(current_df['Insurance'].value_counts().head(10))"
            elif "date" in instruction:
                return "print('Date range:')\nprint(f'From: {current_df[\"Appoinment Date\"].min()}')\nprint(f'To: {current_df[\"Appoinment Date\"].max()}')"
            else:
                return "print('Available columns for filtering:')\nprint(list(current_df.columns))"
        
        elif "count" in instruction:
            if "insurance" in instruction:
                return "print('Insurance counts:')\nprint(current_df['Insurance'].value_counts())"
            elif "office" in instruction:
                return "print('Office counts:')\nprint(current_df['Office Name'].value_counts())"
            elif "provider" in instruction:
                return "print('Provider counts:')\nprint(current_df['Provider Name'].value_counts())"
            else:
                return "print(f'Total records: {len(current_df)}')"
        
        elif "summary" in instruction or "report" in instruction:
            return """
print("=== SUMMARY REPORT ===")
print(f"Total appointments: {len(current_df)}")
print(f"Date range: {current_df['Appoinment Date'].min()} to {current_df['Appoinment Date'].max()}")
print(f"Unique offices: {current_df['Office Name'].nunique()}")
print(f"Unique providers: {current_df['Provider Name'].nunique()}")
print(f"Unique patients: {current_df['Patient ID'].nunique()}")
print("\\nTop 5 Insurance types:")
print(current_df['Insurance'].value_counts().head())
print("\\nTop 5 Offices:")
print(current_df['Office Name'].value_counts().head())
"""
        
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
current_df['Insurance New'] = current_df['Insurance'].apply(format_insurance_name)

print("✅ Insurance column reformatted to match expected format!")
print("Sample of original vs reformatted:")
sample_df = current_df[['Insurance', 'Insurance New']].head(15)
print(sample_df.to_string(index=False))
print(f"\\nTotal reformatted entries: {current_df['Insurance New'].notna().sum()}")
print("\\nUnique reformatted values:")
print(current_df['Insurance New'].value_counts().head(25))
"""
        
        elif "export" in instruction or "save" in instruction:
            return """
# Export current data
output_file = f'processed_data_{datetime.now().strftime(\"%Y%m%d_%H%M%S\")}.xlsx'
current_df.to_excel(output_file, index=False)
print(f'Data exported to {output_file}')
"""
        
        else:
            return f"""
print("Available commands:")
print("- show/display: Show data")
print("- info/describe: Show data information")
print("- filter: Filter data")
print("- count: Count records")
print("- summary/report: Generate summary")
print("- export/save: Export data")
print("\\nYour instruction: {instruction}")
"""
    
    def extract_number(self, text):
        """Extract number from text"""
        numbers = re.findall(r'\d+', text)
        return int(numbers[0]) if numbers else None
    
    def execute_instruction(self, instruction):
        """Execute the given instruction"""
        print(f"\n🤖 Processing instruction: {instruction}")
        
        # Add to conversation history
        self.conversation_history.append({
            'timestamp': datetime.now(),
            'instruction': instruction,
            'sheet': self.current_sheet
        })
        
        try:
            # Get code from AI or basic processor
            code = self.ask_ai(instruction)
            
            # Prepare execution environment
            current_df = self.data[self.current_sheet].copy()
            
            # Create execution context
            exec_globals = {
                'pd': pd,
                'current_df': current_df,
                'self': self,
                'datetime': datetime,
                'print': print
            }
            
            # Execute the code
            print("📝 Generated code:")
            print("-" * 40)
            print(code)
            print("-" * 40)
            
            exec(code, exec_globals)
            
            # Update the data if it was modified
            if 'current_df' in exec_globals:
                self.data[self.current_sheet] = exec_globals['current_df']
            
            print("✅ Instruction executed successfully!")
            
        except Exception as e:
            print(f"❌ Error executing instruction: {e}")
            print("💡 Try rephrasing your instruction or use one of the basic commands.")
    
    def switch_sheet(self, sheet_name):
        """Switch to a different sheet"""
        if sheet_name in self.data:
            self.current_sheet = sheet_name
            print(f"📊 Switched to sheet: {sheet_name}")
            print(f"   Shape: {self.data[sheet_name].shape}")
        else:
            print(f"❌ Sheet '{sheet_name}' not found. Available sheets: {list(self.data.keys())}")
    
    def list_sheets(self):
        """List all available sheets"""
        print("📋 Available sheets:")
        for i, sheet_name in enumerate(self.data.keys(), 1):
            shape = self.data[sheet_name].shape
            current = " (current)" if sheet_name == self.current_sheet else ""
            print(f"   {i}. {sheet_name}: {shape[0]} rows × {shape[1]} columns{current}")
    
    def export_data(self, filename=None):
        """Export current data to Excel"""
        if not filename:
            filename = f"processed_{self.excel_file_path.split('/')[-1].replace('.xlsx', '')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                for sheet_name, df in self.data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"✅ Data exported to: {filename}")
            return filename
            
        except Exception as e:
            print(f"❌ Error exporting data: {e}")
            return None
    
    def interactive_mode(self):
        """Start interactive mode"""
        print("\n🚀 AI Excel Automation - Interactive Mode")
        print("=" * 50)
        print("Available commands:")
        print("  - Type any instruction (e.g., 'show first 10 rows')")
        print("  - 'switch [sheet_name]' - Switch to different sheet")
        print("  - 'list' - List all sheets")
        print("  - 'export' - Export current data")
        print("  - 'quit' - Exit program")
        print("=" * 50)
        
        while True:
            try:
                instruction = input(f"\n[{self.current_sheet}] Enter instruction: ").strip()
                
                if instruction.lower() in ['quit', 'exit', 'q']:
                    print("👋 Goodbye!")
                    break
                elif instruction.lower() == 'list':
                    self.list_sheets()
                elif instruction.lower().startswith('switch '):
                    sheet_name = instruction[7:].strip()
                    self.switch_sheet(sheet_name)
                elif instruction.lower() == 'export':
                    self.export_data()
                elif instruction:
                    self.execute_instruction(instruction)
                else:
                    print("Please enter an instruction or command.")
                    
            except KeyboardInterrupt:
                print("\n👋 Goodbye!")
                break
            except Exception as e:
                print(f"❌ Error: {e}")

def main():
    """Main function"""
    print("🤖 AI Excel Automation Program")
    print("=" * 40)
    
    # Excel file path
    excel_file = "Imagen-IV-CR-August-2025.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"❌ Excel file '{excel_file}' not found!")
        return
    
    # Initialize the automation
    automation = AIExcelAutomation(excel_file)
    
    # Start interactive mode
    automation.interactive_mode()

if __name__ == "__main__":
    main()
