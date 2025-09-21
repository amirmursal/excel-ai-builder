#!/usr/bin/env python3
"""
Demo script showing how to use the AI Excel Automation
"""

from ai_excel_automation import AIExcelAutomation

def main():
    print("🤖 AI Excel Automation Demo")
    print("=" * 40)
    
    # Initialize with your Excel file
    excel_file = "Imagen-IV-CR-August-2025.xlsx"
    automation = AIExcelAutomation(excel_file)
    
    print("\n📊 Let's explore the data with some sample instructions:")
    print("-" * 50)
    
    # Sample instructions
    instructions = [
        "show first 5 rows",
        "show data info",
        "count insurance types",
        "generate summary report",
        "filter by office name"
    ]
    
    for instruction in instructions:
        print(f"\n🔍 Instruction: {instruction}")
        automation.execute_instruction(instruction)
        print("\n" + "="*50)
    
    print("\n✅ Demo completed!")
    print("💡 To use interactively, run: python ai_excel_automation.py")
    print("🌐 To use web interface, run: python web_automation.py")

if __name__ == "__main__":
    main()
