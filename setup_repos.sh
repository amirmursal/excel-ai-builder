#!/bin/bash

echo "🚀 Setting up Excel AI Builder with separate repositories"
echo "========================================================"
echo ""

# Colors
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m'

print_step() {
    echo -e "${BLUE}Step $1:${NC} $2"
}

print_info() {
    echo -e "${GREEN}ℹ️  $1${NC}"
}

print_warning() {
    echo -e "${YELLOW}⚠️  $1${NC}"
}

echo "This script will help you set up two separate repositories:"
echo "1. Main Excel AI Builder (current repo)"
echo "2. Excel Comparison Tool (new repo)"
echo ""

# Step 1: Check current status
print_step "1" "Checking current setup"
echo ""

if [ -d "../excel-comparison-tool" ]; then
    print_info "✅ Comparison tool repository already created"
    echo "   Location: ../excel-comparison-tool/"
else
    print_warning "❌ Comparison tool repository not found"
    echo "   Please run the previous setup first"
    exit 1
fi

# Step 2: Show repository structure
print_step "2" "Repository structure"
echo ""
echo "📁 Main App (current directory):"
echo "   - web_excel_automation.py"
echo "   - requirements.txt"
echo "   - runtime.txt"
echo "   - Procfile"
echo "   - Already deployed at: https://web-production-9e92a.up.railway.app"
echo ""
echo "📁 Comparison Tool (../excel-comparison-tool/):"
echo "   - excel_comparison.py"
echo "   - requirements.txt"
echo "   - runtime.txt"
echo "   - Procfile"
echo "   - Ready for deployment"
echo ""

# Step 3: Next steps
print_step "3" "Next steps"
echo ""
echo "🔧 To complete the setup:"
echo ""
echo "1. Push comparison tool to GitHub:"
echo "   cd ../excel-comparison-tool"
echo "   git remote add origin https://github.com/YOUR_USERNAME/excel-comparison-tool.git"
echo "   git push -u origin main"
echo ""
echo "2. Deploy comparison tool on Railway:"
echo "   - Go to https://railway.app"
echo "   - Create new project from GitHub repo"
echo "   - Select excel-comparison-tool"
echo "   - Set environment variable: MAIN_APP_URL = https://web-production-9e92a.up.railway.app/"
echo ""
echo "3. Update main app with comparison tool URL:"
echo "   - In Railway dashboard, go to your main app project"
echo "   - Add environment variable: COMPARISON_URL = https://your-comparison-app.railway.app/comparison"
echo ""
echo "4. Test the integration:"
echo "   - Visit main app: https://web-production-9e92a.up.railway.app"
echo "   - Click 'Comparison Tool' button"
echo "   - Should open comparison tool in new tab"
echo ""

print_info "🎉 Setup complete! Follow the steps above to deploy both apps."
