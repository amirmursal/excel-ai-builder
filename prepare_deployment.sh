#!/bin/bash

echo "🚀 Preparing Apps for Railway Deployment..."

# Create directories for each app
mkdir -p main_app_deployment
mkdir -p comparison_app_deployment

echo "📁 Creating Main App deployment files..."
# Main App files
cp web_excel_automation.py main_app_deployment/
cp ai_excel_automation.py main_app_deployment/
cp requirements.txt main_app_deployment/
cp runtime.txt main_app_deployment/
cp Procfile_main main_app_deployment/Procfile

echo "📁 Creating Comparison App deployment files..."
# Comparison App files
cp excel_comparison.py comparison_app_deployment/
cp requirements.txt comparison_app_deployment/
cp runtime.txt comparison_app_deployment/
cp Procfile_comparison comparison_app_deployment/Procfile

echo "✅ Deployment files prepared!"
echo ""
echo "📂 Main App files are in: main_app_deployment/"
echo "📂 Comparison App files are in: comparison_app_deployment/"
echo ""
echo "🚀 Next steps:"
echo "1. Create two separate Git repositories"
echo "2. Copy files from each directory to respective repository"
echo "3. Deploy each repository separately on Railway"
echo ""
echo "📖 See DEPLOYMENT_MAIN_APP.md and DEPLOYMENT_COMPARISON_APP.md for detailed instructions"
