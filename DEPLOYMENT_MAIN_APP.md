# Main Excel AI Builder - Deployment Guide

## Current Status

Your main app is already deployed at: `https://web-production-9e92a.up.railway.app`

## After Deploying Comparison Tool

Once you have deployed the comparison tool and have its URL:

1. **Get the Comparison Tool URL:**

   - It will look like: `https://your-comparison-app.railway.app`

2. **Update Environment Variable:**

   - Go to your Railway project dashboard
   - Go to "Variables" tab
   - Add: `COMPARISON_URL` = `https://your-comparison-app.railway.app/comparison`

3. **Redeploy (if needed):**
   - Railway will automatically redeploy when you add environment variables
   - Or you can trigger a manual redeploy

## Testing the Integration

1. **Test Main App:**

   - Visit: `https://web-production-9e92a.up.railway.app`
   - Click "📊 Comparison Tool" button
   - Should open the comparison tool in a new tab

2. **Test Comparison Tool:**
   - In the comparison tool, click "🏠 Main App" button
   - Should redirect back to the main app

## Local Development

To test locally with both apps:

1. **Start Main App:**

   ```bash
   export COMPARISON_URL="http://localhost:5002/comparison"
   python web_excel_automation.py
   ```

2. **Start Comparison Tool:**
   ```bash
   cd ../excel-comparison-tool
   export MAIN_APP_URL="http://localhost:5001/"
   python excel_comparison.py
   ```

## Environment Variables

- `COMPARISON_URL`: URL of the comparison tool (default: http://localhost:5002/comparison)

## File Structure

```
excel-ai-builder/          # Main app repository
├── web_excel_automation.py
├── requirements.txt
├── runtime.txt
├── Procfile
└── DEPLOYMENT_MAIN_APP.md

excel-comparison-tool/     # Comparison tool repository
├── excel_comparison.py
├── requirements.txt
├── runtime.txt
├── Procfile
└── DEPLOYMENT.md
```
