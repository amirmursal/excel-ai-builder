# Excel Comparison Tool - Railway Deployment Guide

## 🚀 Deploy the Comparison Tool to Railway

### Step 1: Create New Railway Project

1. Go to [Railway.app](https://railway.app)
2. Click "New Project"
3. Select "Deploy from GitHub repo" or "Deploy from folder"

### Step 2: Upload Files

Upload these files to your new Railway project:

- `excel_comparison.py` (main application)
- `requirements.txt` (dependencies)
- `runtime.txt` (Python version)
- `Procfile` (start command)

### Step 3: Configure Environment

1. In Railway dashboard, go to your project
2. Click on "Variables" tab
3. Add these environment variables:
   - `MAIN_APP_URL`: `https://web-production-9e92a.up.railway.app/`
   - `PORT`: `5002` (optional, Railway will set this automatically)

### Step 4: Deploy

1. Railway will automatically detect Python and install dependencies
2. The app will start using the Procfile command
3. Railway will provide you with a URL like: `https://your-comparison-app.railway.app`

### Step 5: Update Main App

After getting the comparison tool URL, update the main app:

1. Go to your main app's Railway project
2. Add environment variable: `COMPARISON_URL`: `https://your-comparison-app.railway.app/comparison`
3. Redeploy the main app

### Step 6: Test

1. Main app: `https://web-production-9e92a.up.railway.app/`
2. Comparison tool: `https://your-comparison-app.railway.app/comparison`
3. Test navigation between both apps

## 📁 Files Included

- `excel_comparison.py` - Main comparison tool application
- `requirements.txt` - Python dependencies
- `runtime.txt` - Python version specification
- `Procfile` - Railway start command
- `DEPLOYMENT_COMPARISON_TOOL.md` - This guide

## 🔗 Navigation

- The comparison tool will have a "🏠 Main App" button that links back to the main app
- The main app will have a "📊 Comparison Tool" button that links to the comparison tool
