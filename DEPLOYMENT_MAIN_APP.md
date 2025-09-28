# Main Excel Automation App - Railway Deployment

## 🚀 Deploy Main App to Railway

### Step 1: Prepare Repository

1. Create a new Git repository for the main app
2. Copy these files to the repository:
   - `web_excel_automation.py`
   - `ai_excel_automation.py`
   - `requirements.txt`
   - `runtime.txt`
   - `Procfile_main` (rename to `Procfile`)

### Step 2: Rename Procfile

```bash
mv Procfile_main Procfile
```

### Step 3: Deploy on Railway

1. Go to [Railway.app](https://railway.app)
2. Create new project
3. Connect your Git repository
4. Railway will automatically detect and deploy

### Step 4: Access Your App

- Railway will provide a URL like: `https://your-main-app.railway.app`
- Your main app will be accessible at this URL

## 📁 Files Needed for Main App Deployment:

- ✅ `web_excel_automation.py` (main app)
- ✅ `ai_excel_automation.py` (core functionality)
- ✅ `requirements.txt` (dependencies)
- ✅ `runtime.txt` (Python version)
- ✅ `Procfile` (deployment command)

## 🎯 Features Available:

- Excel file upload and processing
- Natural language instructions
- Insurance name formatting
- Sheet switching
- Data export
- All original functionality
