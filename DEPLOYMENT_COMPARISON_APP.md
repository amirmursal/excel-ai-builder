# Excel Comparison Tool - Railway Deployment

## 🚀 Deploy Comparison Tool to Railway

### Step 1: Prepare Repository

1. Create a new Git repository for the comparison tool
2. Copy these files to the repository:
   - `excel_comparison.py`
   - `requirements.txt`
   - `runtime.txt`
   - `Procfile_comparison` (rename to `Procfile`)

### Step 2: Rename Procfile

```bash
mv Procfile_comparison Procfile
```

### Step 3: Deploy on Railway

1. Go to [Railway.app](https://railway.app)
2. Create new project
3. Connect your Git repository
4. Railway will automatically detect and deploy

### Step 4: Access Your App

- Railway will provide a URL like: `https://your-comparison-app.railway.app`
- Your comparison tool will be accessible at: `https://your-comparison-app.railway.app/comparison`

## 📁 Files Needed for Comparison Tool Deployment:

- ✅ `excel_comparison.py` (comparison tool)
- ✅ `requirements.txt` (dependencies)
- ✅ `runtime.txt` (Python version)
- ✅ `Procfile` (deployment command)

## 🎯 Features Available:

- Upload raw Excel file
- Upload previous Excel file
- Compare Patient Names between files
- Add Status column with "Done" for matches
- Download comparison results
- Reset application functionality
