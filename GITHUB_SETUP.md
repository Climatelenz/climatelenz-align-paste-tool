# GitHub Setup Guide for ClimateLenz Align & Paste Tool

## Step 1: Create a GitHub Repository

1. Go to [github.com](https://github.com) and sign in
2. Click the **+** icon (top right) → **New repository**
3. Repository name: `climatelenz-align-paste-tool`
4. Description: `Web tool to align copied material into Word, Excel, and PowerPoint documents for CSRD/ESRS reporting`
5. Visibility: **Private** (recommended for now)
6. Check: **Add a README file** (we'll replace it with ours)
7. Click **Create repository**

## Step 2: Upload Your Code

### Option A: Using Git Command Line

```bash
# Navigate to your project folder
cd climatelenz-align-paste-tool

# Initialize git
git init

# Add all files
git add .

# Commit
git commit -m "Initial commit: ClimateLenz Align & Paste Tool"

# Connect to GitHub (replace with your actual repo URL)
git remote add origin https://github.com/YOUR_USERNAME/climatelenz-align-paste-tool.git

# Push
git branch -M main
git push -u origin main