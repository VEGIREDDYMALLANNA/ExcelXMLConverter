# ExcelXML Converter — Complete Git Setup Guide

## Project Location

The project is currently local at:

```text
C:\XMLConverter
```

Recommended structure:

```text
C:\XMLConverter
│
├── main.py
├── excelxml.ico
├── requirements.txt
├── .gitignore
│
└── docs
    └── ExcelXML_Converter_Build_Deployment_Guide.md
```

Generated PyInstaller folders such as `build/` and `dist/` should normally NOT be committed.

---

# 1. Open the Project

Open Command Prompt or PowerShell:

```cmd
cd C:\XMLConverter
```

Check the files:

```cmd
dir
```

---

# 2. Create `.gitignore`

Create the file:

```cmd
notepad .gitignore
```

Paste:

```gitignore
# Python
__pycache__/
*.py[cod]

# Virtual environments
venv/
.venv/

# PyInstaller
build/
dist/
*.spec

# IDE
.vscode/

# Logs
*.log
```

Save and close Notepad.

Check:

```cmd
dir /a
```

You should see `.gitignore`.

---

# 3. Initialize Git

Run:

```cmd
git init
```

Expected result:

```text
Initialized empty Git repository in C:/XMLConverter/.git/
```

This creates:

```text
C:\XMLConverter\.git
```

The `.git` directory contains your local Git history and configuration.

---

# 4. Check Git Status

```cmd
git status
```

You should see your project files as untracked.

You should normally NOT see:

```text
build/
dist/
__pycache__/
```

because they are ignored.

---

# 5. Configure Git Identity

If Git is not already configured:

```cmd
git config --global user.name "Mallanna Vegireddy"
```

Set the email associated with your GitHub account:

```cmd
git config --global user.email "YOUR_GITHUB_EMAIL"
```

Check:

```cmd
git config --global --list
```

---

# 6. Create the Documentation Folder

Create:

```cmd
mkdir docs
```

Move the documentation file:

```cmd
move ExcelXML_Converter_Build_Deployment_Guide.md docs\
```

The project should now look like:

```text
C:\XMLConverter
│
├── main.py
├── excelxml.ico
├── requirements.txt
├── .gitignore
│
├── docs
│   └── ExcelXML_Converter_Build_Deployment_Guide.md
│
├── build\
└── dist\
```

`build` and `dist` will be ignored by Git.

---

# 7. Add the Project to Git

Check first:

```cmd
git status
```

Then stage the files:

```cmd
git add .
```

Check what is staged:

```cmd
git status
```

You should see files such as:

```text
new file: .gitignore
new file: main.py
new file: excelxml.ico
new file: requirements.txt
new file: docs/ExcelXML_Converter_Build_Deployment_Guide.md
```

---

# 8. Create the First Commit

Run:

```cmd
git commit -m "Initial commit - ExcelXML Converter"
```

Check the history:

```cmd
git log --oneline
```

Example:

```text
abc1234 Initial commit - ExcelXML Converter
```

---

# 9. Create the GitHub Repository

Open:

```text
https://github.com/
```

Sign in.

Select:

```text
+ → New repository
```

Use:

**Repository name**

```text
ExcelXMLConverter
```

**Description**

```text
Windows desktop application for converting Excel worksheets to XML
```

Choose `Public` or `Private`.

For a portfolio project, `Public` can be useful.

## Important

Do NOT select:

```text
☐ Add a README file
☐ Add .gitignore
☐ Choose a license
```

The local project already contains the required files.

Click:

```text
Create repository
```

---

# 10. Copy the GitHub Repository URL

GitHub will show a URL similar to:

```text
https://github.com/YOUR_USERNAME/ExcelXMLConverter.git
```

Use the actual URL from your GitHub repository.

---

# 11. Connect Local Git to GitHub

Return to:

```text
C:\XMLConverter
```

Run:

```cmd
git remote add origin https://github.com/YOUR_USERNAME/ExcelXMLConverter.git
```

Replace `YOUR_USERNAME` with your actual GitHub username.

---

# 12. Verify the Remote

Run:

```cmd
git remote -v
```

Expected:

```text
origin  https://github.com/YOUR_USERNAME/ExcelXMLConverter.git (fetch)
origin  https://github.com/YOUR_USERNAME/ExcelXMLConverter.git (push)
```

---

# 13. Rename the Branch to `main`

Run:

```cmd
git branch -M main
```

Check:

```cmd
git branch
```

Expected:

```text
* main
```

---

# 14. Push the Project to GitHub

Run:

```cmd
git push -u origin main
```

Authenticate with GitHub if requested.

After a successful push, refresh your GitHub repository.

You should see:

```text
ExcelXMLConverter
│
├── .gitignore
├── main.py
├── excelxml.ico
├── requirements.txt
│
└── docs
    └── ExcelXML_Converter_Build_Deployment_Guide.md
```

---

# 15. Complete First-Time Command Sequence

From the local project:

```cmd
cd C:\XMLConverter
```

Create `.gitignore`:

```cmd
notepad .gitignore
```

Then:

```cmd
git init
```

Configure identity if required:

```cmd
git config --global user.name "Mallanna Vegireddy"
git config --global user.email "YOUR_GITHUB_EMAIL"
```

Create documentation folder:

```cmd
mkdir docs
```

Move the documentation:

```cmd
move ExcelXML_Converter_Build_Deployment_Guide.md docs\
```

Check:

```cmd
git status
```

Stage:

```cmd
git add .
```

Check:

```cmd
git status
```

Commit:

```cmd
git commit -m "Initial commit - ExcelXML Converter"
```

Rename branch:

```cmd
git branch -M main
```

Connect GitHub:

```cmd
git remote add origin https://github.com/YOUR_USERNAME/ExcelXMLConverter.git
```

Verify:

```cmd
git remote -v
```

Push:

```cmd
git push -u origin main
```

---

# 16. After the First Push

For future changes, you normally only need:

```cmd
git status
```

```cmd
git add .
```

```cmd
git commit -m "Describe the change"
```

```cmd
git push
```

Example:

```cmd
git add .
git commit -m "Add XML validation"
git push
```

---

# 17. Updating the Documentation

If you modify:

```text
docs\ExcelXML_Converter_Build_Deployment_Guide.md
```

run:

```cmd
git status
git add docs/ExcelXML_Converter_Build_Deployment_Guide.md
git commit -m "Update build and deployment documentation"
git push
```

---

# 18. Useful Git Commands

### Check current status

```cmd
git status
```

### Add all changed files

```cmd
git add .
```

### Commit

```cmd
git commit -m "Your commit message"
```

### Push

```cmd
git push
```

### Pull latest changes

```cmd
git pull
```

### Show branches

```cmd
git branch
```

### Show GitHub remote

```cmd
git remote -v
```

### Show commit history

```cmd
git log --oneline
```

---

# 19. Git Tags for Releases

When the application reaches a stable release:

```cmd
git tag v1.0.0
```

Push the tag:

```cmd
git push origin v1.0.0
```

Later:

```cmd
git tag v1.1.0
git push origin v1.1.0
```

Recommended versions:

```text
v1.0.0
v1.1.0
v1.2.0
```

---

# 20. PyInstaller Files

Your application uses PyInstaller, for example:

```cmd
python -m PyInstaller --onedir --windowed --icon=excelxml.ico --add-data "excelxml.ico;." --name ExcelXMLConverter main.py
```

PyInstaller creates:

```text
build/
dist/
ExcelXMLConverter.spec
```

These are build artifacts and are excluded by `.gitignore`.

Do not normally commit them to the source repository.

---

# 21. Production Release Structure

Keep source code and production releases separate.

GitHub repository:

```text
ExcelXMLConverter
│
├── Source Code
│   ├── main.py
│   ├── excelxml.ico
│   └── requirements.txt
│
├── Documentation
│   └── docs/
│
└── Releases
    └── v1.0.0
        └── ExcelXMLConverter_Setup.exe
```

The final installer can later be distributed through GitHub Releases.

---

# 22. Recommended Repository Structure

The source repository should ideally be:

```text
ExcelXMLConverter/
│
├── main.py
├── excelxml.ico
├── requirements.txt
├── .gitignore
├── README.md
│
└── docs/
    └── ExcelXML_Converter_Build_Deployment_Guide.md
```

Generated files should remain ignored:

```text
build/
dist/
__pycache__/
*.spec
```

---

# 23. Development Workflow

For every future feature:

```text
Modify code
    ↓
Test application
    ↓
git status
    ↓
git add .
    ↓
git commit -m "Description"
    ↓
git push
```

Example:

```cmd
git status
git add .
git commit -m "Add column selection improvements"
git push
```

---

# 24. Final Git Workflow

First setup:

```text
C:\XMLConverter
       ↓
git init
       ↓
Local Git Repository
       ↓
Create GitHub Repository
       ↓
git remote add origin
       ↓
git branch -M main
       ↓
git push -u origin main
       ↓
GitHub
```

Future development:

```text
Change code
    ↓
git status
    ↓
git add .
    ↓
git commit
    ↓
git push
```

---

# 25. Final Checklist

- [ ] Project is at `C:\XMLConverter`
- [ ] `main.py` exists
- [ ] `excelxml.ico` exists
- [ ] `requirements.txt` exists
- [ ] `.gitignore` exists
- [ ] Documentation is under `docs/`
- [ ] `build/` is ignored
- [ ] `dist/` is ignored
- [ ] Git is initialized
- [ ] Git identity is configured
- [ ] GitHub repository is created
- [ ] GitHub remote is configured
- [ ] Branch is `main`
- [ ] Initial commit is created
- [ ] Project is pushed to GitHub
- [ ] GitHub repository has been verified

---

# 26. Final Goal

Your professional project should eventually look like:

```text
GitHub Repository
        │
        ├── Source Code
        │     ├── main.py
        │     ├── excelxml.ico
        │     └── requirements.txt
        │
        ├── Documentation
        │     └── docs/
        │
        ├── Git History
        │
        └── Releases
              └── v1.0.0
                    └── ExcelXMLConverter_Setup.exe
```

This keeps your **source code, documentation, Git history, and production releases organized separately**.
