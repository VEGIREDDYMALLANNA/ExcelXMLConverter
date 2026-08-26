# ExcelXML Converter — Build & Deployment Guide

## 1. Overview

**ExcelXML Converter** is a Windows desktop application built with Python and Tkinter. It converts Excel worksheets into XML and can be packaged as a standalone Windows application using **PyInstaller**.

This guide covers:

- Project structure
- Python dependencies
- Running the application
- One-file PyInstaller build
- One-directory PyInstaller build
- Application icon handling
- Clean rebuilds
- Debugging
- Production packaging with an installer
- Testing on another Windows machine
- Antivirus considerations

---

## 2. Project Structure

Recommended project directory:

```text
C:\XMLConverter
│
├── main.py
├── excelxml.ico
├── requirements.txt
│
├── build\
├── dist\
└── installer\
```

| File / Folder | Purpose |
|---|---|
| `main.py` | Main Python application |
| `excelxml.ico` | Windows application icon |
| `requirements.txt` | Python dependencies |
| `build\` | Temporary PyInstaller files |
| `dist\` | Generated application |
| `installer\` | Final installer output |

---

## 3. Open the Project Directory

Open PowerShell or Command Prompt:

```powershell
cd C:\XMLConverter
```

Check Python:

```powershell
python --version
```

Check pip:

```powershell
python -m pip --version
```

---

## 4. Install Dependencies

For the Excel-to-XML application:

```powershell
python -m pip install openpyxl
```

Install PyInstaller:

```powershell
python -m pip install pyinstaller
```

Verify:

```powershell
python -m PyInstaller --version
```

### requirements.txt

Create:

```text
C:\XMLConverter\requirements.txt
```

Example:

```text
openpyxl
pyinstaller
```

Install everything with:

```powershell
python -m pip install -r requirements.txt
```

> `tkinter` is normally included with the standard Windows Python installation and does not need to be added to `requirements.txt`.

---

# 5. Test the Python Application First

Before creating an EXE, run:

```powershell
python main.py
```

Verify:

- Application opens.
- Excel file can be selected.
- Worksheets are detected.
- Data preview works.
- Empty cells are omitted by default.
- `Include Empty` works.
- Deleted columns are omitted.
- XML root and record names work.
- XML export works.
- Generated XML is valid.

Only create the EXE after the Python application works correctly.

---

# 6. Application Icon

Place the icon here:

```text
C:\XMLConverter\excelxml.ico
```

The icon can be used for:

1. EXE icon.
2. Tkinter application/window icon.
3. Installer icon.
4. Desktop shortcut.
5. Start Menu shortcut.

---

# 7. Why Both `--icon` and `--add-data` Are Used

These two options have different purposes.

### `--icon`

```text
--icon=excelxml.ico
```

Sets the Windows icon of the generated EXE.

### `--add-data`

```text
--add-data "excelxml.ico;."
```

Includes the actual `.ico` file inside the packaged application so that your Python code can access it at runtime.

On Windows, PyInstaller uses `;` between the source and destination:

```text
--add-data "source;destination"
```

Therefore:

```text
--add-data "excelxml.ico;."
```

means:

```text
Source:      excelxml.ico
Destination: application root
```

---

# 8. One-File Build

Use this command when you want a single executable:

```powershell
python -m PyInstaller --onefile --windowed --icon=excelxml.ico --add-data "excelxml.ico;." --name ExcelXMLConverter main.py
```

## Meaning of the options

| Option | Meaning |
|---|---|
| `python -m PyInstaller` | Runs PyInstaller from the active Python environment |
| `--onefile` | Creates one executable |
| `--windowed` | Hides the console window |
| `--icon=excelxml.ico` | Sets the EXE icon |
| `--add-data "excelxml.ico;."` | Includes the icon as runtime data |
| `--name ExcelXMLConverter` | Sets the application/EXE name |
| `main.py` | Main Python program |

---

# 9. One-File Output

After a successful build:

```text
C:\XMLConverter
│
├── build\
├── dist\
│   └── ExcelXMLConverter.exe
│
├── excelxml.ico
└── main.py
```

The executable is:

```text
C:\XMLConverter\dist\ExcelXMLConverter.exe
```

The user sees only one EXE.

### Important

`--onefile` does not mean the application has no dependencies. PyInstaller packages the Python runtime and required libraries into the executable and extracts runtime components when the program starts.

This can cause:

- Slightly slower startup.
- Larger EXE size.
- More antivirus false positives in some environments.

---

# 10. One-Directory Build

For your production application, the recommended PyInstaller command is:

```powershell
python -m PyInstaller --onedir --windowed --icon=excelxml.ico --add-data "excelxml.ico;." --name ExcelXMLConverter main.py
```

`--onedir` creates a complete application directory instead of one large executable.

---

# 11. One-Directory Output

The result will normally look similar to:

```text
C:\XMLConverter
│
├── dist
│   └── ExcelXMLConverter
│       ├── ExcelXMLConverter.exe
│       └── internal
│           └── ...
│
└── build
```

Depending on your PyInstaller version, runtime files may be placed in an `internal` directory.

### Important

Do **not** copy only:

```text
ExcelXMLConverter.exe
```

The complete:

```text
dist\ExcelXMLConverter
```

directory is required.

---

# 12. One-File vs One-Directory

| Feature | `--onefile` | `--onedir` |
|---|---|---|
| Output | One EXE | EXE + runtime files |
| Easy to share directly | Yes | Less convenient |
| Startup | Usually slower | Usually faster |
| Troubleshooting | More difficult | Easier |
| Installer friendly | Yes | **Recommended** |
| Production deployment | Possible | **Recommended** |

For ExcelXML Converter, use:

```text
--onedir
```

and then package it into a Windows installer.

---

# 13. Clean Previous PyInstaller Builds

Before creating a clean release, remove old build files.

From:

```powershell
cd C:\XMLConverter
```

run:

```powershell
Remove-Item -Recurse -Force build -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force dist -ErrorAction SilentlyContinue
Remove-Item -Force *.spec -ErrorAction SilentlyContinue
```

Then build again.

---

# 14. Recommended Production Build

Run:

```powershell
python -m PyInstaller --onedir --windowed --icon=excelxml.ico --add-data "excelxml.ico;." --name ExcelXMLConverter main.py
```

Then test:

```text
C:\XMLConverter\dist\ExcelXMLConverter\ExcelXMLConverter.exe
```

---

# 15. Debugging a Failed EXE

During development, do not immediately use `--windowed`.

Build without it:

```powershell
python -m PyInstaller --onedir --icon=excelxml.ico --add-data "excelxml.ico;." --name ExcelXMLConverter main.py
```

Run:

```text
dist\ExcelXMLConverter\ExcelXMLConverter.exe
```

The console can display Python exceptions.

Once everything works, rebuild using:

```text
--windowed
```

---

# 16. Runtime Icon Path

If your application accesses `excelxml.ico` from Python, use a resource-path helper:

```python
import os
import sys

def resource_path(filename):
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, filename)
```

Then:

```python
icon_path = resource_path("excelxml.ico")
```

For Tkinter:

```python
root = tk.Tk()

icon_path = resource_path("excelxml.ico")

try:
    root.iconbitmap(icon_path)
except Exception:
    pass
```

Make sure `sys` is imported:

```python
import sys
```

Otherwise you may get:

```text
NameError: name 'sys' is not defined
```

---

# 17. Common Problem — `openpyxl` Not Found

If you get:

```text
ModuleNotFoundError: No module named 'openpyxl'
```

install it:

```powershell
python -m pip install openpyxl
```

Verify:

```powershell
python -c "import openpyxl; print(openpyxl.__version__)"
```

Then rebuild:

```powershell
python -m PyInstaller --onedir --windowed --icon=excelxml.ico --add-data "excelxml.ico;." --name ExcelXMLConverter main.py
```

---

# 18. Common Problem — Icon Not Updated

Verify:

```text
C:\XMLConverter\excelxml.ico
```

exists.

Then clean the previous build:

```powershell
Remove-Item -Recurse -Force build -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force dist -ErrorAction SilentlyContinue
Remove-Item -Force *.spec -ErrorAction SilentlyContinue
```

Build again:

```powershell
python -m PyInstaller --onedir --windowed --icon=excelxml.ico --add-data "excelxml.ico;." --name ExcelXMLConverter main.py
```

Windows Explorer can cache old icons. If the old icon remains, restart Explorer or Windows and check again.

---

# 19. Production Packaging

For a professional application, do not give users the raw PyInstaller `dist` folder.

Instead:

```text
main.py
   ↓
PyInstaller --onedir
   ↓
dist\ExcelXMLConverter\
   ↓
Inno Setup
   ↓
ExcelXMLConverter_Setup.exe
```

The user receives only:

```text
ExcelXMLConverter_Setup.exe
```

The installer places the application and its dependencies under something like:

```text
C:\Program Files\ExcelXML Converter\
```

The user does not need:

- Python
- pip
- openpyxl
- PyInstaller
- VS Code

---

# 20. Suggested Production Folder Structure

Development machine:

```text
C:\XMLConverter
│
├── main.py
├── excelxml.ico
├── requirements.txt
├── ExcelXMLConverter.iss
│
├── build\
│
├── dist\
│   └── ExcelXMLConverter\
│       ├── ExcelXMLConverter.exe
│       └── internal\
│
└── installer\
    └── ExcelXMLConverter_Setup.exe
```

Only distribute:

```text
ExcelXMLConverter_Setup.exe
```

---

# 21. Testing on Another Computer

Before production distribution, test the installer on another Windows PC or VM.

The target machine should not need:

```text
Python
pip
VS Code
openpyxl
PyInstaller
```

Install:

```text
ExcelXMLConverter_Setup.exe
```

Then test:

1. Application launches.
2. Application icon is correct.
3. Excel file selection works.
4. Worksheets are detected.
5. Data preview works.
6. Empty cells are omitted by default.
7. `Include Empty` works.
8. Deleted columns are omitted.
9. XML root name works.
10. XML record name works.
11. XML export works.
12. Generated XML opens correctly.
13. Application can be uninstalled.

---

# 22. Antivirus / Windows Defender

PyInstaller executables can sometimes trigger antivirus or endpoint-security false positives.

Do not simply disable antivirus protection.

If Windows Defender reports a threat:

1. Open **Windows Security**.
2. Go to **Virus & threat protection**.
3. Open **Protection history**.
4. Record the exact detection name.
5. Determine whether it is the installer, your EXE, or a dependency.
6. Investigate before distributing the application.

For professional organizational deployment, code signing and controlled software distribution can also improve trust and reduce reputation-related warnings.

---

# 23. Final Commands

### Test source application

```powershell
python main.py
```

### One-file build

```powershell
python -m PyInstaller --onefile --windowed --icon=excelxml.ico --add-data "excelxml.ico;." --name ExcelXMLConverter main.py
```

Output:

```text
dist\ExcelXMLConverter.exe
```

### One-directory build

```powershell
python -m PyInstaller --onedir --windowed --icon=excelxml.ico --add-data "excelxml.ico;." --name ExcelXMLConverter main.py
```

Output:

```text
dist\ExcelXMLConverter\
```

### Clean production rebuild

```powershell
Remove-Item -Recurse -Force build -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force dist -ErrorAction SilentlyContinue
Remove-Item -Force *.spec -ErrorAction SilentlyContinue

python -m PyInstaller --onedir --windowed --icon=excelxml.ico --add-data "excelxml.ico;." --name ExcelXMLConverter main.py
```

---

# 24. Recommended Release Process

Use this process for every release:

```text
1. Modify main.py
       ↓
2. Test with:
   python main.py
       ↓
3. Clean build:
   Remove-Item build/dist/*.spec
       ↓
4. Build:
   PyInstaller --onedir
       ↓
5. Test EXE
       ↓
6. Create installer
       ↓
7. Test installer on another PC/VM
       ↓
8. Check Windows Defender / endpoint security
       ↓
9. Distribute:
   ExcelXMLConverter_Setup.exe
```

---

# 25. Final Checklist

Before sharing the application:

- [ ] `main.py` works correctly.
- [ ] `openpyxl` is installed.
- [ ] `excelxml.ico` exists.
- [ ] Application icon appears correctly.
- [ ] Excel files load correctly.
- [ ] All worksheets are detected.
- [ ] Empty cells are omitted by default.
- [ ] User can include empty columns when required.
- [ ] User can delete/omit columns.
- [ ] XML export works.
- [ ] XML output is valid.
- [ ] `--onedir` build works.
- [ ] Complete PyInstaller output is included in the installer.
- [ ] Installer creates the required shortcuts.
- [ ] Installer uses the correct icon.
- [ ] Application works on another Windows machine.
- [ ] Uninstall works.
- [ ] Antivirus/endpoint-security results have been checked.
- [ ] Final installer has been tested.

---

# 26. Quick Reference

## Project

```text
C:\XMLConverter
```

## Run

```powershell
python main.py
```

## One-file

```powershell
python -m PyInstaller --onefile --windowed --icon=excelxml.ico --add-data "excelxml.ico;." --name ExcelXMLConverter main.py
```

## One-directory

```powershell
python -m PyInstaller --onedir --windowed --icon=excelxml.ico --add-data "excelxml.ico;." --name ExcelXMLConverter main.py
```

## Production recommendation

```text
PyInstaller --onedir
        ↓
Inno Setup
        ↓
ExcelXMLConverter_Setup.exe
```

---

# 27. Final Distribution File

The final file you give to another user should be:

```text
ExcelXMLConverter_Setup.exe
```

The user should be able to install and run ExcelXML Converter without manually installing Python or Python packages.
