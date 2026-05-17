# VCM CLI App

VCM (VBA Code Manager) is a Python CLI tool for formatting, importing, and exporting VBA code. It is designed for fast command-line automation to work with your VBA Project outside VBEditor.

---

## 🚀 Features

- Format VBA code with indentation control
- Import VBA files or projects
- Export VBA code to files or bundles
- Single-file or batch operations
- Force overwrite support
- Clean modular CLI architecture

---

## ⚙️ Installation & Setup

### 1. Clone the repository

```bash
git clone
cd VCM
```

---

### 2. Create virtual environment (recommended)

```bash
python -m venv .venv
```

Activate it:

**Windows (PowerShell):**
```bash
.venv\Scripts\activate
```
---

### 3. Install dependencies

Using pip:

```bash
pip install -r requirements.txt
```

OR if using uv:

```bash
uv pip install -r requirements.txt
```

---

## ▶️ Running the CLI

```bash
python main.py
```

---

## 📌 Commands Reference

---

### 🔹 Help

```bash
vcm -h
vcm --help
```

📖 **Description:**
Displays all available commands and usage information for VCM CLI.

---

### 🔹 Version

```bash
vcm -v
vcm --version
```

📖 **Description:**
Shows the current installed version of VCM.

---

## 📤 Export Commands

Export VBA code from the system.

### Basic Export

```bash
vcm export
```

📖 **Description:**
Exports all available VBA code using default settings.
Note: This command will not overwrite code if it already exists.

---

### Force Export

```bash
vcm export -f
vcm export --force
```

📖 **Description:**
Forces export and overwrites existing output files if they already exist.

---

### Export Single File (Required)

```bash
vcm export -o filename
vcm export --onefile filename
```

📖 **Description:**
Exports a specific VBA file by name.
⚠️ The filename is required for this operation.

---

## 📥 Import Commands

Import VBA code into the VBA Project.

### Basic Import

```bash
vcm import
```

📖 **Description:**
Imports all VBA files into the project.

---

### Force Import

```bash
vcm import -f
vcm import --force
```

📖 **Description:**
Forces import - Normal Import + Deleting VBA Components that are not in src folder.

---

### Import Single File (Required)

```bash
vcm import -o filename
vcm import --onefile filename
```

📖 **Description:**
Imports a specific VBA file by name.
⚠️ Filename is required.

---

## 🎨 Format Commands

Format VBA code with indentation control.

### Format All Files

```bash
vcm format -a
vcm format --all
```

📖 **Description:**
Formats all VBA files in the project with default indentation rules.

---

### Format Single File (Required)

```bash
vcm format -o filename
vcm format --onefile filename
```

📖 **Description:**
Formats a specific VBA file.
⚠️ Filename is required.

---

### Format with Indentation

```bash
vcm format -o filename -i 4
vcm format --onefile filename --indent 4
```

📖 **Description:**
Formats a file and applies custom indentation level.
Default indentation is 2

---

## 🏗️ Build Executable (PyInstaller)

After installing dependencies, you can build a standalone executable.

### Build command:

```bash
python -m PyInstaller --onefile --name vcm --icon=VCM.ico --paths=. main.py
```

📖 **Description:**
Creates a standalone executable (`vcm.exe`) inside the `dist/` folder.
This allows users to run VCM without needing Python installed.

---

### Output:

```bash
dist/vcm.exe
```

---

## 📁 Project Structure

```
VBA_Code_Manager/
│── main.py              # CLI entry point
│── commands/           # CLI command handlers
│── utils/              # Helper utilities
│── VCM.ico             # Application icon
```

---

## 🧠 Notes

- `-o / --onefile` → required for single-file operations
- `-f / --force` → overwrites existing data
- `-a / --all` → processes all files
- `-i / --indent` → controls formatting indentation

---

## 📄 License
Raven Mark S. Eduardo Software License (Proprietary)

---

## 👤 Author
Raven Eduardo