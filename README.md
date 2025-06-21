# Archival Automation for InDesign Projects

This Python-based automation tool is designed to streamline the archival and packaging process for InDesign-based publishing projects on macOS. It interacts with the filesystem, manages folder structures, and automates interactions with Extensis Connect and Adobe InDesign via AppleScript.

---

## 🧩 Features

- 📁 **Project Folder Detection & Setup**  
  Prompts the user to select a project folder (e.g., `11492_S24_Monroe_Color`) and tokenizes the folder name to extract metadata.

- 🗃️ **Automated Archival Workflow**  
  - Creates a project archive in `~/Documents/Archived_Projects/`
  - Copies over key subfolders: Digital_Content, Logs, Manuscript, and Office
  - Prepares layout and printer-ready folders

- 🖨️ **Print File Organization**  
  Moves and validates printer-ready PDF files with naming conventions like `11492_Print_*.pdf`.

- 🧠 **Extensis Connect & InDesign Automation**  
  - Checks if Extensis Connect is running and refreshes fonts
  - Opens InDesign files and automates the "Package" process into a standardized format
  - Skips missing font dialogs automatically

- 🔎 **Validation & QA Checks**  
  - Scans all archived files to ensure they are non-empty  
  - Alerts if expected print PDFs are missing

---

## 💻 Requirements

- macOS with:
  - **Python 3.9+**
  - **Extensis Connect** installed
  - **Adobe InDesign 2025**
- `tkinter` (comes with Python on macOS)
- AppleScript support (native on macOS)

---

## 🚀 How to Run

From Terminal:

```bash
cd ~/Documents/Executables
python3 run_cnt.py
