# 📄 PDF Manager

<div align="center">


![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey.svg)

**A powerful and user-friendly GUI application for PDF manipulation**

</div>

---

## 🌟 Features

### ✂️ **Slice PDF**
Extract specific page ranges from PDF files with ease. Perfect for extracting chapters, sections, or individual pages.

- Select any PDF file
- Automatically detect total page count
- Choose custom page ranges (e.g., pages 5-10)
- Smart file browser remembers your last location

### 📄 **Merge PDFs**
Combine multiple PDF files into a single document seamlessly.

- Add unlimited PDF files
- Drag-and-drop multiple files at once
- Reorder files before merging
- Preview file list before processing

### 🔄 **PPTX to PDF**
Convert PowerPoint presentations to PDF format and merge multiple presentations.

- Support for .pptx files
- Batch conversion of multiple presentations
- Automatic merging into single PDF
- Cross-platform support (Windows, Linux, macOS)

### 💾 **Smart File Management**
- Intelligent default directories based on your workflow
- File browsers remember recent locations
- Custom output path selection
- Clean and intuitive interface

---

## 🚀 Installation

### Prerequisites

- **Python 3.7+** (Python 3.8+ recommended)
- **tkinter** (Usually included with Python)

### System-Specific Requirements

<details>
<summary><b>🐧 Linux (Ubuntu/Debian)</b></summary>

```bash
# Install tkinter
sudo apt-get update
sudo apt-get install python3-tk

# For PPTX to PDF conversion (optional)
sudo apt-get install libreoffice
```
</details>

<details>
<summary><b>🍎 macOS</b></summary>

```bash
# Tkinter is usually included with Python
# If needed, reinstall Python with Homebrew
brew install python-tk

# For PPTX to PDF conversion (optional)
brew install libreoffice
```
</details>

<details>
<summary><b>🪟 Windows</b></summary>

- Tkinter is included with Python from [python.org](https://www.python.org/)
- For PPTX conversion: Ensure Microsoft PowerPoint is installed
</details>

### Quick Install

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/pdf-manager.git
   cd pdf-manager
   ```

2. **Run the launcher script** (Recommended)
   ```bash
   chmod +x run.sh
   ./run.sh
   ```
   
   The launcher automatically:
   - Creates a virtual environment
   - Installs all dependencies
   - Launches the application

### Manual Installation

```bash
# Create virtual environment
python3 -m venv .venv

# Activate virtual environment
source .venv/bin/activate  # Linux/Mac
.venv\Scripts\activate     # Windows

# Install dependencies
pip install -r requirements.txt

# Run the application
python pdf_manager.py
```

---

## 📖 Usage

### Quick Start

Simply run the launcher:
```bash
./run.sh
```

Or run directly:
```bash
python pdf_manager.py
```

### Detailed Usage Guide

#### 1️⃣ Slicing PDFs

1. Open the **"Slice PDF"** tab
2. Click **Browse** to select your PDF file
3. The app automatically detects total pages
4. Enter your desired page range:
   - **From:** Starting page (e.g., 1)
   - **To:** Ending page (e.g., 10)
5. Click **Browse** next to Output Location
6. Name your output file
7. Click **✂️ SLICE PDF**

**Example Use Cases:**
- Extract a single chapter from a textbook
- Save specific pages from a large document
- Create sample PDFs from full versions

#### 2️⃣ Merging PDFs

1. Open the **"Merge PDFs"** tab
2. Click **Add Files** to select multiple PDFs
   - You can select multiple files at once
   - Files appear in merge order
3. Use **Remove Selected** to remove unwanted files
4. Use **Clear All** to start over
5. Click **Browse** to choose output location
6. Click **📄 MERGE PDFs**

**Example Use Cases:**
- Combine scanned documents
- Merge reports from different sources
- Create complete documentation from chapters

#### 3️⃣ Converting PPTX to PDF

1. Open the **"PPTX to PDF"** tab
2. Click **Add Files** to select PowerPoint files
3. Multiple files will be converted and merged
4. Use **Remove Selected** or **Clear All** as needed
5. Click **Browse** to choose output location
6. Click **🔄 CONVERT TO PDF**

**Note:** 
- Windows: Requires Microsoft PowerPoint
- Linux/Mac: Requires LibreOffice

---

## 🛠️ Technologies Used

- **Python 3.7+** - Core programming language
- **Tkinter** - GUI framework
- **PyPDF2** - PDF manipulation
- **python-pptx** - PowerPoint file handling
- **Pillow** - Image processing
- **LibreOffice** (Linux/Mac) - PPTX conversion

---

## 📋 Requirements

### Python Packages
```
PyPDF2==3.0.1
python-pptx==0.6.23
Pillow==10.1.0
comtypes==1.4.1  # Windows only
```

See `requirements.txt` for complete list.

---

## 🐛 Troubleshooting

### Common Issues

<details>
<summary><b>❌ "ModuleNotFoundError: No module named 'tkinter'"</b></summary>

**Solution:**
```bash
# Linux
sudo apt-get install python3-tk

# Mac
brew install python-tk
```
</details>

<details>
<summary><b>❌ PPTX conversion fails on Linux/Mac</b></summary>

**Solution:**
Install LibreOffice:
```bash
# Ubuntu/Debian
sudo apt-get install libreoffice

# Mac
brew install libreoffice
```

Test installation:
```bash
libreoffice --version
```
</details>

<details>
<summary><b>❌ "Permission denied" when saving files</b></summary>

**Solution:**
- Check write permissions for the output directory
- Try saving to your home directory or Desktop
- On Linux: `chmod +w /path/to/directory`
</details>

<details>
<summary><b>❌ "Invalid page range" error</b></summary>

**Solution:**
- Ensure "From" page ≤ "To" page
- Check that page numbers are within document range
- Page numbers must be positive integers
</details>

---

## 🤝 Contributing

### How to Continue This Project:

1. **Fork** the repository to your own GitHub account
2. **Create** a feature branch (`git checkout -b feature/AmazingFeature`)
3. **Commit** your changes (`git commit -m 'Add some AmazingFeature'`)
4. **Push** to your fork (`git push origin feature/AmazingFeature`)
5. Share your improvements with the community

### Ideas for Future Development
- Add drag-and-drop file support
- Implement PDF rotation
- Add PDF compression
- Create dark mode theme
- Add command-line interface
- Improve error handling
- Add unit tests

**If you create an actively maintained fork, please open an issue to have it listed here.**

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

```
MIT License

Copyright (c) 2025 PDF Manager

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 👨‍💻 Author

**Original Author:** This project was created as a standalone utility tool.

---


<div align="center">

**⭐ Star this repository if you find it helpful!**

Made with ❤️ and Python

[Back to Top](#-pdf-manager)

</div>
