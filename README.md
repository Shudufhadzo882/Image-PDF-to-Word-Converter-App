# AeroConvert - Premium Image & PDF to Word Converter

A sleek, glassmorphic dark-mode web application built with Python (Flask) that converts images and PDF files into fully editable Microsoft Word (`.docx`) documents.

## Features

- **High-Fidelity UI**: Frosted glassmorphism elements, dynamic radial glowing hover effects, custom typography, and responsive animations.
- **Single File Conversion**: Convert a single Image or PDF directly and stream the `.docx` download instantly.
- **Batch ZIP Conversion**: Drag and drop multiple files to convert them simultaneously, producing a consolidated `.zip` download.
- **Memory-Safe File Streaming**: Temp files are loaded into memory and deleted immediately on disk to avoid file locking on Windows.
- **No-Poppler PDF to Image Fallback**: Automatically renders PDF pages to images using PyMuPDF (fitz) out-of-the-box, eliminating complex Windows Poppler binary installations.
- **OCR Text Extraction**: Converts scanned PDFs and images to editable text via Tesseract-OCR, while embedding the original pages/images for references.

---

## Installation & Setup

### 1. Prerequisites
- **Python 3.8+**
- **Tesseract-OCR** (Optional, required to enable the OCR text extraction feature):
  - Download and run the Windows installer from [UB-Mannheim Tesseract Wiki](https://github.com/UB-Mannheim/tesseract/wiki).
  - Install using default options (it will install to `C:\Program Files\Tesseract-OCR\`).
  - The app will automatically detect it and unlock the OCR capability!

### 2. Clone the Repository
```bash
git clone https://github.com/Shudufhadzo882/Image-PDF-to-Word-Converter-App.git
cd Image-PDF-to-Word-Converter-App
```

### 3. Set Up Virtual Environment
```powershell
python -m venv .venv
.venv\Scripts\activate
```

### 4. Install Dependencies
```powershell
pip install -r converter_app/requirements.txt
```

### 5. Run the Server
```powershell
python converter_app/app.py
```
Open your browser and navigate to `http://127.0.0.1:5000`.

---

## License

This project is licensed under the MIT License.
