# 📄 PDF Table Extractor

A lightweight Python tool to extract tabular data from system-generated PDFs and save them into an Excel file — **without using Tabula or Camelot**.

**🔗 GitHub Repository**: [https://github.com/codersbliss/scoreme_hackathon](https://github.com/codersbliss/scoreme_hackathon)

**🔗  PRESENTATION FILE**: [https://docs.google.com/presentation/d/1WA5K4Aa7p0yFbrAmEU7MEM-PU58eOPFD/edit?usp=sharing&ouid=108849095306524663572&rtpof=true&sd=true](https://docs.google.com/presentation/d/1WA5K4Aa7p0yFbrAmEU7MEM-PU58eOPFD/edit?usp=sharing&ouid=108849095306524663572&rtpof=true&sd=true)

**🔗 OUTPUT EXCEL FILE**: [https://drive.google.com/drive/folders/1y3hN5p28K6C9jaizybiKdsTgrQnt9vxu?usp=sharing](https://drive.google.com/drive/folders/1y3hN5p28K6C9jaizybiKdsTgrQnt9vxu?usp=sharing)



## 🚀 Features

- Extracts tables from **any system-generated PDF** (not scanned)
- Saves each page's table in **separate Excel sheets**
- **No external Java dependencies** (unlike Tabula)
- Clean and readable output using Pandas

## 🛠️ Tech Stack

- Python 3
- [PyMuPDF (fitz)](https://pymupdf.readthedocs.io/) – For parsing PDFs
- Pandas – Data structure and table handling
- OpenPyXL – Excel file creation

## 🧾 How It Works

1. Scans the PDF and groups text lines by Y-coordinate (rows)
2. Sorts text within each row by X-coordinate
3. Builds tables dynamically for each page
4. Outputs everything to Excel

## 📦 Installation

```bash
pip install pymupdf pandas openpyxl
```

## 📂 Usage

```bash
python pdftoexcel.py test3.pdf --output tables.xlsx
```

**Arguments:**
- `input.pdf` – Path to your PDF file
- `--output` – (Optional) Name of output Excel file (default: `extracted_tables.xlsx`)

### 🔍 How to Give the Correct File Path

Make sure to provide the **full absolute path** or correct **relative path** to your PDF. Here are examples:

**Same folder as script:**
```bash
python pdftoexcel.py test3.pdf
```

**From Downloads folder (Windows):**
```bash
python pdftoexcel.py "C:/Users/YourUsername/Downloads/test3.pdf"
```

**From a different folder:**
```bash
python pdftoexcel.py "D:/Documents/Projects/test3.pdf"
```

⚠️ Always wrap the path in quotes if it contains spaces.

## ✅ Example

```bash
python pdftoexcel.py "test3.pdf" --output tables_final.xlsx
```

## 📄 Output

- One Excel file with multiple sheets: Each sheet corresponds to one page in the PDF.


