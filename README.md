# Automated Resume & Cover Letter Data Pipeline

An automated data extraction and processing pipeline developed in Python to parse unstructured candidate cover letters from Microsoft Word documents (`.docx`), extract key applicant details using Regular Expressions (`re`), and aggregate the structured outputs into a standardized Excel spreadsheet (`.xlsx`).

---

## 🏗️ Pipeline Architecture & Workflow

```text
+-------------------------------------------------------------+
|                     Input Directory                         |
|                    (cover_letters/)                         |
|   [le_van_c.docx]  [pham_thi_d.docx]  [tran_van_k.docx] ... |
+------------------------------+------------------------------+
                               |
                               v
+-------------------------------------------------------------+
|               Orchestration Layer (main.py)                 |
|  - Initialize CoverLetterProcessor instance                 |
|  - Target folder path & output Excel specifications         |
+------------------------------+------------------------------+
                               |
                               v
+-------------------------------------------------------------+
|             Core Processor Module (python-docx)             |
|  - Iterate through all .docx files in folder                |
|  - Ingest document body and extract raw text paragraphs     |
+------------------------------+------------------------------+
                               |
                               v
+-------------------------------------------------------------+
|             Data Extraction & Regex Engine (re)             |
|  - Pattern Matching: Name, Email, Phone, Role, Experience   |
|  - Text Normalization & String Cleaning                     |
+------------------------------+------------------------------+
                               |
                               v
+-------------------------------------------------------------+
|               Export & Storage Layer (openpyxl)             |
|  - Initialize tabular schema (Headers & Formatting)         |
|  - Append extracted candidate records row-by-row            |
|  - Persist to disk: so_yeu_ly_lich.xlsx                     |
+-------------------------------------------------------------+

```

---

## 📌 Key Features

* **Batch Document Ingestion**: Scans and parses multiple candidate Word documents concurrently from the designated input directory (`cover_letters/`).
* **Entity & Text Extraction**:
* Leverages `python-docx` to extract raw paragraphs and textual content.
* Applies pattern matching via Regular Expressions (`re`) to parse applicant profile attributes (Full Name, Contact Email, Phone Number, Position Applied, Professional Background).


* **Structured Data Export**:
* Automatically initializes and formats destination workbooks using `openpyxl`.
* Appends cleaned, structured records into `so_yeu_ly_lich.xlsx` for streamlined human resources analysis and reporting.



---

## 📂 Project Structure

```text
Session_8/
│
├── cover_letters/             # Input directory containing candidate Word (.docx) files
├── cover_letters_processor.py # Core processor module (defines CoverLetterProcessor class)
├── main.py                    # Pipeline execution and orchestration entry point
├── read_docx.py               # Exploratory script for document inspection & testing
├── so_yeu_ly_lich.xlsx        # Consolidated output spreadsheet with extracted data
├── .gitignore                 # Git ignore configuration (Python cache, OS metadata)
└── README.md                  # Comprehensive project documentation

```

---

## 🛠️ Requirements & Installation

Developed and tested on **Python 3.x**. Install the required dependencies:

```bash
pip install python-docx openpyxl

```

---

## 🚀 Execution Guide

### 1. Data Ingestion Setup

Place all target `.docx` cover letters into the `cover_letters/` directory.

### 2. Run the Processing Pipeline

Execute the workflow via the main entry script:

```bash
python main.py

```

Underlying execution logic (`main.py`):

```python
from cover_letters_processor import CoverLetterProcessor

# Initialize processor with target input folder and output spreadsheet
processor = CoverLetterProcessor(
    folder_path="cover_letters",
    excel_file="so_yeu_ly_lich.xlsx"
)

# Parse documents and export structured records
processor.process_documents()

```

### 3. Review Generated Data

Upon task completion, inspect `so_yeu_ly_lich.xlsx` to review the tabular dataset containing parsed candidate records.

---

## 🧰 Technology Stack

* **Python**: Primary programming language for pipeline implementation.
* **python-docx**: DOCX file parser for programmatic text extraction.
* **re (Regular Expressions)**: Pattern matching for unstructured text normalization.
* **openpyxl**: Automated Excel workbook generation, cell formatting, and tabular output export.

```

---

### Các bước cập nhật lên GitHub qua Terminal:
1. Mở file `README.md` trong VS Code, dán đè toàn bộ nội dung trên rồi nhấn `Ctrl + S`.
2. Chạy lệnh sau ở Terminal để cập nhật:
   ```bash
   git add README.md
   git commit -m "Docs: Add ASCII architecture workflow diagram to README"
   git push

```
