# CoverLetterProcessor: Word-to-Excel Profile Data Extraction

A small Python tool that reads a folder of Word (`.docx`) personal-profile forms (*Sơ yếu lý lịch tự thuật*), extracts key fields with regular expressions, and saves them as structured rows in an Excel file.

> Course project. All sample data in this repository is for practice.

## Overview

Personal-profile forms arrive as unstructured Word documents. Reading and re-typing each one into a spreadsheet is slow and easy to get wrong. This project turns that task into a repeatable pipeline:

**Word files → text → extracted fields → Excel rows → saved workbook**

The logic is wrapped in a single class, `CoverLetterProcessor`, so it can be reused with any folder and any output file.

## Features

- Reads every `.docx` file in a folder (other file types are skipped).
- Extracts **8 fields** from each document:
  `Họ và tên`, `Giới tính`, `Ngày sinh`, `Nơi sinh`, `Nguyên quán`, `Hộ khẩu thường trú`, `Chỗ ở hiện nay`, `Điện thoại`
- Normalizes extra whitespace in every extracted value.
- Leaves a field blank when it is not found, so every row has the same columns.
- Opens the output workbook if it exists, or creates a new one with a header row.
- Appends one row per document and saves the result to Excel.

## Workflow

| Step | Action | Method |
|---|---|---|
| 1 | Open or create the Excel workbook | `initialize_excel()` |
| 2 | List the files in the input folder | `os.listdir()` |
| 3 | Loop through each `.docx` file | `process_documents()` |
| 4 | Read the document text | `read_docx()` |
| 5 | Extract the 8 fields | `extract_info()` |
| 6 | Append the values as a new row | `ws.append()` |
| 7 | Save the workbook | `wb.save()` |

## Tech Stack

- **Python 3**
- [`python-docx`](https://python-docx.readthedocs.io/) for reading Word files
- `re` (regular expressions) for field extraction
- [`openpyxl`](https://openpyxl.readthedocs.io/) for reading and writing Excel
- `os` for file and folder handling

## Project Structure

```
buoi_08/
├── cover_letter_processor.py   # CoverLetterProcessor class (main logic)
├── main.py                     # Entry point: sets the folder and Excel file, runs the pipeline
├── read_docx.py                # Short demo of how python-docx reads a Word file
├── cover_letters/              # Input: 12 sample .docx profile forms
└── so_yeu_ly_lich.xlsx         # Output: extracted data
```

## Getting Started

### 1. Install dependencies

```bash
pip install python-docx openpyxl
```

### 2. Run

From inside the `Session_08` folder:

```bash
python main.py
```

`main.py` uses these settings:

```python
from cover_letter_processor import CoverLetterProcessor

processor = CoverLetterProcessor(
    folder_path="cover_letters",
    excel_file="so_yeu_ly_lich.xlsx"
)
processor.process_documents()
```

On success, the program prints `Xử lý thành công` and writes the results to `so_yeu_ly_lich.xlsx`.

To use your own data, change `folder_path` to the folder holding your `.docx` files and `excel_file` to the output name you want.

## Input Format

Each Word file is expected to follow the same labeled layout, for example:

```
1. Họ và tên: LÊ VĂN C          Nam/Nữ: Nam
2. Sinh ngày: 20 tháng 08 năm 2013   Nơi sinh: Thành phố Đà Nẵng
3. Nguyên quán: Phường Hải Châu I, Quận Hải Châu, Thành phố Đà Nẵng
4. Nơi đăng ký hộ khẩu thường trú: Phường Hải Châu I, Quận Hải Châu, Thành phố Đà Nẵng
5. Chỗ ở hiện nay: Toà Beta Building, Quận Sơn Trà, Đà Nẵng
6. Điện thoại liên hệ: 9234567890
```

The tool only reads the 8 fields listed above. Other fields on the form are ignored.

## How the Extraction Works

`read_docx()` joins all paragraphs of a document into one text string. `extract_info()` then applies one regular expression per field (case-insensitive), takes the first capture group, collapses repeated whitespace, and returns a dictionary:

| Field | Pattern idea |
|---|---|
| Họ và tên | text after `Họ và tên:` up to `Nam/Nữ:` |
| Giới tính | text after `Nam/Nữ:` to end of line |
| Ngày sinh | text after `Sinh ngày:` up to `Nơi sinh:` |
| Nơi sinh | text after `Nơi sinh:` to end of line |
| Nguyên quán | text after `Nguyên quán:` to end of line |
| Hộ khẩu thường trú | text after `Nơi đăng ký hộ khẩu thường trú:` to end of line |
| Chỗ ở hiện nay | text after `Chỗ ở hiện nay:` to end of line |
| Điện thoại | text after `Điện thoại:` or `Điện thoại liên hệ:` to end of line |

If a pattern finds no match, the value is an empty string.

## Example Output

Dictionary returned for one document:

```python
{
    "Họ và tên": "LÊ VĂN C",
    "Giới tính": "Nam",
    "Ngày sinh": "20 tháng 08 năm 2013",
    "Nơi sinh": "Thành phố Đà Nẵng",
    "Nguyên quán": "Phường Hải Châu I, Quận Hải Châu, Thành phố Đà Nẵng",
    "Hộ khẩu thường trú": "Phường Hải Châu I, Quận Hải Châu, Thành phố Đà Nẵng",
    "Chỗ ở hiện nay": "Toà Beta Building, Quận Sơn Trà, Đà Nẵng",
    "Điện thoại": "9234567890"
}
```

Running the tool on the 12 sample files produces one Excel sheet named `Thông tin người dùng` with **12 rows and 8 columns**.

## Current Limitations

- Re-running the program on an existing Excel file **appends the same records again**, because the file is opened and new rows are added without checking for duplicates.
- The output has no column showing which source file each row came from.
- A field that is not found is left blank rather than flagged.
- One unreadable or malformed file is not handled separately, so it can stop the whole run.
- Excel column widths are not adjusted automatically.
- The regular expressions depend on the exact labels used in the forms. A different layout needs updated patterns.

## Possible Improvements

- [ ] Add a `Tên file` column to trace each record to its source document.
- [ ] Write `Không tìm thấy` for missing fields instead of leaving them blank.
- [ ] Wrap file processing in `try/except` so one bad file does not stop the run.
- [ ] Print the number of files processed successfully and the number that failed.
- [ ] Auto-fit Excel column widths.
- [ ] Skip records that already exist in the workbook.

## Author

TPA-Analyst
