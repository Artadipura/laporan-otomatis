
# OJK Daily Report Automation

A Python-based desktop application developed during an internship at the Financial Services Authority (OJK) to automate the consolidation of daily Excel reports into an existing master workbook.

The application processes multiple Excel reports, including reports packaged in ZIP archives, maps the required data to the appropriate sheets, checks previously processed records, and generates an updated master workbook.

## Overview

The reporting process originally involved manually opening and consolidating multiple daily Excel reports into an existing master template.

This project automates that workflow by:

- Processing multiple Excel reports in one run
- Supporting `.xlsx`, `.xls`, and ZIP archives
- Reading required sheets and general information
- Mapping source columns to the existing master workbook
- Checking previously processed records to help prevent duplicates
- Handling different sheet structures and data formats
- Generating a timestamped output workbook
- Providing a desktop GUI for selecting input, template, and output folders

## Key Features

### Batch Excel Processing

The application can process multiple Excel reports from a selected input folder.

### ZIP Archive Processing

ZIP archives containing multiple Excel reports can be extracted and processed automatically.

### Existing Master Template

The application does not create the workbook structure from scratch.

It uses the existing master workbook template, including its predefined sheets, headers, and structure, and inserts the processed data into the appropriate locations.

### Duplicate Prevention

Previously processed records are checked using sheet-specific identity rules and normalized reporting periods to help prevent duplicate entries.

### Legacy Excel Support

The application supports both `.xlsx` and `.xls` files. Legacy `.xls` files can be converted to `.xlsx` during processing when required.

### Desktop Interface

A desktop GUI allows users to:

- Select the input folder
- Select the master Excel template
- Select the output folder
- Start the processing workflow
- Monitor processing logs and progress
- Open the output folder after processing

## Workflow

```text
Input Excel / ZIP Reports
          │
          ▼
Read Each Report
          │
          ▼
Extract Required Information
          │
          ▼
Map Source Data to Master Template
          │
          ▼
Validate Records
          │
     ┌────┴────┐
     │         │
   Valid     Invalid
     │         │
     ▼         ▼
Insert Data   Log / Skip
     │
     ▼
Check Existing Records
     │
     ▼
Generate Timestamped Workbook
````

## Machine Learning / Data Processing Approach

Although this project is primarily an automation application rather than a machine learning project, the implementation required structured data processing and validation.

The main processing stages are:

1. Identify the required workbook structure
2. Read sheet headers and data locations
3. Extract general information from each report
4. Normalize reporting periods and identifiers
5. Map source columns to the master template
6. Check existing records
7. Insert valid records into the appropriate sheets
8. Save the processed master workbook

## Technology Stack

* Python
* Tkinter
* Pandas
* OpenPyXL
* pathlib
* zipfile
* win32com.client
* PyInstaller

## Project Structure

```text
.
├── app.py
├── backend/
│   ├── __init__.py
│   └── report_processor.py
├── gui/
│   ├── __init__.py
│   ├── main_window.py
│   └── worker.py
├── assets/
├── .github/
│   └── workflows/
│       └── build-macos.yml
├── build_windows.bat
├── build_macos.command
├── LaporanHarianOtomatis.spec
├── requirements.txt
├── README.md
└── .gitignore
```

## Installation

Clone the repository:

```bash
git clone https://github.com/Artadipura/laporan-otomatis.git
cd laporan-otomatis
```

Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Run in Development Mode

Run the desktop application with:

```bash
python app.py
```

The application will open a desktop interface where the user can select:

1. Input folder containing Excel or ZIP reports
2. Existing master Excel template
3. Output folder

Then start the processing workflow from the application.

## Building the Application

### Windows

Run:

```bash
build_windows.bat
```

The packaged application will be generated in:

```text
dist/
```

### macOS

Run:

```bash
./build_macos.command
```

The packaged application will be generated in:

```text
dist/
```

The repository also includes a GitHub Actions workflow for the macOS build process.

## Output

The application generates a timestamped master workbook after processing.

The output workbook contains the consolidated data inserted into the corresponding sheets of the existing master template.

## Project Context

This project was developed during an internship at:

**Otoritas Jasa Keuangan (OJK) — Financial Services Authority**

Department of Financial Technology Innovation, Digital Financial Assets and Crypto Assets Supervision.

The application was designed around an existing reporting workflow and was tested against the actual process used during the internship.

The tool was used by relevant staff to make the consolidation process more consistent and substantially faster than the previous manual workflow.

## Learning Outcomes

Through this project, I learned how to:

* Analyze an existing business workflow before designing automation
* Work with complex Excel workbook structures
* Process multiple files in batches
* Handle different Excel formats
* Design sheet-specific data mapping
* Prevent duplicate records
* Implement validation and error handling
* Separate application logic from the graphical interface
* Package a Python application as a desktop application

## Important Note

The repository does **not** contain confidential OJK reports, company-specific data, production workbooks, or other sensitive information from the internship.

Only the application source code and supporting project files are included.

The application depends on the expected structure and sheet names of the master workbook used by the reporting workflow.

## License

This project is provided for portfolio and educational purposes.


