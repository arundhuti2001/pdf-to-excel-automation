# pdf-to-excel-automation

# PDF to Excel Data Extraction Automation

## Overview

This project automates the extraction of financial and portfolio data from PDF factsheets and converts it into structured Excel reports. The solution extracts tables, parses textual information, organizes data into multiple worksheets, and generates an Excel workbook for further analysis and reporting.

## Features

- Automated PDF table extraction
- Text parsing from financial factsheets
- Portfolio data structuring
- Multi-sheet Excel report generation
- Entry and exit load extraction
- Performance table extraction
- IDCW data extraction
- Fund manager information reporting
- Portfolio holdings consolidation

## Technologies Used

- Python
- PDFPlumber
- Pandas
- OpenPyXL

## Workflow

1. Load PDF factsheet.
2. Extract tables from specified pages.
3. Parse textual information.
4. Extract performance metrics.
5. Extract IDCW details.
6. Extract fund load information.
7. Organize portfolio holdings.
8. Generate structured Excel workbook.

## Output Sheets

### PERFORMANCE
Performance metrics extracted from PDF tables.

### IDCW
Dividend and payout-related information.

### ENTRY_EXIT_LOAD
Entry and exit load details extracted from text.

### FUND_MANAGER
Fund manager information and experience details.

### PORTFOLIO
Portfolio holdings with industry classification and NAV allocation.

## Business Applications

- Mutual Fund Reporting
- Financial Data Extraction
- Investment Research
- Portfolio Analysis
- Regulatory Reporting
- Fund Factsheet Processing
- Data Digitization

## Installation

```bash
pip install pdfplumber pandas openpyxl
```

## Run

```bash
python main.py
```

## Output

The automation generates an Excel workbook containing structured data extracted from PDF factsheets, ready for reporting and analytics.
