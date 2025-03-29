# Bank Statement Extractor

A Python tool to extract transaction data from bank statement PDFs and export to Excel. Supports multiple bank formats.

## Features

- Extract transaction data from PDF bank statements  
- Parse key details: dates, amounts, descriptions, balances  
- Export clean Excel files with organized data  
- Supports multiple bank formats automatically  

## Supported Banks

- Standard Indian bank statements (most major banks)  
- Special "Operating Account Consolidated Statement" format  

## Installation

1. Ensure Python 3.6+ is installed  
2. Install required packages:  
   ```bash
   pip install pdfplumber pandas openpyxl

##   Usage
Place your bank statement PDF in the project folder

Run the extractor:

bash
Copy
python pdf_table_extractor.py
Find results in complete_bank_statement.xlsx

##  Known Limitations
Poor quality scans may not be read accurately

Non-standard formats may require adjustments

Requires proper date formats in source PDFs

Complex statements may miss some transactions

##  Future Roadmap
Add support for more bank formats

Implement OCR for scanned statements

Add spending visualization features

Improve error handling for edge cases
