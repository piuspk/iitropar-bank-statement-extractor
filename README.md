# Bank Statement Extractor

A simple tool to extract transaction data from bank statement PDFs and save as Excel. Works with multiple bank formats!

## What it Do? 

- Read your boring bank PDF statements 
- Pull out all the important transaction details 
- Save everything in nice Excel file 
- Handle different bank formats automatically 

## How to Use? 

1. Put your bank statement PDF in same folder as this script
2. Run the script (it's Python, so you need that installed)
3. Boom! You get `complete_bank_statement.xlsx` with all your transactions

```bash
python pdf_table_extractor.py

Supported Banks 
Standard Bank Statements (most Indian banks)

Special format for "Operating Account Consolidated Statement"

##  Why Use This?
Save time from manual copy-paste 

No need to pay for expensive software 

Simple enough for non-techies to use


##  Known Issues 
Some PDFs might be tricky to read (but I try my best!)

Very messy statements might miss few transactions

Dates formats must be proper or it get confused


Future Plans
Add more bank formats support

Make it work with scanned statements (OCR magic!)

Maybe make pretty graphs from your spending