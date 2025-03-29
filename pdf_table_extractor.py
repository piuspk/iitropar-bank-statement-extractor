import pdfplumber
import pandas as pd
import re
from datetime import datetime

def extract_bank_metadata(text):
    """Extract bank and account metadata from header"""
    metadata = {
        'Bank Name': '',
        'Branch Name': '',
        'Branch Address': '',
        'City': '',
        'PIN Code': '',
        'IFSC Code': '',
        'Account No': '',
        'Account Name': '',
        'Account Holder': '',
        'Statement Period': '',
        'Extracted On': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    
    if not text:
        return metadata
        
    lines = text.split('\n')
    for i, line in enumerate(lines):
        line = line.strip()
        if 'BANK NAME :' in line:
            metadata['Bank Name'] = line.split('BANK NAME :')[-1].strip()
        elif 'BRANCH NAME :' in line:
            metadata['Branch Name'] = line.split('BRANCH NAME :')[-1].strip()
        elif 'ADDRESS :' in line:
            metadata['Branch Address'] = line.split('ADDRESS :')[-1].strip()
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                if not next_line.startswith(('CITY :', 'IFSC')):
                    metadata['Branch Address'] += ' ' + next_line
        elif 'CITY :' in line:
            metadata['City'] = line.split('CITY :')[-1].strip()
        elif 'PIN CODE :' in line:
            metadata['PIN Code'] = line.split('PIN CODE :')[-1].strip()
        elif 'IFSC Code :' in line:
            metadata['IFSC Code'] = line.split('IFSC Code :')[-1].strip()
        elif 'Account No :' in line:
            metadata['Account No'] = line.split('Account No :')[-1].strip()
        elif 'A/C Name :' in line:
            metadata['Account Name'] = line.split('A/C Name :')[-1].strip()
        elif 'A/C Holder :' in line:
            metadata['Account Holder'] = line.split('A/C Holder :')[-1].strip()
        elif 'Statement of account for the period of' in line:
            metadata['Statement Period'] = line.split('period of')[-1].strip()
    
    return metadata

def extract_transactions_from_text(text):
    """Extract transactions from plain text for borderless tables"""
    transactions = []
    lines = text.split('\n')
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        
        if not line or any(x in line for x in ['BANK NAME', 'BRANCH NAME', 'ADDRESS', 
                                             'IFSC Code', 'MICR Code', 'REPORT PRINTED BY',
                                             'Account No', 'A/C Name', 'Statement of account',
                                             'Please examine', 'Manager', 'NOTE', '***END']):
            i += 1
            continue
            
        date_match = re.match(r'(\d{2}-[A-Za-z]{3}-\d{4})', line)
        if date_match:
            date = date_match.group(1)
            remaining = line[len(date):].strip()
            
            trans_type = ''
            if remaining.startswith(('T ', 'C ')):
                trans_type = remaining[0]
                remaining = remaining[2:].strip()
                
            amounts = re.findall(r'([\d,]+\.\d{2})(Dr|Cr)?', remaining.replace(',', ''))
            if len(amounts) >= 2:
                description = remaining[:remaining.find(amounts[0][0])].strip()
                amount = amounts[0][0]
                balance = amounts[1][0] + (amounts[1][1] if len(amounts[1]) > 1 else '')
                
                transactions.append({
                    'Date': date,
                    'Type': trans_type,
                    'Description': description,
                    'Amount': amount,
                    'Balance': balance,
                    'Cheque No': ''
                })
                i += 1
                continue
                
            description = remaining
            i += 1
            if i < len(lines):
                next_line = lines[i].strip()
                amounts = re.findall(r'([\d,]+\.\d{2})(Dr|Cr)?', next_line.replace(',', ''))
                if len(amounts) >= 2:
                    amount = amounts[0][0]
                    balance = amounts[1][0] + (amounts[1][1] if len(amounts[1]) > 1 else '')
                    transactions.append({
                        'Date': date,
                        'Type': trans_type,
                        'Description': description,
                        'Amount': amount,
                        'Balance': balance,
                        'Cheque No': ''
                    })
                elif len(amounts) == 1:
                    transactions.append({
                        'Date': date,
                        'Type': trans_type,
                        'Description': description,
                        'Amount': amounts[0][0],
                        'Balance': '',
                        'Cheque No': ''
                    })
            i += 1
        else:
            i += 1
            
    return transactions

def extract_bank_statement(pdf_path):
    """Extract complete bank statement with metadata, handling both bordered and borderless tables"""
    transactions = []
    metadata = {}
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text() or ""
            metadata = extract_bank_metadata(text)
            
            for page_num, page in enumerate(pdf.pages):
                try:
                    # Try text-based extraction first (for borderless tables)
                    text = page.extract_text()
                    if text:
                        page_transactions = extract_transactions_from_text(text)
                        transactions.extend(page_transactions)
                    
                    # If text extraction fails or yields few results, try table extraction (for bordered tables)
                    if not page_transactions or len(page_transactions) < 5:
                        tables = page.extract_tables()
                        for table in tables:
                            for row in table[1:]:  # Skip header row
                                if len(row) >= 5:  # Assuming minimum columns: Date, Type, Desc, Amount, Balance
                                    transactions.append({
                                        'Date': row[0] or '',
                                        'Type': row[1] or '',
                                        'Description': row[2] or '',
                                        'Amount': row[3] or '',
                                        'Balance': row[4] or '',
                                        'Cheque No': ''
                                    })
                except Exception as e:
                    print(f"Warning: Could not process page {page_num + 1}: {str(e)}")
                    continue
    
    except Exception as e:
        raise Exception(f"Failed to open PDF: {str(e)}")
    
    # Post-processing
    df = pd.DataFrame(transactions)
    
    if not df.empty:
        # Remove duplicates that might occur from dual extraction methods
        df = df.drop_duplicates(subset=['Date', 'Description', 'Amount', 'Balance'])
        
        # Extract cheque numbers from description
        df['Cheque No'] = df['Description'].str.extract(r'(\d{6,7})')[0]
        
        # Handle opening balance (B/F)
        if 'B/F' in df['Description'].values:
            bf_index = df[df['Description'] == 'B/F'].index[0]
            df.at[bf_index, 'Balance'] = df.at[bf_index, 'Amount']
            df.at[bf_index, 'Amount'] = ''
        
        # Reorder columns
        df = df[['Date', 'Type', 'Description', 'Cheque No', 'Amount', 'Balance']]
    
    return metadata, df

def save_to_excel(metadata, transactions, output_path):
    """Save extracted data to Excel with proper formatting"""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        pd.DataFrame([metadata]).to_excel(writer, sheet_name='Account Info', index=False)
        transactions.to_excel(writer, sheet_name='Transactions', index=False)
        
        workbook = writer.book
        info_sheet = writer.sheets['Account Info']
        for col in ['A', 'B']:
            info_sheet.column_dimensions[col].width = 25
        
        trans_sheet = writer.sheets['Transactions']
        column_widths = {'A': 12, 'B': 5, 'C': 50, 'D': 12, 'E': 15, 'F': 15}
        for col, width in column_widths.items():
            trans_sheet.column_dimensions[col].width = width

if __name__ == '__main__':
    pdf_file = 'normal_statement.pdf'
    print(f"Processing {pdf_file}...")
    
    try:
        metadata, transactions = extract_bank_statement(pdf_file)
        if not transactions.empty:
            output_file = 'complete_bank_statement.xlsx'
            save_to_excel(metadata, transactions, output_file)
            
            print(f"\nExtracted Account Information:")
            for key, value in metadata.items():
                print(f"{key:>20}: {value}")
            
            print(f"\nSuccess! Extracted {len(transactions)} transactions to {output_file}")
            print("\nFirst 5 transactions:")
            print(transactions.head(). comprehensionto_string(index=False))
        else:
            print("No transactions were extracted.")
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")