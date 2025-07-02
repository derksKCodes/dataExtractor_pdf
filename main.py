import re
import requests
import pandas as pd
from io import BytesIO
import os

def install_packages():
    """Install required packages if not already installed"""
    try:
        import PyPDF2
        import pandas
        import openpyxl
    except ImportError:
        print("Installing required packages...")
        import subprocess
        subprocess.check_call(["pip", "install", "pypdf2", "pandas", "openpyxl", "requests"])
        print("Packages installed successfully.")

install_packages()

from PyPDF2 import PdfReader

def extract_school_data(text_block):
    """Extract structured data from a school text block"""
    data = {
        'School Name': 'NA',
        'Location': 'NA',
        'Address': 'NA',
        'City/ Town': 'NA',
        'County': 'NA',
        'Country': 'NA',
        'Website': 'NA',
        'Phone': 'NA',
        'Email': 'NA',
        'Fax': 'NA'
    }

    # Normalize text block
    text_block = re.sub(r'\s+', ' ', text_block).strip()

    # Extract school name (first line before any field)
    name_match = re.match(r'^(.*?)(?=Location:|Address:|City/ Town:|County:|Country:|Website:|Phone:|Email:|Fax:|$)', text_block)
    if name_match:
        data['School Name'] = name_match.group(1).strip()

    # Extract all fields
    fields = {
        'Location': r'Location:\s*(.*?)(?=Address:|City/ Town:|County:|Country:|Website:|Phone:|Email:|Fax:|$)',
        'Address': r'Address:\s*(.*?)(?=Location:|City/ Town:|County:|Country:|Website:|Phone:|Email:|Fax:|$)',
        'City/ Town': r'City/ Town:\s*(.*?)(?=Location:|Address:|County:|Country:|Website:|Phone:|Email:|Fax:|$)',
        'County': r'County:\s*(.*?)(?=Location:|Address:|City/ Town:|Country:|Website:|Phone:|Email:|Fax:|$)',
        'Country': r'Country:\s*(.*?)(?=Location:|Address:|City/ Town:|County:|Website:|Phone:|Email:|Fax:|$)',
        'Website': r'Website:\s*(.*?)(?=Location:|Address:|City/ Town:|County:|Country:|Phone:|Email:|Fax:|$)',
        'Phone': r'Phone:\s*(.*?)(?=Location:|Address:|City/ Town:|County:|Country:|Website:|Email:|Fax:|$)',
        'Email': r'Email:\s*(.*?)(?=Location:|Address:|City/ Town:|County:|Country:|Website:|Phone:|Fax:|$)',
        'Fax': r'Fax:\s*(.*?)(?=Location:|Address:|City/ Town:|County:|Country:|Website:|Phone:|Email:|$)'
    }

    for field, pattern in fields.items():
        match = re.search(pattern, text_block, re.IGNORECASE)
        if match:
            value = match.group(1).strip()
            if value:
                # Clean phone numbers
                if field == 'Phone':
                    phones = re.findall(r'(?:\b\d{3}[-.]?\d{3}[-.]?\d{3,4}\b|\b0\d{2,3}[- ]?\d{3,4}[- ]?\d{3,4}\b)', value)
                    cleaned_phones = []
                    for phone in phones:
                        cleaned = re.sub(r'[^\d]', '', phone)
                        if len(cleaned) == 9 and cleaned[0] in '23456789':
                            cleaned = '0' + cleaned
                        elif len(cleaned) == 12 and cleaned.startswith('254'):
                            cleaned = '0' + cleaned[3:]
                        cleaned_phones.append(cleaned)
                    data[field] = ', '.join(set(cleaned_phones)) if cleaned_phones else 'NA'
                # Clean emails
                elif field == 'Email':
                    emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', value, re.IGNORECASE)
                    data[field] = ', '.join(set(emails)) if emails else 'NA'
                # Clean websites
                elif field == 'Website':
                    websites = re.findall(r'(?:https?://|www\.)[^\s,]+', value)
                    data[field] = ', '.join(set(websites)) if websites else 'NA'
                else:
                    data[field] = value

    return data

def process_pdf(url):
    """Process PDF and extract school data"""
    try:
        print("Downloading PDF...")
        response = requests.get(url, timeout=30)
        response.raise_for_status()

        print("Extracting text...")
        pdf_file = BytesIO(response.content)
        reader = PdfReader(pdf_file)
        full_text = "\n".join(page.extract_text() or "" for page in reader.pages)

        # Pre-process text
        full_text = re.sub(r'(\w)-\s*\n\s*(\w)', r'\1\2', full_text)  # Fix hyphenated words
        full_text = re.sub(r'\s+', ' ', full_text)  # Normalize whitespace

        # Split into school blocks
        school_blocks = re.split(r'(?=\b[A-Z][a-zA-Z\s,&\.\'-]+(?:School|Academy|Sch\.|High|Centre|Foundation)\b)', full_text)[1:]

        schools = []
        for i in range(0, len(school_blocks), 2):
            if i+1 < len(school_blocks):
                block = school_blocks[i] + school_blocks[i+1]
            else:
                block = school_blocks[i]
            schools.append(extract_school_data(block))

        return schools

    except Exception as e:
        print(f"Error processing PDF: {e}")
        return []

def save_to_excel(data, filename):
    """Save extracted data to Excel file"""
    df = pd.DataFrame(data)
    columns = [
        'School Name', 'Location', 'Address', 'City/ Town', 'County', 'Country',
        'Website', 'Phone', 'Email', 'Fax'
    ]
    df = df.reindex(columns=columns, fill_value='NA')
    
    if os.path.exists(filename):
        print(f"Warning: {filename} already exists")
        response = input("Overwrite? (y/n): ").lower()
        if response != 'y':
            print("Operation cancelled")
            return False
    
    try:
        df.to_excel(filename, index=False, engine='openpyxl')
        print(f"Successfully saved {len(df)} records to {filename}")
        return True
    except Exception as e:
        print(f"Error saving to Excel: {e}")
        return False

if __name__ == "__main__":
    pdf_url = "https://www.devolutionhub.or.ke/file/f6ca1e61-nairobi-city-county-private-secondar.pdf"
    output_file = "nairobi_schools_data.xlsx"

    print(f"Processing PDF from: {pdf_url}")
    school_data = process_pdf(pdf_url)
    
    if school_data:
        save_to_excel(school_data, output_file)
    else:
        print("No data extracted. Please check the PDF format.")