import PyPDF2
import pandas as pd
import re

def find_toc_page(pdf_path):
    """
    Finds the page number containing the Table of Contents in a PDF.
    
    Parameters:
        pdf_path (str): Path to the PDF file.
        
    Returns:
        int or None: Page number if found, else None.
    """
    with open(pdf_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        for page_num, page in enumerate(reader.pages, start=1):
            try:
                text = page.extract_text()
                if "Table of Contents" in text or "Contents" in text:
                    return page_num
            except Exception as e:
                print(f"Error reading page {page_num}: {e}")
    return None

# Example usage
pdf_path = r"C:\Users\ADITYA\Downloads\NIST.SP.800-53r5[1].pdf"  # Replace with your file path
toc_page = find_toc_page(pdf_path)
if toc_page:
    print(f"Table of Contents found on page {toc_page}.")
else:
    print("Table of Contents not found.")

import PyPDF2

def extract_toc_to_text(pdf_path, toc_page_num, text_file_path):
    """
    Extracts the Table of Contents from a specific page of a PDF and saves it to a text file.
    
    Parameters:
        pdf_path (str): Path to the PDF file.
        toc_page_num (int): Page number containing the Table of Contents.
        text_file_path (str): Path to save the text file.
    """
    with open(pdf_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        # Extract the text from the TOC page
        toc_page = reader.pages[toc_page_num - 1]
        toc_text = toc_page.extract_text()
        
    # Save the TOC text to a file
    with open(text_file_path, 'w', encoding='utf-8') as text_file:
        text_file.write(toc_text)
    
    print(f"Table of Contents saved to {text_file_path}")

# Example usage
toc_page_num = toc_page  # Replace with the page number you found
text_file_path = 'table_of_contents.txt'  # Desired text file output path
extract_toc_to_text(pdf_path, toc_page_num, text_file_path)

def process_toc_and_save_to_excel(txt_file_path, excel_file_path):
    """
    Processes a Table of Contents text file and saves it as a structured Excel sheet.
    
    Parameters:
        txt_file_path (str): Path to the input text file containing TOC data.
        excel_file_path (str): Path to save the structured Excel file.
    """
    # Read the text file
    with open(txt_file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    # Prepare to store structured TOC data
    toc_data = []

    # Regular expression to match TOC lines (adjust based on TOC format)
    toc_pattern = re.compile(r"^(?P<section>[0-9]+(\.[0-9]+)*)?\s*(?P<section_name>.+?)\s+(?P<page>\d+)$")
    
    for line in lines:
        # Strip leading and trailing whitespace
        line = line.strip()
        
        # Try to match the TOC pattern
        match = toc_pattern.match(line)
        if match:
            section = match.group('section') if match.group('section') else ''
            section_name = match.group('section_name')
            page = int(match.group('page'))
            
            # Append to data list
            toc_data.append({'Section': section, 'Section Name': section_name, 'Page Number': page})
    
    # Convert to DataFrame
    df = pd.DataFrame(toc_data)
    
    # Save to Excel
    df.to_excel(excel_file_path, index=False)
    print(f"Structured TOC saved to {excel_file_path}")

# Example usage
txt_file_path = 'table_of_contents.txt'  # Replace with the path to your TOC text file
excel_file_path = 'structured_toc.xlsx'  # Desired output Excel file path
process_toc_and_save_to_excel(txt_file_path, excel_file_path)

def find_offset(pdf_path):
    """
    Finds the page number where "Chapter One" appears as a heading in a PDF.
    
    Parameters:
        pdf_path (str): Path to the PDF file.
        
    Returns:
        int or None: The actual PDF page number where "Chapter One" appears as a heading, or None if not found.
    """
    heading_pattern = re.compile(r'^\s*chapter one\s*$', re.IGNORECASE)

    with open(pdf_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)

        for pdf_page_num, page in enumerate(reader.pages, start=1):
            text = page.extract_text()
            # Check if "Chapter One" appears as a heading
            lines = text.splitlines()
            for line in lines:
                if heading_pattern.match(line.strip()):
                    return pdf_page_num
    
    return None  # Return None if "Chapter One" is not found as a heading

# Example usage
chapter_one_heading_page = find_offset(pdf_path)

if chapter_one_heading_page is not None:
    print(f"'Chapter One' appears as a heading on actual PDF page {chapter_one_heading_page}.")
else:
    print("Unable to find 'Chapter One' as a heading in the PDF.")
    
excel_file = 'structured_toc.xlsx'

def get_page_number(file_path: str, section: str) -> int | None:
    """
    Returns the page number for a given section from a structured TOC Excel file.

    Parameters:
        file_path (str): Path to the Excel file.
        section (str): The section number to look up.

    Returns:
        int or None: Page number if found, otherwise None.
    """
    try:
        df = pd.read_excel(file_path)

        df.columns = df.columns.str.strip()  # Clean column names
        df['Section'] = df['Section'].astype(str).str.strip()
        result = df[df['Section'] == section]

        if not result.empty:
            return result.iloc[0]['Page Number']
        else:
            return None

    except Exception as e:
        print(f"Error: {e}")
        return None

def extract_text_by_toc_and_offset(pdf_path, toc_file, start_heading, end_heading, offset, output_file):
    """
    Extracts text from a PDF between two headings by parsing the TOC file to find their pages.
    
    Parameters:
        pdf_path (str): Path to the PDF file.
        toc_file (str): Path to the TOC text file.
        start_heading (str): The starting heading.
        end_heading (str): The ending heading.
        offset (int): Page offset to adjust TOC numbering to actual PDF page numbers.
        output_file (str): Path to save the extracted text.
    """
    # Parse TOC to get page numbers
    start_toc_page = get_page_number(excel_file, start_heading)
    end_toc_page = get_page_number(excel_file, end_heading)

    if start_toc_page is None or end_toc_page is None:
        print("Unable to find one or both headings in the TOC.")
        return

    # Adjust TOC page numbers using the offset
    start_pdf_page = start_toc_page + offset
    end_pdf_page = end_toc_page + offset

    extracted_text = []

    with open(pdf_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)

        for page_num in range(start_pdf_page - 1, end_pdf_page - 1):  # Convert to 0-based index
            page = reader.pages[page_num]
            extracted_text.append(page.extract_text())
    
    # Save extracted text to the output file
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write("\n".join(extracted_text))
    
    print(f"Text from '{start_heading}' to '{end_heading}' saved to {output_file}.")

# Example usage
toc_file = 'table_of_contents.txt'  # Replace with your TOC text file path
start_heading = '3.2'  # Starting heading
end_heading = '3.3'  # Ending heading
offset = find_offset(pdf_path) - 1  # Replace with the actual offset (difference between TOC and PDF numbering)
output_file = 'section_3_1_to_3_2.txt'  # Output text file
extract_text_by_toc_and_offset(pdf_path, toc_file, start_heading, end_heading, offset, output_file)






