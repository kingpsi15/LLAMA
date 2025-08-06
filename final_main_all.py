import PyPDF2
import pandas as pd
import re
import os

def find_toc_page(pdf_path):
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

def extract_toc_to_text(pdf_path, toc_page_num, text_file_path):
    with open(pdf_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        toc_page = reader.pages[toc_page_num - 1]
        toc_text = toc_page.extract_text()

    with open(text_file_path, 'w', encoding='utf-8') as text_file:
        text_file.write(toc_text)
    
    print(f"Table of Contents saved to {text_file_path}")

def process_toc_and_save_to_excel(txt_file_path, excel_file_path):
    with open(txt_file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    toc_data = []
    toc_pattern = re.compile(r"^(?P<section>[0-9]+(\.[0-9]+)*)?\s*(?P<section_name>.+?)\s+(?P<page>\d+)$")
    
    for line in lines:
        line = line.strip()
        match = toc_pattern.match(line)
        if match:
            section = match.group('section') if match.group('section') else ''
            # Remove trailing dots and spaces from section name
            section_name = match.group('section_name').rstrip('. ').strip()
            page = int(match.group('page'))
            toc_data.append({'Section': section, 'Section Name': section_name, 'Page Number': page})

    df = pd.DataFrame(toc_data)
    df.to_excel(excel_file_path, index=False)
    print(f"Structured TOC saved to {excel_file_path}")

def find_offset(pdf_path):
    heading_pattern = re.compile(r'^\s*chapter one\s*$', re.IGNORECASE)
    with open(pdf_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        for pdf_page_num, page in enumerate(reader.pages, start=1):
            text = page.extract_text()
            lines = text.splitlines()
            for line in lines:
                if heading_pattern.match(line.strip()):
                    return pdf_page_num
    return None

def get_page_number(file_path: str, section: str) -> int | None:
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
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
    start_toc_page = get_page_number(excel_file, start_heading)
    end_toc_page = get_page_number(excel_file, end_heading)

    if start_toc_page is None or end_toc_page is None:
        print("Unable to find one or both headings in the TOC.")
        return

    start_pdf_page = start_toc_page + offset
    end_pdf_page = end_toc_page + offset

    extracted_text = []

    with open(pdf_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)

        for page_num in range(start_pdf_page - 1, end_pdf_page - 1):
            page = reader.pages[page_num]
            extracted_text.append(page.extract_text())
    
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write("\n".join(extracted_text))
    
    print(f"Text from '{start_heading}' to '{end_heading}' saved to {output_file}")

def extract_all_sections(pdf_path, toc_excel_file, offset, output_dir):
    """
    Extracts text for all sections listed in a structured TOC Excel file.
    Saves each section as: <section number> <section name>.txt inside subfolders named <section number> <section name>
    """
    os.makedirs(output_dir, exist_ok=True)

    df = pd.read_excel(toc_excel_file, dtype={'Section': str})
    df = df.sort_values('Page Number').reset_index(drop=True)

    with open(pdf_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        total_pages = len(reader.pages)

        for i in range(len(df) - 1):
            start_section = df.iloc[i]
            end_section = df.iloc[i + 1]

            start_page = int(start_section['Page Number']) + offset - 1
            end_page = int(end_section['Page Number']) + offset - 1

            extracted_text = []
            for page_num in range(start_page, end_page):
                if 0 <= page_num < total_pages:
                    text = reader.pages[page_num].extract_text()
                    extracted_text.append(text)

            section_id = str(start_section['Section']).strip()
            section_name = str(start_section['Section Name']).strip().replace('/', '-')
            section_folder = os.path.join(output_dir, f"{section_id} {section_name}")
            os.makedirs(section_folder, exist_ok=True)

            filename = os.path.join(section_folder, f"{section_id} {section_name}.txt")

            with open(filename, 'w', encoding='utf-8') as file:
                file.write("\n".join(extracted_text))
            print(f"Saved: {filename}")

        # Handle the last section
        last_section = df.iloc[-1]
        start_page = int(last_section['Page Number']) + offset - 1
        extracted_text = []
        for page_num in range(start_page, total_pages):
            text = reader.pages[page_num].extract_text()
            extracted_text.append(text)

        section_id = str(last_section['Section']).strip()
        section_name = str(last_section['Section Name']).strip().replace('/', '-')
        section_folder = os.path.join(output_dir, f"{section_id} {section_name}")
        os.makedirs(section_folder, exist_ok=True)

        filename = os.path.join(section_folder, f"{section_id} {section_name}.txt")

        with open(filename, 'w', encoding='utf-8') as file:
            file.write("\n".join(extracted_text))
        print(f"Saved: {filename}")


# ==== MAIN EXECUTION ====

pdf_path = r"C:\Users\ADITYA\Downloads\NIST.SP.800-53r5[1].pdf"
toc_file = 'table_of_contents.txt'
excel_file = 'structured_toc.xlsx'
output_dir = 'extracted_sections'

# Step 1: Find TOC page
toc_page = find_toc_page(pdf_path)
if toc_page:
    print(f"TOC found on page {toc_page}.")
    extract_toc_to_text(pdf_path, toc_page, toc_file)
    process_toc_and_save_to_excel(toc_file, excel_file)
else:
    print("TOC not found.")
    exit()

# Step 2: Find offset using "Chapter One"
offset = find_offset(pdf_path)
if offset:
    offset -= 1
    print(f"Offset calculated: {offset}")
else:
    print("Unable to determine offset.")
    exit()

# Step 3: Extract all sections
extract_all_sections(pdf_path, excel_file, offset, output_dir)
