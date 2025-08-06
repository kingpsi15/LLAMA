import re
import json
from PyPDF2 import PdfReader

# Step 1: Extract text from the PDF
def extract_text_from_pdf(pdf_path):
    print("Extracting text from PDF...")
    reader = PdfReader(pdf_path)
    pdf_text = ""
    for i, page in enumerate(reader.pages):
        print(f"Processing page {i+1}...")
        pdf_text += page.extract_text() + "\n\n"
    print("Text extraction complete.")
    return pdf_text

# Step 2: Detect first-level subheadings
def detect_first_level_headings(pdf_text):
    # Regex to detect first-level subheadings:
    # - AC-1, AC-2, etc. (two uppercase letters, dash, number(s))
    # - 3.1 ... style headings (section number + uppercase heading)
    pattern = re.compile(r"^([A-Z]{2}-\d+.*|3\.\d+.*)$", re.MULTILINE)

    headings = []
    for match in pattern.finditer(pdf_text):
        heading = match.group(0).strip()
        start_idx = match.start()
        headings.append({"start_index": start_idx, "heading": heading})

    print(f"Detected {len(headings)} first-level subheadings:\n")
    for h in headings:
        print(f"Start Index {h['start_index']}: {h['heading']}")

    return headings

# Step 3: Save headings as JSON
def save_headings_json(headings, out_path):
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(headings, f, indent=2)
    print(f"\nHeadings saved to {out_path}")

# Main Script
if __name__ == "__main__":
    # Input PDF file
    pdf_path = r"C:\Users\ADITYA\Downloads\NIST.SP.800-53r5_testing_2[1].pdf"  # Change this to your PDF file
    headings_json_file = "detected_headings.json"

    # Step 1: Extract text from PDF
    pdf_text = extract_text_from_pdf(pdf_path)

    # Step 2: Detect first-level subheadings
    headings = detect_first_level_headings(pdf_text)

    # Step 3: Save detected headings to JSON
    save_headings_json(headings, headings_json_file)
