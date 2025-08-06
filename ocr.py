import os
import re
import json
from pdf2image import convert_from_path
from pytesseract import image_to_string

# Step 1: Convert PDF to images
def pdf_to_images(pdf_path, output_folder="output_images", dpi=300):
    print("Converting PDF to images...")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    images = convert_from_path(pdf_path, dpi=dpi)
    image_paths = []
    for i, image in enumerate(images, start=1):
        image_path = os.path.join(output_folder, f"page_{i}.png")
        image.save(image_path, "PNG")
        image_paths.append(image_path)
    print(f"PDF converted to {len(images)} images and saved to '{output_folder}'")
    return image_paths

# Step 2: Perform OCR on images
def perform_ocr(image_paths, output_file="ocr_output.txt"):
    print("Performing OCR on images...")
    ocr_text = ""
    for image_path in image_paths:
        print(f"Processing image: {os.path.basename(image_path)}...")
        text = image_to_string(image_path, lang="eng")
        ocr_text += text + "\n\n"
    
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(ocr_text)
    print(f"OCR completed. Text saved to '{output_file}'")
    return ocr_text

# Step 3: Detect first-level subheadings
def detect_first_level_headings(ocr_text_path):
    with open(ocr_text_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    headings = []
    # Regex to detect first-level subheadings:
    #  - AC-1, AC-2, etc. (two uppercase letters, dash, number(s))
    #  - 3.1 ... style headings (section number + uppercase heading)
    pattern = re.compile(r"^([A-Z]{2}-\d+.*|3\.\d+.*)$")

    for idx, line in enumerate(lines):
        line = line.strip()
        if pattern.match(line):
            headings.append({"line": idx + 1, "heading": line})

    # Print detected headings
    print(f"Detected {len(headings)} first-level subheadings:\n")
    for h in headings:
        print(f"Line {h['line']}: {h['heading']}")

    return headings

# Step 4: Save headings as JSON
def save_headings_json(headings, out_path):
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(headings, f, indent=2)
    print(f"\nHeadings saved to {out_path}")

# Main Script
if __name__ == "__main__":
    # Input PDF file
    pdf_path = r"C:\Users\ADITYA\Downloads\NIST.SP.800-53r5_testing_2[1].pdf"  # Change this to your PDF file
    output_folder = "output_images"
    ocr_output_file = "ocr_output.txt"
    headings_json_file = "detected_headings.json"

    # Step 1: Convert PDF to images
    image_paths = pdf_to_images(pdf_path, output_folder=output_folder)

    # Step 2: Perform OCR
    ocr_text = perform_ocr(image_paths, output_file=ocr_output_file)

    # Step 3: Detect first-level subheadings
    headings = detect_first_level_headings(ocr_output_file)

    # Step 4: Save detected headings to JSON
    save_headings_json(headings, headings_json_file)
