import os
import re
import json
from pdf2image import convert_from_path
from pytesseract import image_to_string, image_to_data, Output


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


# Step 3: Detect first-level subheadings based on visual appearance
def detect_first_level_headings(image_paths):
    first_heading_style = None
    first_regex_heading = None
    headings = []

    for page_num, image_path in enumerate(image_paths, start=1):
        data = image_to_data(image_path, output_type=Output.DICT)

        lines = {}
        for i, text in enumerate(data['text']):
            if not text.strip():
                continue
            line_num = data['line_num'][i]
            if line_num not in lines:
                lines[line_num] = []
            lines[line_num].append({
                "text": text,
                "left": data['left'][i],
                "top": data['top'][i],
                "width": data['width'][i],
                "height": data['height'][i],
                "conf": int(data['conf'][i])
            })

        for line_num, words in lines.items():
            line_text = " ".join([w['text'] for w in words])
            avg_height = sum(w['height'] for w in words) / len(words)
            avg_top = sum(w['top'] for w in words) / len(words)

            # First detect heading by regex
            if first_regex_heading is None:
                if re.match(r"^[A-Z]{2}-\d+.*$", line_text):
                    first_regex_heading = line_text
                    # Debug: Print the first heading detected by regex
                    print(f"\nDEBUG: First heading detected by regex on page {page_num}, line {line_num}: '{line_text}'")

            # Fallback: Use the visually largest text block as the first heading
            if first_heading_style is None and first_regex_heading is None:
                first_heading_style = {
                    "avg_height": avg_height,
                    "avg_top": avg_top,
                    "text": line_text
                }
                print(f"\nDEBUG: Fallback - First visually detected heading: '{line_text}' on page {page_num}")

            # Assign the style for the first visually identified heading
            if first_heading_style is None and first_regex_heading:
                first_heading_style = {
                    "avg_height": avg_height,
                    "avg_top": avg_top,
                    "text": first_regex_heading
                }
                print(f"DEBUG: Visual heading style set based on regex match: '{first_regex_heading}'")

            # Compare visual appearance to first heading style
            if first_heading_style:
                height_diff = abs(avg_height - first_heading_style["avg_height"])
                height_threshold = first_heading_style["avg_height"] * 0.15  # 15% tolerance

                if height_diff <= height_threshold:
                    headings.append({"page": page_num, "line_num": line_num, "heading": line_text})

    print(f"\nDetected {len(headings)} headings based on visual similarity to the first heading.\n")
    for h in headings:
        print(f"Page {h['page']} Line {h['line_num']}: {h['heading']}")

    return headings


# Step 4: Save headings as JSON
def save_headings_json(headings, out_path):
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(headings, f, indent=2)
    print(f"\nHeadings saved to {out_path}")


# Main Script
if __name__ == "__main__":
    pdf_path = r"C:\Users\ADITYA\Downloads\NIST.SP.800-53r5_testing_2[1].pdf"  # Change this to your PDF file
    output_folder = "output_images"
    ocr_output_file = "ocr_output.txt"
    headings_json_file = "detected_headings.json"

    # Step 1: Convert PDF to images
    image_paths = pdf_to_images(pdf_path, output_folder=output_folder)

    # Step 2: Perform OCR (optional, for full text output)
    ocr_text = perform_ocr(image_paths, output_file=ocr_output_file)

    # Step 3: Detect first-level subheadings based on visual style
    headings = detect_first_level_headings(image_paths)

    # Step 4: Save detected headings to JSON
    save_headings_json(headings, headings_json_file)
