# ========================
# === subheading_isolator.py ===
# ========================
import os
import pdfplumber
import pandas as pd
from pypdf import PdfReader
import re

def extract_text_with_styles(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text_data = []
        for page_num, page in enumerate(pdf.pages, start=1):
            for char in page.chars:
                text_data.append({
                    "text": char["text"],
                    "page": page_num,
                    "fontname": char.get("fontname", ""),
                    "fontsize": char.get("size", 0),
                    "x": char.get("x0", 0),
                    "y": char.get("top", 0),
                })
        return text_data

def group_text_by_position(text_data, line_tolerance=2):
    grouped_lines = []
    current_line = []

    for i, char_data in enumerate(sorted(text_data, key=lambda x: (x["page"], x["y"], x["x"]))):
        if not current_line:
            current_line.append(char_data)
        else:
            last_char = current_line[-1]
            same_line = (
                abs(char_data["y"] - last_char["y"]) <= line_tolerance and
                char_data["page"] == last_char["page"]
            )
            if same_line:
                current_line.append(char_data)
            else:
                grouped_lines.append(current_line)
                current_line = [char_data]

    if current_line:
        grouped_lines.append(current_line)

    return grouped_lines

def is_all_caps(text):
    filtered = ''.join(c for c in text if c.isalpha())
    return filtered.isupper() if filtered else False

def is_bold(fontname):
    return "Bold" in fontname or "bold" in fontname

def extract_all_subheadings_with_style(lines, fontname, fontsize, all_caps, bold):
    matched_subheadings = []
    for line in lines:
        text = "".join(char["text"] for char in line).strip()
        if not text:
            continue
        total_chars = len(line)
        total_bold_chars = sum(1 for char in line if is_bold(char.get("fontname", "")))
        overall_line_bold = total_bold_chars / total_chars >= 0.7
        font_counts = {}
        for char in line:
            fnt = char.get("fontname", "")
            font_counts[fnt] = font_counts.get(fnt, 0) + 1
        majority_fontname = max(font_counts, key=font_counts.get)

        char_style_matches = []
        for char in line:
            c_fontname = char.get("fontname", "")
            c_fontsize = char.get("fontsize", char.get("size", 0))
            c_text = char["text"]
            c_all_caps = c_text.isalpha() and c_text.isupper()
            c_bold = is_bold(c_fontname)
            fontname_match = (c_fontname == fontname)
            fontsize_match = abs(c_fontsize - fontsize) < 1
            all_caps_match = (c_all_caps == all_caps)
            bold_match = (overall_line_bold == bold)
            is_char_matching = fontname_match and fontsize_match and all_caps_match and bold_match
            char_style_matches.append(is_char_matching)

        start_noise = 0
        for match in char_style_matches:
            if not match:
                start_noise += 1
            else:
                break

        end_noise = 0
        for match in reversed(char_style_matches):
            if not match:
                end_noise += 1
            else:
                break

        if (start_noise + end_noise) / total_chars > 0.3:
            continue

        start_idx = start_noise
        end_idx = total_chars - 1 - end_noise

        while start_idx <= end_idx and line[start_idx]["text"].isspace():
            start_idx += 1
        while end_idx >= start_idx and line[end_idx]["text"].isspace():
            end_idx -= 1

        cleaned_text = "".join(char["text"] for char in line[start_idx:end_idx+1]).strip()

        if all_caps and not is_all_caps(cleaned_text):
            continue

        matched_subheadings.append({
            "text": cleaned_text,
            "page": line[0]["page"],
            "fontname": majority_fontname,
            "fontsize": fontsize,
            "all_caps": all_caps,
            "bold": bold,
            "y": line[0]["y"],
        })
    return matched_subheadings

def merge_successive_subheadings(subheadings, y_tolerance=25):
    if not subheadings:
        return []
    merged = []
    prev = subheadings[0]
    for curr in subheadings[1:]:
        same_page = (curr["page"] == prev["page"])
        y_close = abs(curr["y"] - prev["y"]) < y_tolerance
        if same_page and y_close:
            prev["text"] += " " + curr["text"]
            prev["y"] = min(prev["y"], curr["y"])
        else:
            merged.append(prev)
            prev = curr
    merged.append(prev)
    return merged

def detect_first_subheading(lines):
    main_heading = None
    found_main_heading = False
    main_heading_page = None
    main_heading_y = None
    candidates = []
    for line in lines:
        text = "".join(char["text"] for char in line).strip()
        if not text:
            continue
        font_counts = {}
        for char in line:
            fnt = char.get("fontname", "")
            font_counts[fnt] = font_counts.get(fnt, 0) + 1
        majority_fontname = max(font_counts, key=font_counts.get)
        fontname = majority_fontname
        fontsize = line[0]["fontsize"]
        page = line[0]["page"]
        y = line[0]["y"]

        if not found_main_heading:
            if len(text.split()) <= 10 and fontsize > 10:
                main_heading = {
                    "text": text,
                    "page": page,
                    "fontname": fontname,
                    "fontsize": fontsize,
                    "y": y,
                }
                main_heading_page = page
                main_heading_y = y
                found_main_heading = True
        else:
            if page == main_heading_page and y > main_heading_y + 0.5:
                total_chars = len(line)
                total_bold_chars = sum(1 for char in line if is_bold(char.get("fontname", "")))
                line_bold = (total_bold_chars / total_chars) >= 0.7
                candidates.append({
                    "text": text,
                    "page": page,
                    "fontname": fontname,
                    "fontsize": fontsize,
                    "y": y,
                    "all_caps": is_all_caps(text),
                    "bold": line_bold,
                })

    if not candidates:
        return []

    filtered_candidates = [c for c in candidates if c["y"] > main_heading["y"] + 0.5]
    if not filtered_candidates:
        return []

    filtered_candidates.sort(key=lambda c: (
        -c["fontsize"],
        -int(c["all_caps"]),
        -int(c["bold"]),
    ))

    first_subheading = filtered_candidates[0]

    matched_subheadings = extract_all_subheadings_with_style(
        lines,
        first_subheading["fontname"],
        first_subheading["fontsize"],
        first_subheading["all_caps"],
        first_subheading["bold"],
    )

    return merge_successive_subheadings(matched_subheadings)

def complete_excel_sheet(pdf_path, excel_path, line_number_map):
    import re

    df = pd.read_excel(excel_path)

    # Read text from PDF pages using pypdf
    reader = PdfReader(pdf_path)
    num_pages = len(reader.pages)

    # Prepare line_number_map by page for quick lookup
    lines_by_page = {}
    for entry in line_number_map:
        lines_by_page.setdefault(entry["page"], []).append(entry)

    updated_rows = []
    for idx, row in df.iterrows():
        subheading_text = row["Subheading"]
        page_no = row["Page No"]
        if page_no not in lines_by_page:
            print(f"Page {page_no} not found in line_number_map for subheading '{subheading_text}'. Skipping.")
            updated_rows.append(row)
            continue

        page_lines = lines_by_page[page_no]
        found_line_number = None

        # Only first 10 chars of subheading, case-insensitive
        subhead_snippet = re.escape(subheading_text[:10].strip().lower())

        # Scan lines for substring match, not full line
        for line_entry in page_lines:
            line_text = "".join(char["text"] for char in line_entry["line"]).strip().lower()
            if re.search(subhead_snippet, line_text):
                found_line_number = line_entry["line_on_page"]
                print(f"Found subheading '{subheading_text}' on page {page_no} at line {found_line_number}.")
                break
        if found_line_number is None:
            print(f"Subheading '{subheading_text}' NOT found on page {page_no}.")
            found_line_number = float('nan')

        # Update dataframe row with found line number
        df.at[idx, "Line on Page"] = found_line_number

    try:
        df.to_excel(excel_path, index=False)
        print(f"Excel sheet updated with line numbers at: {excel_path}")
    except PermissionError:
        print(f"❌ Cannot write to Excel file: {excel_path}. Please close it if it's open and try again.")

def run_subheading_isolator():
    pdf_path = r"C:\Users\ADITYA\Desktop\VSC\LLAMA\All clauses\3.18.pdf"
    output_excel = r"C:\Users\ADITYA\Desktop\VSC\LLAMA\3.18_subheadings.xlsx"

    text_data = extract_text_with_styles(pdf_path)
    grouped_lines = group_text_by_position(text_data)

    # Track line number on each page
    page_line_counter = {}
    line_number_map = []

    for line in grouped_lines:
        page = line[0]["page"]
        page_line_counter[page] = page_line_counter.get(page, 0) + 1
        line_number_map.append({
            "line": line,
            "page": page,
            "line_on_page": page_line_counter[page]
        })

    subheadings = detect_first_subheading(grouped_lines)

    if not subheadings:
        print("No subheadings detected!")
        return

    # Prepare dataframe with subheadings and pages
    data_for_excel = []
    for sh in subheadings:
        data_for_excel.append({
            "Subheading": sh["text"],
            "Page No": sh["page"]
        })

    df = pd.DataFrame(data_for_excel)
    try:
        df.to_excel(output_excel, index=False)
        print(f"Subheadings exported to Excel at: {output_excel}")
    except PermissionError:
        print(f"❌ Cannot write to Excel file: {output_excel}. Please close it if it's open and try again.")
        return

    complete_excel_sheet(pdf_path, output_excel, line_number_map)

def run_subheading_text_extraction():
    import pandas as pd
    import re

    pdf_path = r"C:\Users\ADITYA\Desktop\VSC\LLAMA\All clauses\3.18.pdf"
    excel_path = r"C:\Users\ADITYA\Desktop\VSC\LLAMA\3.18_subheadings.xlsx"

    text_data = extract_text_with_styles(pdf_path)
    grouped_lines = group_text_by_position(text_data)

    # Track line number on each page
    page_line_counter = {}
    line_number_map = []

    for line in grouped_lines:
        page = line[0]["page"]
        page_line_counter[page] = page_line_counter.get(page, 0) + 1
        line_number_map.append({
            "line": line,
            "page": page,
            "line_on_page": page_line_counter[page]
        })

    df = pd.read_excel(excel_path)
    excel_subheadings = df['Subheading'].dropna().tolist()

    def save_text_to_file(text_lines, output_path):
        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(text_lines))
        print(f"Saved extracted text to {output_path}")

    for i in range(len(excel_subheadings) - 1):
        # Regex for first 10 chars of current and next headings, case insensitive
        start_heading_regex = re.escape(excel_subheadings[i][:10].strip().lower())
        next_heading_regex = re.escape(excel_subheadings[i + 1][:10].strip().lower())

        extracted_lines = []
        start_found = False

        for entry in line_number_map:
            line_text = "".join(char["text"] for char in entry["line"]).strip()
            line_text_lower = line_text.lower()

            if len(line_text) < 4:
                continue

            # Start extracting from line where start heading regex matches (include this line)
            if not start_found:
                if re.search(start_heading_regex, line_text_lower):
                    start_found = True
                    extracted_lines.append(line_text)
                continue
            else:
                # Stop extracting if next heading regex matches
                if re.search(next_heading_regex, line_text_lower):
                    break

                extracted_lines.append(line_text)

        if extracted_lines:
            out_file = f"extracted_{i+1}_{excel_subheadings[i][:10].replace(' ','_')}_to_{excel_subheadings[i+1][:10].replace(' ','_')}.txt"
            save_text_to_file(extracted_lines, out_file)
        else:
            print(f"No text found between '{excel_subheadings[i]}' and '{excel_subheadings[i+1]}'")

if __name__ == "__main__":
    run_subheading_isolator()
    run_subheading_text_extraction()
