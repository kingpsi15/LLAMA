# === START: final_main_all.py logic ===
import sys
sys.setrecursionlimit(3000)  # default is 1000, this raises it safely

import os
import re
import pandas as pd
import PyPDF2
import glob

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

# === MAIN RUN ===

pdf_path = r"C:\Users\ADITYA\Downloads\NIST.SP.800-53r5[1].pdf"
toc_file = 'table_of_contents.txt'
excel_file = 'structured_toc.xlsx'
output_dir = 'extracted_sections'

toc_page = find_toc_page(pdf_path)
if toc_page:
    extract_toc_to_text(pdf_path, toc_page, toc_file)
    process_toc_and_save_to_excel(toc_file, excel_file)
else:
    print("TOC not found.")
    exit()

offset = find_offset(pdf_path)
if offset:
    offset -= 1
else:
    print("Offset not found.")
    exit()

# === SPLIT MAIN PDF INTO SECTION PDFs ===
reader = PyPDF2.PdfReader(pdf_path)
total_pages = len(reader.pages)
df = pd.read_excel(excel_file, dtype={'Section': str})
df = df.sort_values('Page Number').reset_index(drop=True)

for i in range(len(df)):
    start_section = df.iloc[i]
    end_page_number = df.iloc[i + 1]['Page Number'] if i < len(df) - 1 else total_pages

    start_page = int(start_section['Page Number']) + offset - 1
    end_page = int(end_page_number) + offset - 1

    section_id = str(start_section['Section']).strip()
    section_name = str(start_section['Section Name']).strip().replace('/', '-')
    section_folder = os.path.join(output_dir, f"{section_id} {section_name}")
    os.makedirs(section_folder, exist_ok=True)

    output_pdf_path = os.path.join(section_folder, f"{section_id}.pdf")
    writer = PyPDF2.PdfWriter()
    for page_num in range(start_page, end_page):
        if 0 <= page_num < total_pages:
            writer.add_page(reader.pages[page_num])
    with open(output_pdf_path, 'wb') as f:
        writer.write(f)
    print(f"[+] Saved section PDF: {output_pdf_path}")

# === CONTINUES in next message with final_sub_all logic and subheading integration ===
# === START: final_sub_all.py logic (slightly adapted for loop) ===

import pdfplumber
import pandas as pd
from pypdf import PdfReader
import re
import json
import os

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

    for char_data in sorted(text_data, key=lambda x: (x["page"], x["y"], x["x"])):
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
    df = pd.read_excel(excel_path)
    reader = PdfReader(pdf_path)
    lines_by_page = {}
    for entry in line_number_map:
        lines_by_page.setdefault(entry["page"], []).append(entry)

    for idx, row in df.iterrows():
        subheading_text = row["Subheading"]
        page_no = row["Page No"]
        if page_no not in lines_by_page:
            continue
        page_lines = lines_by_page[page_no]
        found_line_number = None
        subhead_snippet = re.escape(subheading_text[:10].strip().lower())
        for line_entry in page_lines:
            line_text = "".join(char["text"] for char in line_entry["line"]).strip().lower()
            if re.search(subhead_snippet, line_text):
                found_line_number = line_entry["line_on_page"]
                break
        df.at[idx, "Line on Page"] = found_line_number if found_line_number else float('nan')

    df.to_excel(excel_path, index=False)

def extract_last_subheading_to_section_end(last_subheading, line_number_map, output_path, json_dir):
    start_regex = re.escape(last_subheading['text'][:10].strip().lower())
    start_found = False
    extracted_lines = []
    for entry in line_number_map:
        line_text = "".join(char["text"] for char in entry["line"]).strip()
        line_text_lower = line_text.lower()
        if len(line_text) < 4:
            continue
        if not start_found:
            if re.search(start_regex, line_text_lower):
                start_found = True
                extracted_lines.append(line_text)
            continue
        else:
            extracted_lines.append(line_text)
    if extracted_lines:
        full_text = "\n".join(extracted_lines)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(full_text)
        json_path = os.path.join(json_dir, os.path.basename(output_path).replace(".txt", ".json"))
        json_data = {
            "subheading": last_subheading['text'],
            "start_text": full_text[:50],
            "content": full_text
        }
        with open(json_path, "w", encoding="utf-8") as jf:
            json.dump(json_data, jf, indent=2)
        print(f"✅ Saved last subheading extract to {output_path} and JSON")

# === PER-SECTION LOOP ===
for folder in os.listdir(output_dir):
    folder_path = os.path.join(output_dir, folder)
    if not os.path.isdir(folder_path):
        continue

    txt_dir = os.path.join(folder_path, "txt_chunks")
    json_dir = os.path.join(folder_path, "json_chunks")
    os.makedirs(txt_dir, exist_ok=True)
    os.makedirs(json_dir, exist_ok=True)

    pdf_file = None
    for f in os.listdir(folder_path):
        if f.endswith(".pdf"):
            pdf_file = os.path.join(folder_path, f)
            break
    if not pdf_file:
        continue

    print(f"\n[🔍] Subheading extraction for: {folder}")

    text_data = extract_text_with_styles(pdf_file)
    grouped_lines = group_text_by_position(text_data)

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
        print(f"[⚠️] No subheadings found in {folder}")
        continue

    excel_path = os.path.join(folder_path, f"{folder.split()[0]}_subheadings.xlsx")
    df = pd.DataFrame([{"Subheading": s["text"], "Page No": s["page"]} for s in subheadings])
    df.to_excel(excel_path, index=False)
    complete_excel_sheet(pdf_file, excel_path, line_number_map)

    for i in range(len(subheadings) - 1):
        start_regex = re.escape(subheadings[i]['text'][:10].strip().lower())
        end_regex = re.escape(subheadings[i+1]['text'][:10].strip().lower())
        extracted_lines = []
        start_found = False
        for entry in line_number_map:
            line_text = "".join(char["text"] for char in entry["line"]).strip()
            line_text_lower = line_text.lower()
            if len(line_text) < 4:
                continue
            if not start_found:
                if re.search(start_regex, line_text_lower):
                    start_found = True
                    extracted_lines.append(line_text)
                continue
            else:
                if re.search(end_regex, line_text_lower):
                    break
                extracted_lines.append(line_text)
        if extracted_lines:
            filename_base = f"extracted_{i+1}_{subheadings[i]['text'][:10].replace(' ','_')}_to_{subheadings[i+1]['text'][:10].replace(' ','_')}"
            txt_path = os.path.join(txt_dir, f"{filename_base}.txt")
            json_path = os.path.join(json_dir, f"{filename_base}.json")
            full_text = "\n".join(extracted_lines)

            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(full_text)

            json_data = {
                "subheading": subheadings[i]['text'],
                "start_text": full_text[:50],
                "content": full_text
            }
            with open(json_path, "w", encoding="utf-8") as jf:
                json.dump(json_data, jf, indent=2)

    # 🔚 Handle last subheading to section end
    last_sh = subheadings[-1]
    last_out_txt = os.path.join(txt_dir, f"extracted_last_{last_sh['text'][:10].replace(' ','_')}_to_end.txt")
    extract_last_subheading_to_section_end(last_sh, line_number_map, last_out_txt, json_dir)

# === CLEANUP: Remove section PDFs and subheadings Excel files ===
for folder in os.listdir(output_dir):
    folder_path = os.path.join(output_dir, folder)
    if not os.path.isdir(folder_path):
        continue

    # Delete PDF file
    for f in os.listdir(folder_path):
        if f.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, f)
            try:
                os.remove(pdf_path)
            except Exception as e:
                print(f"[❌] Failed to delete PDF {pdf_path}: {e}")

    # Delete Excel file
    section_id = folder.split()[0]  # e.g., '3.1' from '3.1 Section Name'
    excel_filename = f"{section_id}_subheadings.xlsx"
    excel_path = os.path.join(folder_path, excel_filename)
    if os.path.exists(excel_path):
        try:
            os.remove(excel_path)
        except Exception as e:
            print(f"[❌] Failed to delete Excel {excel_path}: {e}")

# Delete main TXT file
    txt_filename = f"{folder}.txt"  # e.g., '3.1 Section Name.txt'
    txt_path = os.path.join(folder_path, txt_filename)
    if os.path.exists(txt_path):
        try:
            os.remove(txt_path)
        except Exception as e:
            print(f"[❌] Failed to delete TXT {txt_path}: {e}")

import os
import json
import pandas as pd
from langchain_ollama import OllamaLLM
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
import re

# === Directories ===
base_dir = "extracted_sections"
output_json_dir = os.path.join("outputs", "json")
output_excel_dir = os.path.join("outputs", "excel")
master_excel_path = os.path.join("outputs", "master_subpolicies.xlsx")

os.makedirs(output_json_dir, exist_ok=True)
os.makedirs(output_excel_dir, exist_ok=True)

# === Initialize LLM ===
llm = OllamaLLM(model="mistral:7b-instruct", options={"num_gpu_layers": 0})

# === Prompt (strict, unchanged) ===
prompt_template = PromptTemplate.from_template("""
You are given a raw policy text with clearly labeled sections such as:

- Control:
- Discussion:
- Related Controls:
- Control Enhancements:
- References:

Your task is to extract and return only the exact text under each section.

⚠️ INSTRUCTIONS:
- Match headings **exactly** (including the colon).
- Copy everything under each heading **verbatim** until the next valid heading.
- If a section is missing, return "None".
- Do NOT summarize or reword anything.

You must return a valid JSON object in this format:

{{
  "PolicyId": "...",              // e.g., AC-2
  "PolicyName": "...",            // title of the policy before first heading
  "Control": "...",               
  "Discussion": "...",
  "ControlEnhancements": "...",
  "RelatedControls": "...",
  "References": "..."
}}

Examples of headings (match exactly):  
Control:    
Discussion:  
Control Enhancements:  
Related Controls:  
References:

Infer only the PolicyId and PolicyName from the first few lines before any section. For all other fields, extract **only what appears under labeled sections**.

Here is the document:

{text}

Return only the JSON object. Do not include any extra explanation or markdown.
""")

chain = LLMChain(llm=llm, prompt=prompt_template)

# === State ===
subpolicy_counter = 1
master_rows = []
last_heading_written = None

# === Traverse all folders ===
for section in sorted(os.listdir(base_dir)):
    section_path = os.path.join(base_dir, section)
    if not os.path.isdir(section_path):
        continue

    txt_dir = os.path.join(section_path, "txt_chunks")
    if not os.path.isdir(txt_dir):
        print(f"⚠️ No txt_chunks in {section}")
        continue

    # Insert heading row once per section
    if last_heading_written != section:
        heading_row = {"PolicyId": f"HEADING {section}"}
        master_rows.append(heading_row)
        last_heading_written = section

    for file in sorted(os.listdir(txt_dir)):
        if not file.endswith(".txt"):
            continue

        txt_path = os.path.join(txt_dir, file)
        print(f"🔍 Processing: {txt_path}")

        with open(txt_path, "r", encoding="utf-8") as f:
            raw_text = f.read()

        normalized_text = raw_text.replace("\r\n", "\n").replace("\n\n", "\n")

        try:
            response = chain.invoke({"text": normalized_text})["text"]

            json_text = response.strip()
            if json_text.startswith("```"):
                json_text = json_text.strip("```").strip()
            if json_text.startswith("json"):
                json_text = json_text[4:].strip()

            import re

            # Remove any markdown code fences ```json ... ```
            json_text = re.sub(r"```(json)?\s*|\s*```", "", json_text)
            json_text = re.sub(r'\\(?![\"\\/bfnrt])', r'\\\\', json_text)

            # Remove any lines not part of JSON (e.g., lines before the first { or after the last })
            json_start = json_text.find('{')
            json_end = json_text.rfind('}')
            if json_start != -1 and json_end != -1:
                json_text = json_text[json_start:json_end+1]

            # Optional: strip leading/trailing whitespace again after slicing
            json_text = json_text.strip()

            try:
                policy_data = json.loads(json_text, strict=False)

                # Your existing handling here...
                for key in policy_data:
                    if isinstance(policy_data[key], list):
                        policy_data[key] = "; ".join(str(item) for item in policy_data[key])

            except json.JSONDecodeError as json_err:
                print(f"❌ JSON decode error in {file}: {json_err}")

                error_log_path = os.path.join("outputs", "json", "error_" + file.replace(".txt", ".json"))
                with open(error_log_path, "w", encoding="utf-8") as ef:
                    ef.write(json_text)
                
                continue  # Skip this file and move on

            # === Save files ===
            # === Clean Policy ID ===
            raw_policy_id = policy_data.get("PolicyId", "").strip()

            # Use fallback if missing or invalid
            if not raw_policy_id:
                raw_policy_id = f"subpolicy_{subpolicy_counter}"

            policy_id = raw_policy_id.replace("/", "-").replace("\\", "-").replace(" ", "_")

            subpolicy_counter += 1

            # === Create section folders ===
            json_section_dir = os.path.join(output_json_dir, section)
            excel_section_dir = os.path.join(output_excel_dir, section)
            os.makedirs(json_section_dir, exist_ok=True)
            os.makedirs(excel_section_dir, exist_ok=True)

            # === Output paths ===
            json_path = os.path.join(json_section_dir, f"{policy_id}.json")
            excel_path = os.path.join(excel_section_dir, f"{policy_id}.xlsx")
            raw_path = os.path.join(json_section_dir, f"{policy_id}_raw.txt")


            with open(json_path, "w", encoding="utf-8") as jf:
                json.dump(policy_data, jf, indent=2)

            pd.DataFrame([policy_data]).to_excel(excel_path, index=False)

            with open(raw_path, "w", encoding="utf-8") as rf:
                rf.write(response)

            print(f"✅ Saved: {policy_id}.json + .xlsx")

            # === Append to master + save incrementally ===
            master_rows.append(policy_data)
            master_df = pd.DataFrame(master_rows)
            master_df.to_excel(master_excel_path, index=False)

        except Exception as e:
            print(f"❌ Failed on {file}: {e}")

print(f"\n✅ Final master Excel saved to: {master_excel_path}")


