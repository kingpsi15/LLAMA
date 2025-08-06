import pdfplumber
import pandas as pd

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
    main_heading_fontsize = None
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
                main_heading_fontsize = fontsize
                main_heading_y = y
                found_main_heading = True

                last_y = y
                for next_line in lines:
                    next_text = "".join(char["text"] for char in next_line).strip()
                    if not next_text:
                        continue
                    next_page = next_line[0]["page"]
                    next_y = next_line[0]["y"]
                    if next_page != main_heading_page:
                        continue
                    if last_y < next_y <= last_y + 25:
                        total_chars = len(next_line)
                        total_bold_chars = sum(1 for char in next_line if is_bold(char.get("fontname", "")))
                        line_bold = (total_bold_chars / total_chars) >= 0.7
                        line_all_caps = is_all_caps(next_text)
                        fontsize_next = next_line[0]["fontsize"]

                        if (fontname == max(set(char["fontname"] for char in next_line), key=lambda f: sum(1 for c in next_line if c["fontname"] == f)) and
                            abs(fontsize_next - fontsize) < 1 and
                            line_all_caps == is_all_caps(main_heading["text"]) and
                            line_bold == is_bold(fontname)):
                            main_heading["text"] += " " + next_text
                            last_y = next_y
                            main_heading["y"] = max(main_heading["y"], next_y)
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

    epsilon = 0.5
    filtered_candidates = [c for c in candidates if c["y"] > main_heading["y"] + epsilon]
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

def main():
    pdf_path = r"C:\Users\ADITYA\Desktop\VSC\LLAMA\All clauses\3.18.pdf"
    output_excel = r"C:\Users\ADITYA\Desktop\VSC\LLAMA\3.18_subheadings.xlsx"

    text_data = extract_text_with_styles(pdf_path)
    grouped_lines = group_text_by_position(text_data)
    subheadings = detect_first_subheading(grouped_lines)

    data = [{
        "Sl. No": idx + 1,
        "Subheading": sh["text"],
        "Page No": sh["page"]
    } for idx, sh in enumerate(subheadings)]

    df = pd.DataFrame(data)
    df.to_excel(output_excel, index=False)
    print(f"Subheadings extracted and saved to:\n{output_excel}")


if __name__ == "__main__":
    main()
