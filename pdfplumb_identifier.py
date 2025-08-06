import pdfplumber

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


def debug_main_and_first_subheading(lines):
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

        fontname = line[0]["fontname"]
        fontsize = line[0]["fontsize"]
        page = line[0]["page"]
        y = line[0]["y"]

        if not found_main_heading:
            # Detect main heading: short line, large font
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
        else:
            # Only lines on same page below main heading
            if page == main_heading_page and y > main_heading_y:
                if fontsize < main_heading_fontsize:
                    candidates.append({
                        "text": text,
                        "page": page,
                        "fontname": fontname,
                        "fontsize": fontsize,
                        "y": y,
                        "all_caps": is_all_caps(text),
                        "bold": is_bold(fontname),
                    })

    if not main_heading:
        print("Main heading not found.")
        return

    print("\n🔷 Detected Main Heading:")
    print(main_heading)

    if not candidates:
        print("\nNo candidate subheadings found after main heading.")
        return

    # Sort candidates by fontsize desc, then all_caps, then bold
    candidates.sort(key=lambda c: (
        -c["fontsize"],
        -int(c["all_caps"]),
        -int(c["bold"]),
    ))

    first_subheading = candidates[0]

    print("\n🔶 First Detected Subheading After Main Heading:")
    print(first_subheading)


def main():
    pdf_path = r"C:\Users\ADITYA\Downloads\NIST.SP.800-53r5_testing_2[1].pdf"
    print("Extracting text and styles from PDF...")
    text_data = extract_text_with_styles(pdf_path)

    print("Grouping text by position...")
    grouped_lines = group_text_by_position(text_data)

    print("Running debug to print main heading and first subheading...")
    debug_main_and_first_subheading(grouped_lines)


if __name__ == "__main__":
    main()
