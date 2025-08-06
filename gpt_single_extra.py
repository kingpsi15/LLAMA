import os
import json
import pandas as pd
from langchain_ollama import OllamaLLM
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
import tiktoken

# === Paths ===
input_txt_path = r"C:\Users\ADITYA\Desktop\VSC\LLAMA\extracted_sections\3.1 ACCESS  CONTROL\txt_chunks\extracted_1_AC-1_POLIC_to_AC-2_ACCOU.txt"
output_json_path = "outputs/single_policy.json"
output_excel_path = "outputs/single_policy.xlsx"
output_raw_path = "outputs/single_policy_raw.txt"
chunk_output_dir = os.path.join("outputs", "chunks")

os.makedirs(os.path.dirname(output_json_path), exist_ok=True)
os.makedirs(os.path.dirname(output_excel_path), exist_ok=True)
os.makedirs(chunk_output_dir, exist_ok=True)

# === Initialize LLM ===
llm = OllamaLLM(model="mistral:7b-instruct")

# === Enhanced Prompt Template ===
prompt_template = PromptTemplate.from_template("""
You are currently extracting the section: {current_heading}

You are given a raw policy text that includes these sections:

- Control:
- Discussion:
- Related Controls:
- Control Enhancements:
- References:

Your task is to extract the text for each of the above sections from the input chunk.

⚠️ STRICT INSTRUCTIONS:
- Extract every line of content verbatim.
- Never skip lines that appear indented or in bullet/numbered list format (e.g. "a.", "i.", "-", "•", ">>", etc.).
- Treat all indents, bullets, and lines as content.
- Do NOT skip or reformat anything.
- Always move forward — never go back to previous sections.
- If a heading appears that is from an earlier section, include it inside the current section.
- Copy everything under the current section until the *next* valid section appears.
- If a section is not present in this chunk, return "None".
- ❌ Do not use ellipses ("...") or summarization.

Return a JSON object like this:

{{
  "PolicyId": "...",
  "PolicyName": "...",
  "Control": "...",
  "Discussion": "...",
  "RelatedControls": "...", 
  "ControlEnhancements": "...",
  "References": "..."
}}

Extract PolicyId and PolicyName as you see fit.

Document chunk starts below:

{text}

Return only the JSON.
""")

chain = LLMChain(llm=llm, prompt=prompt_template)

# === Load input file ===
with open(input_txt_path, "r", encoding="utf-8") as f:
    raw_text = f.read()
normalized_text = raw_text.replace("\r\n", "\n").replace("\n\n", "\n")

# === Chunking ===
encoding = tiktoken.get_encoding("gpt2")
tokens = encoding.encode(normalized_text)
chunk_size = 500
overlap = 200  # Increased overlap for bullet continuity

chunks = []
start = 0
while start < len(tokens):
    end = min(start + chunk_size, len(tokens))
    chunk_tokens = tokens[start:end]
    chunk_text = encoding.decode(chunk_tokens)
    chunks.append(chunk_text)
    if end == len(tokens):
        break
    start = end - overlap

# Save chunks for debug
for i, chunk in enumerate(chunks, 1):
    chunk_path = os.path.join(chunk_output_dir, f"chunk_{i:02d}.txt")
    with open(chunk_path, "w", encoding="utf-8") as cf:
        cf.write(chunk)
    print(f"Saved chunk {i} to: {chunk_path}")

# === Process chunks ===
partial_results = []
raw_responses = []
current_heading = None

for i, chunk in enumerate(chunks, 1):
    print(f"Processing chunk {i}/{len(chunks)}")
    try:
        response = chain.invoke({"text": chunk, "current_heading": current_heading or "None"})["text"]
        json_text = response.strip().strip("```").replace("json", "", 1).strip()
        partial = json.loads(json_text)
        partial_results.append(partial)
        raw_responses.append(f"--- CHUNK {i} ---\n{chunk}\n\n--- RESPONSE ---\n{response}")

        last_section = None
        for sec in reversed(["Control", "Discussion", "RelatedControls", "ControlEnhancements", "References"]):
            val = partial.get(sec)
            if isinstance(val, list):
                val_str = " ".join(val)
            else:
                val_str = str(val or "")
            if val_str.lower() != "none" and val_str.strip():
                last_section = sec
                break
        if last_section:
            current_heading = last_section

    except Exception as e:
        print(f"Failed on chunk {i}: {e}")

# === Merge Function ===
def merge_sections(results):
    merged = {
        "PolicyId": None,
        "PolicyName": None,
        "Control": [],
        "Discussion": [],
        "RelatedControls": [],
        "ControlEnhancements": [],
        "References": []
    }

    for r in results:
        if merged["PolicyId"] is None and r.get("PolicyId") and r["PolicyId"] != "None":
            merged["PolicyId"] = r["PolicyId"]
        if merged["PolicyName"] is None and r.get("PolicyName") and r["PolicyName"] != "None":
            merged["PolicyName"] = r["PolicyName"]

        for key in ["Control", "Discussion", "RelatedControls", "ControlEnhancements", "References"]:
            val = r.get(key)
            if val:
                cleaned = " ".join(str(x).strip() for x in val if str(x).strip()) if isinstance(val, list) else str(val).strip()
                if cleaned:
                    merged[key].append(cleaned)

    for key in merged:
        if isinstance(merged[key], list):
            merged[key] = "\n".join(merged[key]) if merged[key] else "None"

    return merged

policy_data = merge_sections(partial_results)

# === Save Outputs ===
with open(output_json_path, "w", encoding="utf-8") as jf:
    json.dump(policy_data, jf, indent=2)

pd.DataFrame([policy_data]).to_excel(output_excel_path, index=False)

with open(output_raw_path, "w", encoding="utf-8") as rf:
    rf.write("\n\n--- RAW RESPONSES PER CHUNK ---\n\n")
    rf.write("\n\n\n".join(raw_responses))

print(f"Extraction complete. JSON and Excel saved.")
