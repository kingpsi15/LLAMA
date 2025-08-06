import json
import pandas as pd
import re
from langchain_ollama import OllamaLLM
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain

# === Load full raw text ===
with open(r"C:\Users\ADITYA\Desktop\VSC\LLAMA\extracted_sections\3.1 ACCESS  CONTROL\extracted_3_AC-3_ACCES_to_AC-4_INFOR.txt", "r", encoding="utf-8") as f:
    raw_text = f.read()

# === Initialize the local SLM model ===
llm = OllamaLLM(model="mistral:7b-instruct")

# === Prompt for extracting PolicyId and PolicyName from heading using LLM ===
heading_prompt = PromptTemplate.from_template("""
You are given the start of a policy document. 

Extract ONLY these two fields as JSON:

- PolicyId: a short ID like AC-1, SC-7 etc.
- PolicyName: the policy title or name (usually near the ID).

Return a valid JSON object with exactly these fields.

Here is the document start:

{text}
""")

heading_chain = LLMChain(llm=llm, prompt=heading_prompt)

# Use only the first 500 characters for this heading extraction
heading_text = raw_text[:500]

print("Extracting PolicyId and PolicyName from heading with LLM...")
heading_response = heading_chain.invoke({"text": heading_text})["text"]

try:
    json_text = heading_response.strip()
    if json_text.startswith("```") and json_text.endswith("```"):
        json_text = json_text[3:-3].strip()
    if json_text.startswith("```json") and "```" in json_text:
        json_text = json_text[7:].split("```")[0].strip()

    heading_info = json.loads(json_text)
except Exception as e:
    print(f"❌ Failed to parse LLM heading extraction: {e}")
    heading_info = {"PolicyId": "Unknown", "PolicyName": "Unknown"}

# === Ordered headings for regex extraction of sections ===
HEADINGS_ORDER = [
    "Control",
    "Discussion",
    "Related Controls",
    "Control Enhancements",
    "References"
]

def extract_section_ordered(label, text):
    label_escaped = re.escape(label)
    current_idx = HEADINGS_ORDER.index(label)

    # Headings after the current heading
    next_headings = HEADINGS_ORDER[current_idx + 1:]

    lookahead = "|".join([rf"{re.escape(h)}:" for h in next_headings]) if next_headings else "$"

    pattern = rf"{label_escaped}:(.*?)(?=({lookahead})|$)"

    match = re.search(pattern, text, re.DOTALL)
    return match.group(1).strip() if match else "None"

# === Extract other sections with regex ===
policy_data = {
    "PolicyId": heading_info.get("PolicyId", "Unknown"),
    "PolicyName": heading_info.get("PolicyName", "Unknown"),
    "Control": extract_section_ordered("Control", raw_text),
    "Discussion": extract_section_ordered("Discussion", raw_text),
    "RelatedControls": extract_section_ordered("Related Controls", raw_text),
    "ControlEnhancements": extract_section_ordered("Control Enhancements", raw_text),
    "References": extract_section_ordered("References", raw_text),
}

# === Save structured output ===
with open("extracted_policy.json", "w", encoding="utf-8") as f:
    json.dump(policy_data, f, indent=2)

df = pd.DataFrame([policy_data])
df.to_excel("extracted_policy.xlsx", index=False)

print("✅ Extraction complete.")
print("- Saved to extracted_policy.json and extracted_policy.xlsx")
