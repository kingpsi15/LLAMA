import os
import json
import pandas as pd
from langchain_ollama import OllamaLLM
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
import tiktoken

# === Paths ===
input_txt_path = r"C:\Users\ADITYA\Desktop\VSC\LLAMA\extracted_sections\3.1 ACCESS  CONTROL\txt_chunks\extracted_4_AC-4_INFOR_to_AC-5_SEPAR.txt"
output_json_path = "outputs/single_policy.json"
output_excel_path = "outputs/single_policy.xlsx"
output_raw_path = "outputs/single_policy_raw.txt"

os.makedirs(os.path.dirname(output_json_path), exist_ok=True)
os.makedirs(os.path.dirname(output_excel_path), exist_ok=True)

# === Initialize LLM ===
llm = OllamaLLM(model="mistral:7b-instruct")

# === Prompt Template ===
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

# === Count tokens ===
encoding = tiktoken.get_encoding("cl100k_base")
tokens = encoding.encode(normalized_text)
token_count = len(tokens)
print(f"Total tokens in input: {token_count}")

# === Process full text ===
try:
    try:
        response = chain.invoke({"text": normalized_text, "current_heading": "None"})["text"]
        ...
    except Exception as e:
        print(f"Failed to extract data: {e}")

    json_text = response.strip().strip("```").replace("json", "", 1).strip()
    policy_data = json.loads(json_text)

    # Save results
    with open(output_json_path, "w", encoding="utf-8") as jf:
        json.dump(policy_data, jf, indent=2)

    pd.DataFrame([policy_data]).to_excel(output_excel_path, index=False)

    with open(output_raw_path, "w", encoding="utf-8") as rf:
        rf.write("--- RAW RESPONSE ---\n\n")
        rf.write(normalized_text + "\n\n")
        rf.write("--- RESPONSE ---\n\n")
        rf.write(response)

    print("Extraction complete. JSON and Excel saved.")

except Exception as e:
    print(f"Failed to extract data: {e}")


