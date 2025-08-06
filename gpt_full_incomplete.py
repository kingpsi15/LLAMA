import os
import json
import pandas as pd
from langchain_ollama import OllamaLLM
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
import re
import tiktoken

# Load a tokenizer compatible with Mistral
enc = tiktoken.get_encoding("cl100k_base")
# === Directories ===
base_dir = "extracted_sections"
output_json_dir = os.path.join("outputs", "json")
output_excel_dir = os.path.join("outputs", "excel")
master_excel_path = os.path.join("outputs", "master_subpolicies.xlsx")

os.makedirs(output_json_dir, exist_ok=True)
os.makedirs(output_excel_dir, exist_ok=True)

# === Initialize LLM ===
llm = OllamaLLM(model="mistral:7b-instruct", options={"num_gpu_layers": 0})

# === Prompt ===
prompt_template = PromptTemplate.from_template("""
The following is a raw policy text that contains clearly marked sections like:

- Control:
- Description:
- Control Enhancements:
- Related Controls:
- References:

Your task is to extract and return a JSON object with these fields:

1. **PolicyId**: A unique identifier if one is mentioned at the start (e.g., AC-1, SC-7). If none is found, return "Unknown".
2. **PolicyName**: The title or name of the policy if present before the Control section. If not available, return "Unknown".
3. **Control**: Full text under the Control section.
4. **Description**: Full text under the Description section.
5. **Discussion**: If there is a separate section labeled “Discussion”, return it; otherwise return "None".
6. **ControlEnhancements**: Full text under the Control Enhancements section. If not present, return "None".
7. **RelatedControls**: List or paragraph from Related Controls section. If not present, return "None".
8. **References**: All citations or documents listed under References. If not present, return "None".

Here is the text:

{text}

Return only a valid JSON object. Do not include any extra explanation or markdown.
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
        # Count tokens
        token_count = len(enc.encode(normalized_text))

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
            # === Clean Policy ID ==
            policy_data["TokenCount"] = token_count
            raw_policy_id = policy_data.get("PolicyId", "").strip()

            # If missing or marked as "Unknown", use the .txt filename (without .txt)
            if not raw_policy_id or raw_policy_id.lower() == "unknown":
                raw_policy_id = os.path.splitext(file)[0]

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

