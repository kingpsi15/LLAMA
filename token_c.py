import tiktoken

# === File path ===
file_path = r"C:\Users\ADITYA\Desktop\VSC\LLAMA\extracted_sections\3.1 ACCESS  CONTROL\txt_chunks\extracted_4_AC-4_INFOR_to_AC-5_SEPAR.txt"

# === Load text ===
with open(file_path, "r", encoding="utf-8") as f:
    text = f.read()

# === Tokenizer (e.g., for GPT-3.5 / GPT-4) ===
encoding = tiktoken.encoding_for_model("gpt2")
tokens = encoding.encode(text)

# === Results ===
print(f"Character count: {len(text)}")
print(f"Token count: {len(tokens)}")
