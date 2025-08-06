import json
from langchain_ollama import OllamaLLM
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain

# Initialize the local model
llm = OllamaLLM(model="mistral:7b-instruct")

# Define the prompt template for summarization
prompt_template = PromptTemplate.from_template("""
You are an expert summarizer. Create a concise summary of the following text, capturing the key points while reducing length significantly.

Text to summarize:
{content}

Provide a brief summary that highlights the most important information.
""")

# Create the chain
chain = LLMChain(llm=llm, prompt=prompt_template)

# Function to summarize content
def summarize_content(text):
    response = chain.invoke({"content": text})["text"]
    return response.strip()

# Example usage
if __name__ == "__main__":
    example_text = """
    Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.
    """
    
    summary = summarize_content(example_text)
    print("Summary:")
    print(summary)
