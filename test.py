import requests

API_KEY = "d808654cdbb2a0b1198468e21ec1ef93f3873c8dc11310336b16acb2fccc7c0f"  # <-- Replace with your actual API key

url = "https://api.together.xyz/v1/chat/completions"

headers = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

data = {
    "model": "mistralai/Mixtral-8x7B-Instruct-v0.1",
    "messages": [
        {
            "role": "user",
            "content": "Give me 3 benefits of using the Mixtral model."
        }
    ],
    "max_tokens": 512,
    "temperature": 0.7,
    "top_p": 0.9
}

response = requests.post(url, headers=headers, json=data)

if response.status_code == 200:
    print("Mixtral says:\n")
    print(response.json()["choices"][0]["message"]["content"])
else:
    print("Error:", response.status_code, response.text)
