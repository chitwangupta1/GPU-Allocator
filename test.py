import requests
import os

API_KEY = os.getenv("GOOGLE_API_KEY")

url = f"https://generativelanguage.googleapis.com/v1beta/models?key={API_KEY}"
res = requests.get(url)

print(res.json())
