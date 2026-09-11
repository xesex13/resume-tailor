import os, anthropic
from dotenv import load_dotenv

load_dotenv(override=True)
client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY").strip())

try:
    models = client.models.list()
    print("Available Models for your Key:")
    for m in models.data:
        print(f" - {m.id}")
except Exception as e:
    print(f"Failed: {e}")