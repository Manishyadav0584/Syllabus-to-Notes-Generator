
import os
from dotenv import load_dotenv
from groq import Groq

load_dotenv()

api_key = os.getenv("GROQ_API_KEY")
model = os.getenv("GROQ_MODEL", "openai/gpt-oss-120b")

print("Checking Groq connection...")

if not api_key:
    print("ERROR: Groq API key not found in .env")
    raise SystemExit(1)

try:
    client = Groq(
        api_key=api_key,
        timeout=60.0,
        max_retries=1
    )

    response = client.chat.completions.create(
        model=model,
        messages=[
            {
                "role": "user",
                "content": "Explain Generative AI in three simple sentences."
            }
        ],
        max_completion_tokens=200
    )

    print("\nGROQ CONNECTION SUCCESSFUL!")
    print("-" * 50)
    print(response.choices[0].message.content)

except Exception as error:
    print("\nGROQ TEST FAILED!")
    print(type(error).__name__)
    print(error)
