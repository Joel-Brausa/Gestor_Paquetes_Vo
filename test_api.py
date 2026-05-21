#!/usr/bin/env python3
"""Test OpenRouter API key and model availability"""

import os
from openai import OpenAI
import config

api_key = os.environ.get("OPENROUTER_API_KEY") or config.load_api_key()
# Test with a simpler model that doesn't use reasoning mode
model = "openai/gpt-4o-mini"  # Alternative: "meta-llama/llama-2-7b-chat:free"

print(f"API Key: {api_key[:20]}..." if api_key else "[ERROR] No API key found")
print(f"Model: {model}")
print(f"OpenRouter URL: {config.OPENROUTER_BASE_URL}")
print()

if not api_key:
    print("[ERROR] No API key configured")
    exit(1)

try:
    client = OpenAI(base_url=config.OPENROUTER_BASE_URL, api_key=api_key)
    print("[OK] OpenAI client created")

    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": "Hello, just testing. Respond with 'OK' only."}],
        max_tokens=10,
    )

    print(f"[OK] API call successful")
    print(f"Response object: {response}")
    print(f"Choices: {response.choices}")
    if response.choices:
        print(f"Message: {response.choices[0].message}")
        print(f"Content: {response.choices[0].message.content}")
    else:
        print("[ERROR] No choices in response")

except Exception as e:
    print(f"[ERROR] {type(e).__name__}: {str(e)}")
    exit(1)

print("\n[OK] All tests passed! The API key and model are working correctly.")
