# -*- coding: utf-8 -*-
"""
KPMG Workbench API Test - Simplified Version
"""

import os
import sys
import io
import requests
from pathlib import Path

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# Load .env
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent.parent / ".env")
except ImportError:
    pass

# Configuration
API_KEY = os.getenv("KPMG_WORKBENCH_API_KEY")
CHARGE_CODE = os.getenv("KPMG_CHARGE_CODE", "000000000")
BASE_URL = "https://api.workbench.kpmg"
API_VERSION = "2024-04-01-preview"

# Available models (check Developer Portal for your subscription)
CHAT_MODEL = "gpt-35-turbo-0125-std-ae"
EMBEDDING_MODEL = "text-embedding-3-large-1-std-ae"


def test_chat():
    """Test Chat Completion API"""
    print("\n[Chat Completion Test]")
    print("-" * 40)

    url = f"{BASE_URL}/genai/azure/inference/chat/completions?api-version={API_VERSION}"
    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "x-kpmg-charge-code": CHARGE_CODE,
        "azureml-model-deployment": CHAT_MODEL,
        "Content-Type": "application/json"
    }
    payload = {
        "messages": [{"role": "user", "content": "Say hello in Japanese"}],
        "max_tokens": 50,
        "temperature": 0.7
    }

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=60)
        print(f"Status: {resp.status_code}")

        if resp.status_code == 200:
            data = resp.json()
            content = data.get("choices", [{}])[0].get("message", {}).get("content", "")
            usage = data.get("usage", {})
            print(f"Response: {content}")
            print(f"Tokens: in={usage.get('prompt_tokens')}, out={usage.get('completion_tokens')}")
            return True
        else:
            print(f"Error: {resp.text[:200]}")
            return False
    except Exception as e:
        print(f"Error: {e}")
        return False


def test_embedding():
    """Test Embeddings API"""
    print("\n[Embeddings Test]")
    print("-" * 40)

    url = f"{BASE_URL}/genai/azure/inference/embeddings?api-version={API_VERSION}"
    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "x-kpmg-charge-code": CHARGE_CODE,
        "azureml-model-deployment": EMBEDDING_MODEL,
        "Content-Type": "application/json"
    }
    payload = {"input": ["Hello KPMG Workbench"]}

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=60)
        print(f"Status: {resp.status_code}")

        if resp.status_code == 200:
            data = resp.json()
            if "data" in data and len(data["data"]) > 0:
                dim = len(data["data"][0].get("embedding", []))
                print(f"Embedding dimension: {dim}")
                return True
        print(f"Error: {resp.text[:200]}")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False


def main():
    print("=" * 50)
    print("KPMG Workbench API Test")
    print("=" * 50)

    if not API_KEY:
        print("Error: KPMG_WORKBENCH_API_KEY not found in .env")
        return

    print(f"API Key: {API_KEY[:8]}...{API_KEY[-4:]}")
    print(f"Model: {CHAT_MODEL}")

    # Run tests
    results = {
        "Chat": test_chat(),
        "Embedding": test_embedding()
    }

    # Summary
    print("\n" + "=" * 50)
    print("Results:")
    for name, passed in results.items():
        icon = "✓" if passed else "✗"
        print(f"  [{icon}] {name}")

    passed = sum(results.values())
    print(f"\nTotal: {passed}/{len(results)} passed")


if __name__ == "__main__":
    main()
