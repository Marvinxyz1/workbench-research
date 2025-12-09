# -*- coding: utf-8 -*-
"""
KPMG Workbench API Test - Comprehensive Version
Tests all available APIs from the KPMG Workbench Developer Portal
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
TRANSLATOR_API_VERSION = "3.0"

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


def test_translator_text():
    """Test Azure AI Translator - Text Translation API"""
    print("\n[Text Translation Test]")
    print("-" * 40)

    url = f"{BASE_URL}/translator/azure/text/translate?api-version={TRANSLATOR_API_VERSION}&to=ja"
    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "x-kpmg-charge-code": CHARGE_CODE,
        "Content-Type": "application/json"
    }
    payload = [{"Text": "Hello, KPMG Workbench!"}]

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=60)
        print(f"Status: {resp.status_code}")

        if resp.status_code == 200:
            data = resp.json()
            if data and len(data) > 0:
                translations = data[0].get("translations", [])
                if translations:
                    translated = translations[0].get("text", "")
                    print(f"Translated (EN->JA): {translated}")
                    return True
        print(f"Error: {resp.text[:200]}")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False


def test_translator_detect():
    """Test Azure AI Translator - Language Detection API"""
    print("\n[Language Detection Test]")
    print("-" * 40)

    url = f"{BASE_URL}/translator/azure/text/detect?api-version={TRANSLATOR_API_VERSION}"
    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "x-kpmg-charge-code": CHARGE_CODE,
        "Content-Type": "application/json"
    }
    payload = [{"Text": "こんにちは、KPMGワークベンチ！"}]

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=60)
        print(f"Status: {resp.status_code}")

        if resp.status_code == 200:
            data = resp.json()
            if data and len(data) > 0:
                lang = data[0].get("language", "")
                score = data[0].get("score", 0)
                print(f"Detected language: {lang} (score: {score})")
                return True
        print(f"Error: {resp.text[:200]}")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False


def test_text_chunking():
    """Test RAG - Text Chunking API"""
    print("\n[Text Chunking Test]")
    print("-" * 40)

    url = f"{BASE_URL}/text-chunking/microsoft/simple/simple"
    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "x-kpmg-charge-code": CHARGE_CODE,
        "Content-Type": "application/json"
    }
    payload = {
        "extractedFileContent": "This is the first paragraph of a document. It contains important information.\n\nThis is the second paragraph. It has different content that should be in a separate chunk.\n\nThe third paragraph concludes the document with final remarks.",
        "format": "text",
        "algorithm": "line",
        "tokensPerChunk": 50,
        "overlapTokens": 10
    }

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=60)
        print(f"Status: {resp.status_code}")

        if resp.status_code in [200, 202]:
            data = resp.json()
            chunks = data.get("splitText", [])
            print(f"Number of chunks: {len(chunks)}")
            if chunks:
                print(f"First chunk preview: {chunks[0][:50]}...")
            return True
        print(f"Error: {resp.text[:200]}")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False


def test_break_sentence():
    """Test Azure AI Translator - Break Sentence API"""
    print("\n[Break Sentence Test]")
    print("-" * 40)

    url = f"{BASE_URL}/translator/azure/text/breaksentence?api-version={TRANSLATOR_API_VERSION}"
    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "x-kpmg-charge-code": CHARGE_CODE,
        "Content-Type": "application/json"
    }
    payload = [{"Text": "How are you? I am fine. What did you do today?"}]

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=60)
        print(f"Status: {resp.status_code}")

        if resp.status_code == 200:
            data = resp.json()
            if data and len(data) > 0:
                sent_len = data[0].get("sentLen", [])
                detected = data[0].get("detectedLanguage", {})
                print(f"Sentence lengths: {sent_len}")
                print(f"Detected language: {detected.get('language', 'N/A')}")
                return True
        print(f"Error: {resp.text[:200]}")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False


def main():
    print("=" * 60)
    print("KPMG Workbench API Comprehensive Test")
    print("=" * 60)

    if not API_KEY:
        print("Error: KPMG_WORKBENCH_API_KEY not found in .env")
        return

    print(f"API Key: {API_KEY[:8]}...{API_KEY[-4:]}")
    print(f"Base URL: {BASE_URL}")
    print(f"Chat Model: {CHAT_MODEL}")
    print(f"Embedding Model: {EMBEDDING_MODEL}")

    # Run all tests
    results = {
        # Azure OpenAI - Inference API
        "Chat Completion": test_chat(),
        "Embedding": test_embedding(),
        # Azure AI Translator - Text API
        "Text Translation": test_translator_text(),
        "Language Detection": test_translator_detect(),
        "Break Sentence": test_break_sentence(),
        # RAG APIs
        "Text Chunking": test_text_chunking(),
    }

    # Summary
    print("\n" + "=" * 60)
    print("Test Results Summary")
    print("=" * 60)

    # Group by API category
    categories = {
        "Azure OpenAI - Inference API": ["Chat Completion", "Embedding"],
        "Azure AI Translator - Text API": ["Text Translation", "Language Detection", "Break Sentence"],
        "RAG - Text Chunking APIs": ["Text Chunking"],
    }

    for category, tests in categories.items():
        print(f"\n{category}:")
        for test_name in tests:
            if test_name in results:
                icon = "✓" if results[test_name] else "✗"
                print(f"  [{icon}] {test_name}")

    passed = sum(results.values())
    total = len(results)
    print(f"\n{'=' * 60}")
    print(f"Total: {passed}/{total} passed ({passed/total*100:.0f}%)")


if __name__ == "__main__":
    main()
