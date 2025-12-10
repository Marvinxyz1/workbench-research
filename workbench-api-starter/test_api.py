# -*- coding: utf-8 -*-
"""
KPMG Workbench API テスト

このスクリプトは以下の API をテストします:
1. Chat Completion - チャット生成 API
2. Embeddings - テキスト埋め込み API
3. Text Translation - テキスト翻訳 API
4. Text Chunking - テキスト分割 API (RAG 前処理)
"""

import os
import sys
import io

# Windows コンソールの文字エンコーディングを UTF-8 に設定
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import requests
from pathlib import Path

# .env ファイルから環境変数を読み込む
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env")
except ImportError:
    pass

# =============================================================================
# 設定
# =============================================================================

# API 認証情報
API_KEY = os.getenv("KPMG_WORKBENCH_API_KEY")
CHARGE_CODE = os.getenv("KPMG_CHARGE_CODE", "000000000")

# API エンドポイント
BASE_URL = "https://api.workbench.kpmg"
API_VERSION = "2024-04-01-preview"
TRANSLATOR_API_VERSION = "3.0"

# 利用可能なモデル (Developer Portal でサブスクリプションを確認)
CHAT_MODEL = "gpt-35-turbo-0125-std-ae"
EMBEDDING_MODEL = "text-embedding-3-large-1-std-ae"


# =============================================================================
# API テスト関数
# =============================================================================

def test_chat():
    """
    Chat Completion API のテスト

    Azure OpenAI の Chat Completion API を使用して、
    テキスト生成が正常に動作するか確認します。
    """
    print("\n[Chat Completion テスト]")
    print("-" * 40)

    url = f"{BASE_URL}/genai/azure/inference/chat/completions?api-version={API_VERSION}"

    # リクエストヘッダー
    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "x-kpmg-charge-code": CHARGE_CODE,
        "azureml-model-deployment": CHAT_MODEL,
        "Content-Type": "application/json"
    }

    # リクエストボディ
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
    """
    Embeddings API のテスト

    テキストをベクトル表現に変換する API をテストします。
    RAG (Retrieval-Augmented Generation) システムの構築に使用されます。
    """
    print("\n[Embeddings テスト]")
    print("-" * 40)

    url = f"{BASE_URL}/genai/azure/inference/embeddings?api-version={API_VERSION}"

    # リクエストヘッダー
    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "x-kpmg-charge-code": CHARGE_CODE,
        "azureml-model-deployment": EMBEDDING_MODEL,
        "Content-Type": "application/json"
    }

    # リクエストボディ
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


def test_translator():
    """
    Text Translation API のテスト

    Azure AI Translator を使用して、
    テキストを別の言語に翻訳します。
    """
    print("\n[Text Translation テスト]")
    print("-" * 40)

    url = f"{BASE_URL}/translator/azure/text/translate?api-version={TRANSLATOR_API_VERSION}&to=ja"

    # リクエストヘッダー
    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "x-kpmg-charge-code": CHARGE_CODE,
        "Content-Type": "application/json"
    }

    # リクエストボディ
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


def test_text_chunking():
    """
    Text Chunking API のテスト

    長いテキストを小さなチャンクに分割します。
    RAG システムでドキュメントを処理する前処理に使用されます。
    """
    print("\n[Text Chunking テスト]")
    print("-" * 40)

    url = f"{BASE_URL}/text-chunking/microsoft/simple/simple"

    # リクエストヘッダー
    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "x-kpmg-charge-code": CHARGE_CODE,
        "Content-Type": "application/json"
    }

    # リクエストボディ
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


# =============================================================================
# メイン処理
# =============================================================================

def main():
    """
    メイン関数

    すべての API テストを実行し、結果を表示します。
    """
    print("=" * 60)
    print("KPMG Workbench API テスト")
    print("=" * 60)

    # API Key の確認
    if not API_KEY:
        print("\n[エラー] KPMG_WORKBENCH_API_KEY が設定されていません。")
        print(".env ファイルに API Key を設定してください。")
        return

    # 設定情報の表示
    print(f"\nAPI Key: {API_KEY[:8]}...{API_KEY[-4:]}")
    print(f"Base URL: {BASE_URL}")
    print(f"Chat Model: {CHAT_MODEL}")
    print(f"Embedding Model: {EMBEDDING_MODEL}")

    # テストの実行
    results = {
        "Chat Completion": test_chat(),
        "Embeddings": test_embedding(),
        "Text Translation": test_translator(),
        "Text Chunking": test_text_chunking(),
    }

    # 結果サマリーの表示
    print("\n" + "=" * 60)
    print("テスト結果サマリー")
    print("=" * 60)

    # API カテゴリ別に表示
    categories = {
        "Azure OpenAI - Inference API": ["Chat Completion", "Embeddings"],
        "Azure AI Translator": ["Text Translation"],
        "RAG - Text Processing": ["Text Chunking"],
    }

    for category, tests in categories.items():
        print(f"\n{category}:")
        for test_name in tests:
            if test_name in results:
                icon = "[OK]" if results[test_name] else "[NG]"
                print(f"  {icon} {test_name}")

    # 合計結果
    passed = sum(results.values())
    total = len(results)
    print(f"\n{'=' * 60}")
    print(f"合計: {passed}/{total} 成功 ({passed/total*100:.0f}%)")


if __name__ == "__main__":
    main()
