# -*- coding: utf-8 -*-
"""
KPMG Workbench API Connection Test Script
Test availability of various API endpoints
"""

import os
import sys
import io
import json
import requests
from pathlib import Path

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# 添加项目根目录到路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# 加载环境变量
try:
    from dotenv import load_dotenv
    load_dotenv(project_root / ".env")
except ImportError:
    print("提示: 未安装 python-dotenv，尝试直接从环境变量读取")

# 配置
API_KEY = os.getenv("KPMG_WORKBENCH_API_KEY")
CHARGE_CODE = os.getenv("KPMG_CHARGE_CODE", "000000000")
BASE_URL = "https://api.workbench.kpmg"

# 可用区域
REGIONS = ["eastus", "westeurope", "australiaeast"]

def get_headers(region_override=None):
    """构建请求headers"""
    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "x-kpmg-charge-code": CHARGE_CODE,
        "Content-Type": "application/json"
    }
    if region_override:
        headers["x-kpmg-region-override"] = region_override
    return headers


def test_api_connection():
    """测试基础API连接"""
    print("=" * 60)
    print("KPMG Workbench API 连接测试")
    print("=" * 60)

    if not API_KEY:
        print("错误: 未找到 API_KEY，请检查 .env 文件")
        return False

    print(f"API Key: {API_KEY[:8]}...{API_KEY[-4:]}")
    print(f"Charge Code: {CHARGE_CODE}")
    print(f"Base URL: {BASE_URL}")
    print("-" * 60)

    return True


def test_openai_models():
    """测试获取可用模型列表"""
    print("\n[测试1] 获取可用模型列表")
    print("-" * 40)

    url = f"{BASE_URL}/openai/models"

    try:
        response = requests.get(url, headers=get_headers(), timeout=30)
        print(f"状态码: {response.status_code}")

        if response.status_code == 200:
            data = response.json()
            print("可用模型:")
            if "data" in data:
                for model in data["data"][:10]:  # 只显示前10个
                    print(f"  - {model.get('id', 'Unknown')}")
                if len(data["data"]) > 10:
                    print(f"  ... 还有 {len(data['data']) - 10} 个模型")
            return True
        else:
            print(f"响应: {response.text[:500]}")
            return False
    except requests.exceptions.RequestException as e:
        print(f"请求错误: {e}")
        return False


def test_chat_completion():
    """测试Chat Completion API"""
    print("\n[测试2] Chat Completion API")
    print("-" * 40)

    # 尝试不同的模型名称
    models_to_try = ["gpt-4", "gpt-4o", "gpt-35-turbo", "gpt-4-turbo"]

    for model in models_to_try:
        url = f"{BASE_URL}/openai/deployments/{model}/chat/completions"

        payload = {
            "messages": [
                {"role": "user", "content": "Say 'Hello KPMG' in one word."}
            ],
            "max_tokens": 50,
            "temperature": 0.7
        }

        try:
            print(f"尝试模型: {model}")
            response = requests.post(
                url,
                headers=get_headers(),
                json=payload,
                timeout=60
            )

            print(f"  状态码: {response.status_code}")

            if response.status_code == 200:
                data = response.json()
                if "choices" in data and len(data["choices"]) > 0:
                    content = data["choices"][0].get("message", {}).get("content", "")
                    print(f"  响应: {content}")
                    print(f"  成功! 模型 {model} 可用")
                    return True
            elif response.status_code == 404:
                print(f"  模型 {model} 不存在，尝试下一个...")
            else:
                print(f"  响应: {response.text[:200]}")

        except requests.exceptions.RequestException as e:
            print(f"  请求错误: {e}")

    print("所有模型都无法访问")
    return False


def test_embeddings():
    """测试Embeddings API"""
    print("\n[测试3] Embeddings API")
    print("-" * 40)

    models_to_try = ["text-embedding-ada-002", "text-embedding-3-small", "text-embedding-3-large"]

    for model in models_to_try:
        url = f"{BASE_URL}/openai/deployments/{model}/embeddings"

        payload = {
            "input": "Hello KPMG Workbench"
        }

        try:
            print(f"尝试模型: {model}")
            response = requests.post(
                url,
                headers=get_headers(),
                json=payload,
                timeout=30
            )

            print(f"  状态码: {response.status_code}")

            if response.status_code == 200:
                data = response.json()
                if "data" in data and len(data["data"]) > 0:
                    embedding = data["data"][0].get("embedding", [])
                    print(f"  向量维度: {len(embedding)}")
                    print(f"  成功! 模型 {model} 可用")
                    return True
            elif response.status_code == 404:
                print(f"  模型 {model} 不存在，尝试下一个...")
            else:
                print(f"  响应: {response.text[:200]}")

        except requests.exceptions.RequestException as e:
            print(f"  请求错误: {e}")

    print("所有Embedding模型都无法访问")
    return False


def test_region_connectivity():
    """测试各区域连接性"""
    print("\n[测试4] 区域连接性测试")
    print("-" * 40)

    results = {}

    for region in REGIONS:
        print(f"\n测试区域: {region}")
        url = f"{BASE_URL}/openai/models"

        try:
            response = requests.get(
                url,
                headers=get_headers(region_override=region),
                timeout=30
            )

            print(f"  状态码: {response.status_code}")
            results[region] = response.status_code == 200

            if response.status_code == 200:
                print(f"  {region} 区域可用")
            else:
                print(f"  {region} 区域不可用: {response.text[:100]}")

        except requests.exceptions.RequestException as e:
            print(f"  请求错误: {e}")
            results[region] = False

    return results


def test_anthropic_claude():
    """测试Anthropic Claude API"""
    print("\n[测试5] Anthropic Claude API")
    print("-" * 40)

    # Claude API 端点可能不同，尝试几种可能的格式
    endpoints = [
        f"{BASE_URL}/anthropic/v1/messages",
        f"{BASE_URL}/claude/v1/messages",
        f"{BASE_URL}/v1/messages"
    ]

    payload = {
        "model": "claude-3-sonnet-20240229",
        "max_tokens": 50,
        "messages": [
            {"role": "user", "content": "Say 'Hello' in one word."}
        ]
    }

    for endpoint in endpoints:
        try:
            print(f"尝试端点: {endpoint}")
            response = requests.post(
                endpoint,
                headers=get_headers(),
                json=payload,
                timeout=30
            )

            print(f"  状态码: {response.status_code}")

            if response.status_code == 200:
                data = response.json()
                print(f"  响应: {json.dumps(data, indent=2)[:200]}")
                print("  Claude API 可用!")
                return True
            elif response.status_code != 404:
                print(f"  响应: {response.text[:200]}")

        except requests.exceptions.RequestException as e:
            print(f"  请求错误: {e}")

    print("Claude API 端点未找到或不可用")
    return False


def main():
    """主测试函数"""
    print("\n" + "=" * 60)
    print("   KPMG Workbench API 可用性测试")
    print("=" * 60 + "\n")

    # 检查基础配置
    if not test_api_connection():
        return

    results = {
        "基础连接": False,
        "模型列表": False,
        "Chat Completion": False,
        "Embeddings": False,
        "Claude API": False
    }

    # 运行测试
    results["模型列表"] = test_openai_models()
    results["Chat Completion"] = test_chat_completion()
    results["Embeddings"] = test_embeddings()
    results["Claude API"] = test_anthropic_claude()

    # 区域测试
    region_results = test_region_connectivity()

    # 汇总结果
    print("\n" + "=" * 60)
    print("   测试结果汇总")
    print("=" * 60)

    print("\nAPI 测试:")
    for test_name, passed in results.items():
        status = "通过" if passed else "失败"
        icon = "✓" if passed else "✗"
        print(f"  [{icon}] {test_name}: {status}")

    print("\n区域连接性:")
    for region, available in region_results.items():
        status = "可用" if available else "不可用"
        icon = "✓" if available else "✗"
        print(f"  [{icon}] {region}: {status}")

    # 总结
    passed_count = sum(1 for v in results.values() if v)
    total_count = len(results)
    print(f"\n总计: {passed_count}/{total_count} 项测试通过")

    if passed_count == total_count:
        print("\n所有API测试通过! KPMG Workbench API 可正常使用。")
    elif passed_count > 0:
        print("\n部分API可用，请查看具体错误信息。")
    else:
        print("\n所有测试失败，请检查:")
        print("  1. API Key 是否正确")
        print("  2. 网络连接是否正常")
        print("  3. IP地址是否已加入白名单")
        print("  4. 成员公司是否已完成onboarding")


if __name__ == "__main__":
    main()
