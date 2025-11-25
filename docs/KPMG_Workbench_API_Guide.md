# KPMG Workbench API 使用指南

> 文档更新日期: 2025-11-25
>
> 本文档总结了KPMG Workbench API的关键信息，帮助开发者快速上手使用平台API。

---

## 1. 平台概述

KPMG Workbench 是 KPMG 的**集中管理的全球AI平台**，为所有成员公司和职能部门提供AI组件、能力和设计系统。

### 核心特点

- **统一全球AI平台**: 集中治理、可扩展的全球AI平台
- **全球数据主权**: 数据存储在特定地理区域（美国、欧洲、澳大利亚）
- **Trusted AI Framework**: AI风险管理和合规框架
- **独特计费系统**: 精确的消费数据和用户收费

---

## 2. API 访问信息

### 2.1 基础URL

| 环境 | URL |
|------|-----|
| **Production API** | `https://api.workbench.kpmg` |
| **Developer Portal** | `https://developer.<region>.workbench.kpmg` |

### 2.2 可用区域 (Hubs)

| 区域代码 | 地理位置 |
|----------|----------|
| `eastus` | 美国东部 |
| `westeurope` | 西欧 |
| `australiaeast` | 澳大利亚东部 |

---

## 3. API 认证

### 3.1 认证方式

KPMG Workbench API 使用 **Azure API Management (APIM)** 进行认证，需要：
1. **IP地址白名单** - 在成员公司onboarding时配置
2. **API Key** - 发放给开发者个人或应用程序

### 3.2 必需的 HTTP Headers

每个API请求都**必须**包含以下headers：

| Header | 值 | 说明 |
|--------|-----|------|
| `Ocp-Apim-Subscription-Key` | `<your-api-key>` | 你的API订阅密钥 |
| `x-kpmg-charge-code` | `000000000` | 费用代码（9位数字） |

### 3.3 可选 HTTP Headers

| Header | 值 | 说明 |
|--------|-----|------|
| `x-kpmg-approval-group` | `<group-value>` | 成员公司提供的审批组值 |
| `x-kpmg-region-override` | `eastus` / `westeurope` / `australiaeast` | 指定路由到的Hub区域，覆盖默认的最低延迟路由 |
| `x-kpmg-failover-deployment` | `<deployment-name>` | 主部署返回错误或429节流错误超过1分钟时使用的故障转移部署名称 |

---

## 4. 可用的 API 产品

### 4.1 LLM 模型 API

| API | 说明 | 文档链接 |
|-----|------|----------|
| **Anthropic Claude API** | Claude模型对话 | [文档](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Workbench-API-Documentation/Anthropic-Claude-API/) |
| **Google Gemini API** | Gemini模型对话 | [文档](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Workbench-API-Documentation/Google-Gemini-API/) |
| **Azure OpenAI Completions & Embeddings API** | 文本补全和向量嵌入 | [文档](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Workbench-API-Documentation/Azure-OpenAI-Completions-Embeddings-API/) |
| **Azure OpenAI Inference API** | 模型推理 | [文档](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Workbench-API-Documentation/Azure-OpenAI-Inference-API/) |
| **Azure OpenAI Assistants API** | AI助手（带记忆和工具） | [文档](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Workbench-API-Documentation/Azure-OpenAI-Assistants-API/) |

### 4.2 数据处理 API

| API | 说明 | 文档链接 |
|-----|------|----------|
| **RAG Ingestion & Retrieval API** | 检索增强生成 | [文档](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Workbench-API-Documentation/Retrieval-Augmented-Generation-%28RAG%29/) |
| **Blob Data Transfer API** | 文件/Blob数据传输 | [文档](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Workbench-API-Documentation/Blob-Data-Transfer-API/) |

### 4.3 翻译 API

| API | 说明 | 文档链接 |
|-----|------|----------|
| **Azure AI Translator Text API** | 文本翻译 | [文档](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Workbench-API-Documentation/Azure-AI-Translator-Text-API/) |
| **Azure AI Translator Document API** | 文档翻译 | [文档](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Workbench-API-Documentation/Azure-AI-Translator-Document-API/) |

### 4.4 其他 API

| API | 说明 | 文档链接 |
|-----|------|----------|
| **Model Reporting API** | 模型使用报告 | [文档](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Workbench-API-Documentation/Azure-OpenAI-Model-Reporting-API/) |
| **Deployments API** | 部署管理 | [文档](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Workbench-API-Documentation/Deployments-API/) |

---

## 5. 代码示例

### 5.1 Python - 基础请求

```python
import requests

# API配置
API_KEY = "<your-api-key>"
CHARGE_CODE = "000000000"
BASE_URL = "https://api.workbench.kpmg"

# 通用headers
headers = {
    "Ocp-Apim-Subscription-Key": API_KEY,
    "x-kpmg-charge-code": CHARGE_CODE,
    "Content-Type": "application/json"
}

# 示例: Azure OpenAI Chat Completion
def chat_completion(messages, model="gpt-4"):
    url = f"{BASE_URL}/openai/deployments/{model}/chat/completions"

    payload = {
        "messages": messages,
        "temperature": 0.7,
        "max_tokens": 1000
    }

    response = requests.post(url, headers=headers, json=payload)
    return response.json()

# 使用示例
result = chat_completion([
    {"role": "system", "content": "You are a helpful assistant."},
    {"role": "user", "content": "Hello, how are you?"}
])
print(result)
```

### 5.2 Python - 指定区域

```python
import requests

headers = {
    "Ocp-Apim-Subscription-Key": "<your-api-key>",
    "x-kpmg-charge-code": "000000000",
    "x-kpmg-region-override": "westeurope",  # 指定欧洲区域
    "Content-Type": "application/json"
}

# 请求将被路由到西欧Hub
response = requests.post(
    "https://api.workbench.kpmg/openai/deployments/gpt-4/chat/completions",
    headers=headers,
    json={"messages": [{"role": "user", "content": "Hello"}]}
)
```

### 5.3 Python - 使用OpenAI SDK (兼容模式)

```python
from openai import AzureOpenAI

client = AzureOpenAI(
    api_key="<your-api-key>",
    api_version="2024-02-15-preview",
    azure_endpoint="https://api.workbench.kpmg"
)

# 注意: 需要在请求中添加自定义headers
response = client.chat.completions.create(
    model="gpt-4",
    messages=[
        {"role": "user", "content": "Hello, how are you?"}
    ],
    extra_headers={
        "x-kpmg-charge-code": "000000000"
    }
)

print(response.choices[0].message.content)
```

### 5.4 cURL 示例

```bash
curl -X POST "https://api.workbench.kpmg/openai/deployments/gpt-4/chat/completions" \
  -H "Ocp-Apim-Subscription-Key: <your-api-key>" \
  -H "x-kpmg-charge-code: 000000000" \
  -H "Content-Type: application/json" \
  -d '{
    "messages": [
      {"role": "user", "content": "Hello, how are you?"}
    ]
  }'
```

---

## 6. 常见错误响应

| HTTP状态码 | 错误原因 | 解决方案 |
|------------|----------|----------|
| **400** | 缺少 `x-kpmg-charge-code` header | 添加费用代码header |
| **401** | 缺少 `Ocp-Apim-Subscription-Key` | 添加API Key header |
| **401** | 无效的API Key | 检查API Key是否正确，是否过期 |
| **404** | 资源未找到（无效URL） | 检查API端点URL是否正确 |
| **429** | 请求频率限制 | 降低请求频率或使用故障转移部署 |

---

## 7. 安全最佳实践

> **重要提示**: API Key 被视为敏感信息（Secrets），必须妥善管理！

### 7.1 密钥管理原则

1. **永远不要**将API Key硬编码在代码中
2. **永远不要**将API Key提交到Git仓库
3. 使用环境变量或安全的密钥管理服务存储API Key
4. 定期轮换API Key
5. 按最小权限原则分配API Key

### 7.2 推荐的密钥存储方式

```python
import os

# 从环境变量读取
API_KEY = os.environ.get("KPMG_WORKBENCH_API_KEY")

# 或使用 .env 文件 (不要提交到Git)
from dotenv import load_dotenv
load_dotenv()
API_KEY = os.getenv("KPMG_WORKBENCH_API_KEY")
```

### 7.3 .gitignore 配置

```gitignore
# 环境变量文件
.env
.env.local
.env.*.local

# API密钥文件
*api_key*
*secrets*
```

更多安全最佳实践请参考: [Secret Management Best Practices](https://handbook.code.kpmg.com/digital-grc/secrets-management-best-practices/)

---

## 8. 开发者资源

### 8.1 官方文档

| 资源 | 链接 |
|------|------|
| **KPMG Workbench 主页** | https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB |
| **Developer Hub** | https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-Developer-and-Product-Manager-Hub.aspx |
| **API 概览** | https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-API-overview.aspx |
| **API 完整文档** | https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Workbench-API-Documentation/ |

### 8.2 学习资源

| 资源 | 链接 |
|------|------|
| **Cookbook (代码示例)** | https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Cookbook/ |
| **Use Cases** | https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/Use-Cases/ |
| **FAQs** | https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/FAQs/ |
| **Developer Wiki** | https://psychic-chainsaw-9pmpg6n.pages.github.io/ |
| **Learning & Development** | https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-learning-development.aspx |

### 8.3 支持渠道

| 类型 | 说明 |
|------|------|
| **技术问题** | 通过成员公司的技术服务台提交 |
| **API Key问题** | 访问 [Workbench Service Catalog](https://kpmggoprod.service-now.com/sp?id=sc_category&sys_id=3cae446893230a10324c76847aba1033) |
| **伦理问题** | 通过成员公司报告程序或 [KPMG International Hotline](https://kpmg.com/xx/en/about/kpmg-international-hotline.html) |

---

## 9. KPMG Workbench 产品生态

### 9.1 核心产品

| 产品 | 说明 |
|------|------|
| **aIQ Chat on KPMG Workbench** | AI聊天应用 |
| **Trusted AI Stamp** | AI系统可信认证标识 |
| **aIQ CaseCraft on KPMG Workbench** | 案例生成工具 |

### 9.2 aIQ CaseCraft Agents

可复用的AI Agent库，用于构建解决方案。详见: [aIQ CaseCraft Agents Inventory](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/aIQ-CaseCraft-Agents-Inventory/)

---

## 10. 附录

### 10.1 API Key 申请流程

1. 确认成员公司已完成KPMG Workbench onboarding
2. 完成必要的培训，获得 **KPMG Workbench Knowledge Badge**
3. 获取成员公司审批人的批准邮件
4. 填写 [Developer Onboarding Request Form](https://kpmggoprod.service-now.com/sp?id=sc_cat_item&sys_id=623c6518c314a61088532485e0013117&sysparm_category=3cae446893230a10324c76847aba1033)
5. 2-3个工作日内通过邮件收到API Key和Developer Portal访问凭证

### 10.2 区域选择建议

| 你的位置 | 推荐区域 |
|----------|----------|
| 亚太地区 | `australiaeast` |
| 欧洲/中东/非洲 | `westeurope` |
| 美洲 | `eastus` |

> 默认情况下，API会自动路由到延迟最低的Hub。只有在特殊需求时才需要使用 `x-kpmg-region-override` 指定区域。

---

## 更新日志

| 日期 | 更新内容 |
|------|----------|
| 2025-11-25 | 初始版本，包含API认证、产品列表、代码示例等 |

---

*本文档基于KPMG Workbench官方文档整理，如有更新请参考官方最新文档。*
