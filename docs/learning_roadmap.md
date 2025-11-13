# KPMG Workbench 学习路线图
## 从零到获得API Key和GitHub EMU访问权限

> **目标**: 获得KPMG Workbench的API Key,开始实际开发测试

---

## 📋 完整流程概览

```
Step 1: GitHub EMU账号
   ↓
Step 2: Prerequisites认证(2选1完成)
   ↓
Step 3: 提交1个Pull Request到GitHub EMU
   ↓
Step 4: Developer Learning Path(9个模块,5.3小时视频)
   ↓
Step 5: Knowledge Badge Assessment(考试)
   ↓
Step 6: 获得Digital Badge + API Key ✅
```

**预计总时间:** 2-7天(取决于是否已有Prerequisites认证)

---

## Step 1: 获得GitHub EMU账号

### 🎯 目标:
获得KPMG的GitHub Enterprise Managed User (EMU)账号

### 📍 操作步骤:

#### 1.1 检查你是否已有GitHub EMU账号
- 访问: https://github.com/enterprises/kpmg
- 尝试用你的KPMG邮箱登录
- 如果能登录,跳到Step 2

#### 1.2 如果没有账号,申请GitHub EMU
- **申请页面**: https://handbook.code.kpmg.com/KPMG-Code/GitHub/Organization%20onboarding/
- **流程**:
  1. 访问KPMG Code Handbook
  2. 找到"GitHub EMU Onboarding"
  3. 填写申请表
  4. 等待IT部门审批(通常1-3个工作日)

#### 1.3 验证GitHub EMU账号可用
- [ ] 能登录 https://github.com/enterprises/kpmg
- [ ] 能看到KPMG的组织仓库
- [ ] 能创建或fork仓库

**⚠️ 如果卡在这一步:** 联系你的IT部门或AI Lead

---

## Step 2: Prerequisites - 推荐认证(2选1完成即可)

### 🎯 目标:
完成至少2个推荐认证,证明你有基础技术能力

### 📍 认证选项:

官方要求是"**2 or more**",以下认证任选2个完成:

#### 选项A: Azure Fundamentals AZ-900 ⭐推荐
- **学习页面**: https://learn.microsoft.com/en-us/credentials/certifications/azure-fundamentals/?practice-assessment-type=certification
- **内容**: Azure基础知识(云计算、存储、网络、安全)
- **难度**: ⭐⭐ (入门级)
- **时间**: 8-12小时学习 + 45分钟考试
- **费用**: $99 USD(KPMG可能报销,询问你的Lead)

#### 选项B: Azure AI Fundamentals AI-900 ⭐⭐推荐(与Workbench最相关)
- **学习页面**: https://learn.microsoft.com/en-us/credentials/certifications/azure-ai-fundamentals/?practice-assessment-type=certification
- **内容**: Azure AI服务、机器学习基础、OpenAI服务
- **难度**: ⭐⭐ (入门级)
- **时间**: 6-10小时学习 + 45分钟考试
- **费用**: $99 USD
- **为什么推荐**: 与Workbench的AI服务直接相关!

#### 选项C: GitHub Foundations ⭐最简单
- **学习页面**: https://learn.microsoft.com/en-us/collections/o1njfe825p602p
- **内容**: GitHub基础操作、Git版本控制、协作流程
- **难度**: ⭐ (最简单)
- **时间**: 4-6小时学习 + 考试
- **费用**: 免费

#### 选项D: GitHub Actions
- **学习页面**: https://learn.microsoft.com/en-us/collections/n5p4a5z7keznp5
- **内容**: CI/CD、自动化工作流
- **难度**: ⭐⭐⭐
- **时间**: 8-12小时

#### 选项E: Responsible AI
- **学习页面**: https://app.pluralsight.com/library/courses/artificial-intelligence-essentials-responsible-ai/table-of-contents
- **内容**: AI伦理、Trusted AI原则
- **难度**: ⭐⭐
- **时间**: 3-5小时
- **注意**: 需要Pluralsight账号(KPMG应该有企业账号)

### 💡 我的推荐组合:

**快速路线(最省时间):**
- ✅ GitHub Foundations (4-6小时,免费)
- ✅ Responsible AI (3-5小时,免费)
- **总时间: 7-11小时**

**最相关路线(最有用):**
- ✅ Azure AI Fundamentals AI-900 (6-10小时)
- ✅ GitHub Foundations (4-6小时)
- **总时间: 10-16小时**

---

## Step 3: 提交1个Pull Request到GitHub EMU

### 🎯 目标:
证明你会用GitHub EMU,满足前置要求

### 📍 操作步骤:

#### 3.1 找一个KPMG的GitHub仓库
- 访问: https://github.com/enterprises/kpmg
- 找到任何一个你有权限的仓库
- 或者创建一个测试仓库

#### 3.2 提交一个简单的Pull Request
**最简单的方法:**
1. Fork一个仓库(或在你有权限的仓库中直接操作)
2. 修改README.md,加一行"Test PR for Workbench onboarding"
3. 创建一个分支
4. 提交Pull Request
5. 自己或让同事review并merge

**目的:** 只是为了满足"至少1个PR"的要求,不需要复杂的代码

---

## Step 4: Developer Learning Path(核心!)

### 🎯 目标:
完成9个学习模块,获得KPMG Workbench知识

### 📍 学习页面:
**主页面**: https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-learning-development.aspx

**Developer Learning Path入口**: https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-learning-and-development-development-track.aspx

### 📚 9个模块详细清单

| # | 模块名称 | 时长 | 核心内容 | 重点关注 |
|---|---------|------|---------|---------|
| **1** | Introduction to KPMG Workbench | 54分钟 | 背景、AI战略、Workbench定位 | 实用性评分、是否可跳过 |
| **2** | Revolutionizing AI Productivity: Dive into KPMG Workbench | 35分钟 | 平台架构、Azure设计、**部署模式**⚠️ | **部署模式!能否导出代码?** |
| **3** | Deep Dive: Inference API | 26分钟 | 推理API的访问、认证、使用 | API使用方式、认证流程 |
| **4** | Deep Dive: Completion API | 28分钟 | Completion API、模型、Payload | API兼容性、支持的模型 |
| **5** | RAG: Overview and Building Blocks | 49分钟 | RAG的6个组件详细介绍 | 组件架构、配置复杂度 |
| **6** | RAG: Leading Practices | 53分钟 | RAG最佳实践、用例 | 实际案例、最佳实践 |
| **7** | Tailoring KPMG Workbench for Global: Feature Flags | 13分钟 | 功能定制、区域合规 | **日本数据中心?数据主权?** |
| **8** | Designing AI Experiences with KPMG Workbench | 39分钟 | Design Systems、组件库 | 组件质量、技术栈 |
| **9** | Building Better, Faster: Guide to Developer Resources | 21分钟 | 开发者资源、文档、**部署**⚠️ | **能否导出?部署方式?** |

**总计:** 318分钟(~5小时18分钟纯视频时间)

### 📝 学习方式:

#### 方式1: 在SharePoint上直接学习(推荐)
- 访问Developer Learning Path页面
- 按顺序观看9个模块视频
- **视频链接都在页面上**

#### 方式2: 在GLMS(Global Learning Management System)上学习
- **从2025年5月1日起,正式路径改为GLMS**
- **Program Name**: GX25_PRO_KPMG Workbench for Developers
- **Program ID**: GX25_CFS_DDF_AI_BLDG_WB_D_PRO
- **GLMS链接**: https://hcm-eu20.hr.cloud.sap/sf/learning?destUrl=https://kpmgic.lms.hr.cloud.sap/learning/user/deeplink_redirect.jsp?linkId%3dPROGRAM_DETAILS%26programID%3dGX25_CFS_DDF_AI_BLDG_WB_D_PRO%26fromSF%3dY&company=KPMGProd
- **如果找不到:** 联系你的L&D部门,提供Program Name和ID

### 🔍 关键学习重点:

**模块2和模块9最关键!** 重点关注:
- [ ] Workbench是"开发平台"还是"托管平台"?
- [ ] 代码能否导出?
- [ ] 能否部署到客户环境?
- [ ] 日本是否有数据中心?
- [ ] API调用能否切换到客户自己的账号?

### 📊 学习时记录模板:

```
模块#: [1-9]
模块名称: [XXX]
观看日期: [YYYY-MM-DD]
实际花费时间: [X小时Y分钟]
实用性评分: [1-5分]
关键收获:
  1. [XXX]
  2. [XXX]
  3. [XXX]
是否可跳过: [是/否]
疑问点:
  - [XXX]
```

---

## Step 5: Knowledge Badge Assessment(考试)

### 🎯 目标:
通过考试,证明你掌握了Workbench知识

### 📍 考试信息:

#### 考试入口:
- 完成所有9个模块后,会出现Assessment链接
- 或在GLMS系统中自动解锁考试

#### 考试形式(推测):
- **题型:** 选择题/判断题
- **题量:** 20-50题
- **时间:** 30-60分钟
- **通过分数:** 70-80%

#### 考试准备:
- 认真观看每个模块,做笔记
- 重点记忆:
  - Workbench的核心价值
  - Trusted AI原则
  - RAG的6个组件
  - API的认证方式
  - Data sovereignty的地区

---

## Step 6: 获得Digital Badge + API Key🎉

### 🎯 目标:
通过考试后自动获得认证和API Key

### 📍 你会得到:

#### 6.1 KPMG Workbench Knowledge Badge
- 数字徽章,发到邮箱
- 可以加到LinkedIn

#### 6.2 API Key(最重要!)
- 获得方式:
  - 邮件通知
  - 在GLMS或Developer Hub显示
  - 在portal申请
- **API Key格式:** 类似 `wb_xxxxxxxxxxxxxxxxxxxxxxxx`

#### 6.3 访问权限
有了API Key后,你可以:
- 调用Completion API
- 调用Inference API
- 使用RAG服务
- 访问Design Systems组件库
- 查看完整开发者文档

### 📍 验证API Key:

测试调用Completion API:
```bash
curl -X POST https://workbench.kpmg.com/api/v1/completion \
  -H "Authorization: Bearer YOUR_API_KEY" \
  -H "Content-Type: application/json" \
  -d '{
    "messages": [{"role": "user", "content": "Hello!"}],
    "model": "gpt-4"
  }'
```

#### 开发者文档:
- **KPMG Code Docs**: https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/
- **Developer Hub**: https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-Developer-and-Product-Manager-Hub.aspx

---

## 🗓️ 时间规划

### 快速路线(已有Prerequisites):
| 天数 | 任务 | 时间 |
|-----|------|------|
| Day 1 | GitHub EMU + PR | 1-2小时 |
| Day 2 | 模块1-5 | 3-4小时 |
| Day 3 | 模块6-9 | 2-3小时 |
| Day 4 | 复习 + 考试 | 2-3小时 |
| **总计** | **4天** | **8-12小时** |

### 完整路线(含Prerequisites):
| 天数 | 任务 | 时间 |
|-----|------|------|
| Day 1-3 | Azure AI-900 | 6-10小时 |
| Day 4 | GitHub Foundations | 4-6小时 |
| Day 5 | GitHub EMU + PR | 1-2小时 |
| Day 6-7 | 9个模块 | 5-7小时 |
| Day 8 | 复习 + 考试 | 2-3小时 |
| **总计** | **8天** | **18-28小时** |

---

## 📞 遇到问题联系

### GitHub EMU:
- **文档**: https://handbook.code.kpmg.com/KPMG-Code/GitHub/Organization%20onboarding/
- **联系**: IT部门

### Workbench学习:
- **Developer Hub**: https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-Developer-and-Product-Manager-Hub.aspx
- **联系**: 页面底部"Contacts"区域

### GLMS访问:
- 联系本地L&D团队
- 提供Program ID: `GX25_CFS_DDF_AI_BLDG_WB_D_PRO`

---

## ✅ 检查清单

### 开始前:
- [ ] 有KPMG邮箱和内网访问权限
- [ ] 能访问SharePoint(spo-global.kpmg.com)
- [ ] 能访问GitHub(github.com/enterprises/kpmg)
- [ ] 已与上司确认这是优先任务

### 提交Assessment前:
- [ ] 完成所有9个模块
- [ ] 做了笔记
- [ ] 理解了核心概念
- [ ] 特别关注了模块2和模块9的部署相关内容

### 获得API Key后:
- [ ] 收到API Key
- [ ] 成功调用一次API
- [ ] 能访问开发者文档

---

## 🎯 下一步:技术测试

获得API Key后,进入实战评估:
- 实验1: API调用效率测试
- 实验2: RAG应用开发测试
- 实验3: UI开发效率测试

参考: `executive_assessment_framework.md`

---

**Good luck!** 🚀 记住:你的目标是"评估",不只是"学会"!
