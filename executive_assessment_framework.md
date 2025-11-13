# KPMG Workbench 战略评估框架(精简版)
## 咨询视角:给上司的决策依据

> **背景**: 上司和Fukushima-san期望评估KPMG Workbench是否能让团队"开发しやすくなる"(开发更轻松),你被专任负责评估。

> **核心问题**: Workbench的主要定位是"社内向けサービスの開発環境"(内部服务开发环境),但可以作为"デモ環境"使用。**关键待验证问题:**用Workbench开发的应用能否直接交付客户?如果不能,迁移成本有多高?

---

## 一、你需要向上司汇报的5个核心观点

### 观点1: 准入门槛与团队适配性
**论证角度:** 在"現状メンバーの空がない状態"(团队很忙)的情况下,ROI是否合理?

**需要证明的数据:**
- 你完成认证花了X小时(包含:Prerequisites + Developer Learning Path + Assessment)
  - **Developer Learning Path包含9个模块**,总时长约318分钟(~5.3小时纯视频时间)
  - 实际学习时间需要加上实践、作业、考试时间
- 预计其他成员需要Y小时,考虑技术背景差异
- 团队中哪些人适合/不适合使用Workbench
- **决策建议:** 是否值得让全员学习,还是只培养2-3个Champion?

**完成后你会获得:**
1. ✅ **KPMG Workbench Knowledge Badge**(数字徽章)
2. ✅ **API Key**(核心!) - 访问Workbench平台的密钥
3. ✅ 访问以下技术能力:
   - Inference API + Completion API(调用AI模型)
   - RAG服务(检索增强生成)
   - Design Systems(UI组件库)
   - Feature Flags(功能定制)

**关键信息来源:**
- Developer Learning Path页面(9个模块详细内容)
- Prerequisites要求(GitHub EMU + 2个推荐认证:Azure Fundamentals/AI Fundamentals/GitHub Foundations等)
- Tech Talks系列(2025年4-6月)的实用性评估

---

### 观点2: Demo开发效率提升(可量化!)
**论证角度:** 作为"デモ環境",它是否显著加速售前PoC开发?

**需要证明的数据:**
- 开发一个典型Demo(如RAG Chatbot/Audit Agent)的时间对比
  - 现有流程:X周
  - Workbench:Y天
  - 提速比例:Z%
- Demo质量对比(界面、功能完整度、客户反馈)
- **核心指标:** 如果提速<30%,不值得切换;如果>50%,强烈推荐

**测试任务建议(3个关键对比实验):**

#### 实验1: API调用效率测试
**对比:** Workbench API vs 团队现有方案
- **测试A:** 用Workbench的**Completion API**调用GPT模型
- **测试B:** 用团队现有的Azure OpenAI直接调用
- **对比指标:**
  - 配置时间(从零到首次调用成功)
  - API响应速度
  - 配额限制与成本
  - 代码复杂度(需要写多少行代码)
- **预期结果:** 如果Workbench配置时间<30分钟 且代码量减少>50%,则有价值

#### 实验2: RAG应用开发测试
**任务:** 开发一个"审计知识库Chatbot"
- **方案A(Workbench):** 用Workbench的**RAG服务**(Ingestion + Retrieval + AI Search)
- **方案B(现有方案):** 用LangChain + Azure AI Search + 自己搭建向量数据库
- **对比指标:**
  - 从零到可演示的时间(小时)
  - 检索准确率
  - 代码维护成本
- **预期结果:** 如果Workbench开发时间<现有方案50%,则强烈推荐

#### 实验3: UI开发效率测试
**任务:** 构建一个标准的AI对话界面
- **方案A(Workbench):** 用**Design Systems组件库**
- **方案B(现有方案):** 从零用React/Vue写组件
- **对比指标:**
  - UI开发时间
  - 组件复用性
  - 是否符合KPMG品牌规范
- **预期结果:** 如果组件库能节省>60% UI开发时间,且符合品牌要求,则有价值

**关键测试要点:**
- 记录每个步骤的时间(精确到小时)
- 截图保存开发过程(作为汇报证据)
- 对比代码行数(Workbench vs 现有方案)

---

### 观点3: 客户交付能力与迁移成本(关键风险点!)
**论证角度:** Workbench开发的应用能否直接部署到客户生产环境?

**核心待验证问题:**
1. **Workbench是"开发平台"还是"托管平台"?**
   - 如果是开发平台:代码可以导出,部署到客户的Azure/AWS环境 → 迁移成本低
   - 如果是托管平台:应用必须运行在KPMG的Workbench上 → **无法交付给客户!**

2. **数据主权问题:**
   - 官方强调"Global data sovereignty"(US/Europe/Australia数据中心)
   - 但日本客户可能要求数据存储在日本境内
   - **关键问题:** Workbench是否支持日本数据中心?如果不支持,能否迁移到客户自己的环境?

3. **迁移成本分析(如果不能直接交付):**
   - 从Workbench迁移到客户生产环境的工作量(代码重写率、数据迁移、安全认证)
   - 迁移成本占总开发时间的比例
     - 如果>50%:Workbench只是"Demo工具",不适合实战
     - 如果20-50%:可用于快速原型,但需考虑迁移成本
     - 如果<20%:可以作为标准开发平台

**必须搞清楚的技术问题:**
- [ ] Workbench开发的代码能否导出为标准Python/Node.js项目?
- [ ] RAG服务是否可以迁移到客户的Azure AI Search?
- [ ] Completion API调用能否切换到客户自己的OpenAI/Azure账号?
- [ ] 数据是否锁定在Workbench平台,还是可以导出?
- [ ] 是否支持Docker容器化部署?(如果支持,迁移就容易)

**关键调研:**
- Secret Management最佳实践(客户环境如何管理密钥)
- SDLC(Software Development Lifecycle)与客户项目流程的差异
- GitHub EMU的权限管理
- Workbench的部署模式(SaaS? PaaS? 还是可导出代码?)

---

### 观点4: Agentic AI能力(核心竞争力!)
**论证角度:** Workbench的核心不是"开发环境",而是"AI Agent平台"。这是否符合团队战略方向?

**需要证明的数据:**
- Workbench的Agent开发能力vs现有方案(如LangChain/AutoGPT/CrewAI)
- 是否有预置Agent模板/框架?
- 与KPMG知识库(审计标准、行业法规)的集成难度
- **战略价值:** 如果Fukushima-san看重Agentic AI,这是最大卖点

**参考资料:**
- [The agentic AI advantage白皮书](https://kpmg.com/us/en/articles/2025/the-agentic-ai-advantage.html)
- Microsoft Keynote - Agentic AI Thinking(Developers Conference录像)
- Design Systems for KPMG Workbench

---

### 观点5: 战略契合度与团队影响力
**论证角度:** 使用Workbench能否提升团队在KPMG Global的地位和资源获取能力?

**需要证明的数据:**
- 成为Workbench"Early Adopters"的政治收益
  - 能否参与Global AI Ninjas/Navigators社区?
  - 能否争取更多预算/资源?
- 与KPMG AI愿景的契合度
  - Workbench被定位为"AI Backbone"(中央管理的AI创新平台)
  - 我们的项目能否成为Champions案例?
- **软性价值:** 团队品牌提升、职业发展机会、跨Member Firm协作

**关键资源:**
- [KPMG Global aIQ Hub](https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-AI)
- [Trusted AI principles](https://spo-global.kpmg.com/sites/go-oi-bus-People/SitePages/Trusted-AI.aspx)
- Leaders and Enablers Hub

---

## 二、给上司的决策框架(Go/No-Go标准)

### ✅ 推荐全面推广的条件:
1. Demo开发速度提升 > 50%
2. 迁移成本 < 20%总开发时间
3. 团队认证时间 < 20小时/人
4. Agentic AI能力显著优于现有方案
5. 获得Fukushima-san明确的战略支持

### ⚠️ 建议小范围试点的条件:
1. Demo开发速度提升 30%-50%
2. 迁移成本 20%-40%
3. 只有2-3个核心成员能快速上手
4. 技术栈与现有方案有一定兼容性

### ❌ 不推荐投入的条件:
1. Demo开发速度提升 < 30%
2. 迁移成本 > 50%
3. 学习成本 > 1个月全职
4. 数据隔离机制不清晰/有合规风险
5. 团队现有方案(如Azure AI Studio)性价比更高

---

## 三、分阶段评估计划(4周)

| 周 | 任务 | 具体测试内容 | 交付物 | 汇报要点 |
|----|------|------------|--------|---------|
| **W1** | 完成Prerequisites + Developer Learning Path(9个模块) | - 记录每个模块学习时间<br>- 评估模块实用性(1-5分)<br>- 完成Assessment获得API Key | - Digital Badge<br>- API Key<br>- 学习时间明细表 | - 总学习时间:X小时<br>- 最有价值模块:RAG/Design Systems<br>- 可跳过模块:[具体名称] |
| **W2** | 实验1+2:API调用+RAG开发测试 | **实验1(API测试):**<br>- Workbench Completion API vs Azure OpenAI<br>- 记录配置时间、响应速度<br>**实验2(RAG测试):**<br>- 用RAG服务开发知识库Chatbot<br>- 对比LangChain方案 | - 2个技术对比报告<br>- 可演示的RAG Demo<br>- 代码行数对比数据 | - API配置时间:Workbench X分钟 vs 传统Y分钟<br>- RAG开发时间:减少Z%<br>- 代码复杂度降低[%] |
| **W3** | 实验3:UI开发+迁移测试 | **实验3(UI测试):**<br>- 用Design Systems构建界面<br>- 对比从零写React的时间<br>**迁移测试:**<br>- 尝试将Demo迁移到客户环境<br>- 评估代码重写率 | - UI开发时间对比报告<br>- 迁移成本评估(代码重写率、工作量) | - UI开发节省X%时间<br>- 迁移成本占总开发Y%<br>- 数据隔离机制[清晰/存在问题] |
| **W4** | 风险分析、战略价值评估、撰写报告 | - 汇总3个实验数据<br>- 计算ROI<br>- 评估战略契合度<br>- 撰写Executive Summary | - 完整评估报告(20页)<br>- Executive Summary(1页)<br>- Go/No-Go建议 | - 综合效率提升:[%]<br>- ROI分析<br>- [✅ Go / ⚠️ 试点 / ❌ No-Go]<br>- 分阶段推广计划 |

---

## 四、给上司的汇报模板(Executive Summary)

```
[上司名字] 您好,

我已完成AI Workbench的评估,以下是核心结论:

【学习成本】
- 我用了[X]小时完成认证,预计其他成员需要[Y]小时
- 建议在[时间段]让[N]个核心成员先完成培训

【Demo开发效率】
- 开发[具体Demo名称]的时间从[X周]缩短到[Y天],提速[Z]%
- [是否/不]显著提升售前竞争力

【迁移成本】
- 从Workbench迁移到客户环境需要[X]%额外工作
- 数据隔离机制[已通过/存在风险]

【Agentic AI能力】(重点!)
- Workbench在[具体功能]上[优于/不如]我们现有方案
- [是否]符合Fukushima-san的AI战略方向

【战略价值】
- 使用Workbench[能/不能]提升团队在Global的影响力
- [是否]有助于争取更多资源

【我的建议】
- [✅ 推荐全面推广 / ⚠️ 建议小范围试点 / ❌ 不推荐投入]
- 理由:[核心数据支撑]
- 下一步行动:[具体计划]

附件:详细评估报告([X]页)

[你的名字]
SC, AI Development
[日期]
```

---

## 五、关键资源速查表

### 必读文档:
- [Developer Learning Path](https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-learning-and-development-development-track.aspx)
- [KPMG Workbench User Guide](https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/)
- [Design Systems](https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/Design-Systems-for-KPMG-Workbench.aspx)

### 必看视频:
- Tech Talks系列(2025年4-6月):Core Platform Overview, Demonstration, Design System
- Developers Conference(2024年11月):Microsoft Keynote(Agentic AI)

### 战略资源:
- [The Agentic AI Advantage白皮书](https://kpmg.com/us/en/articles/2025/the-agentic-ai-advantage.html)
- [Global AI Ninjas and Navigators](https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-AI/SitePages/Global-AI-Ninjas-Navigators.aspx)

---

## 六、关键差异:你的评估 vs 普通培训

| 普通培训 | 你的战略评估 |
|---------|------------|
| "我完成了认证" | "认证需要X小时,其中[模块]最有价值,[模块]可跳过" |
| "Workbench很好用" | "Demo开发速度提升X%,但迁移成本占Y%,ROI是Z" |
| "学到了很多" | "Agentic AI能力[优于/不如]现有方案,建议[采用/不采用]" |
| (没有行动建议) | "基于数据,建议[Go/No-Go],分阶段推广计划如下..." |

**记住:** 你的价值不是"完成培训",而是"用咨询思维给出决策依据"。保持批判性,用数据说话,不要被营销话术迷惑!

---

## 七、Workbench核心技术能力清单(获得API Key后可用)

### 你完成学习后会拿到什么?
1. **KPMG Workbench Knowledge Badge**(数字徽章) - 证明你完成了培训
2. **API Key**(关键!) - 这是访问Workbench平台的密钥,没有它无法使用任何功能

### 有了API Key后可以使用的核心服务:

#### 1. AI模型调用服务
- **Inference API** - 推理API,用于AI模型调用
- **Completion API** - 文本生成API(类似OpenAI的Chat Completion接口)
- **支持的模型:** GPT-4, GPT-3.5等(具体需要在学习中确认)
- **关键测试:** 对比Workbench API vs 直接用Azure OpenAI的配置复杂度和响应速度

#### 2. RAG(检索增强生成)服务 - 重点!
Workbench提供完整的RAG服务链:
- **Data Transfer Service** - 数据传输服务
- **Ingestion Orchestrator** - 数据摄取编排器
- **Chunking Service** - 文本分块服务
- **Extraction Service** - 信息提取服务
- **AI Search Indexing Service** - AI搜索索引服务
- **Retrieval Service** - 检索服务
- **关键测试:** 用RAG服务开发知识库Chatbot vs 用LangChain自己搭建的时间差异

#### 3. Design Systems(UI组件库)
- 符合KPMG品牌规范的UI组件
- AI对话界面组件
- 可能包括:Chatbot界面、文件上传、结果展示等
- **关键测试:** 用组件库构建界面 vs 从零写React的时间节省比例

#### 4. Feature Flags(功能定制)
- 根据Member Firm需求启用/禁用特定功能
- 支持区域合规性定制
- **关键测试:** 评估这对日本KC的适配性

#### 5. Developer Resources
- 开发者文档、API文档
- 示例代码
- Troubleshooting指南
- **关键测试:** 文档质量是否足够支持独立开发

### 重点评估问题清单:

#### 关于API Key和访问权限:
- [ ] API Key有配额限制吗?(每月调用次数/Token数)
- [ ] 是否有成本?(按使用量收费还是免费?)
- [ ] API Key是个人的还是团队共享的?
- [ ] 如果成员离职,API Key如何管理?

#### 关于技术能力:
- [ ] Workbench的Completion API是否与OpenAI API兼容?(代码是否可以无缝切换?)
- [ ] RAG服务支持哪些文档格式?(PDF/Word/Excel/图片?)
- [ ] 向量数据库是什么?(Azure AI Search还是其他?)
- [ ] Design Systems组件库是什么技术栈?(React/Vue/Angular?)

#### 关于数据隔离:
- [ ] 用API Key上传的数据存储在哪里?
- [ ] 是否有多租户隔离机制?
- [ ] 训练/调优模型时,数据会被KPMG Global共享吗?
- [ ] 如何确保Demo数据不泄露到客户项目?

#### 关于迁移成本:
- [ ] 用Workbench开发的应用,能否导出为标准代码?(还是锁定在Workbench平台?)
- [ ] 迁移到客户环境时,哪些服务需要重写?(RAG? API调用?)
- [ ] 是否可以用Docker容器化?(方便迁移到客户的Azure/AWS环境)

---

## 八、第一周学习任务详细清单(获得API Key前)

### Developer Learning Path - 9个模块详细内容

| # | 模块名称 | 时长 | 核心内容 | 你需要记录的 |
|---|---------|------|---------|-------------|
| 1 | Introduction to KPMG Workbench | 54分钟 | 背景、AI战略、Workbench定位 | 实用性评分(1-5分),是否可跳过 |
| 2 | Revolutionizing AI Productivity: Dive into KPMG Workbench | 35分钟 | 平台架构、Azure设计、可用性/性能/合规 | Azure架构复杂度,与现有环境差异 |
| 3 | Deep Dive: Inference API | 26分钟 | 推理API的访问、认证、使用方法 | 与Azure OpenAI差异,代码示例质量 |
| 4 | Deep Dive: Completion API | 28分钟 | Completion API、模型可用性、Payload格式 | API兼容性,是否支持Streaming |
| 5 | RAG: Overview and Building Blocks | 49分钟 | RAG服务的6个组件详细介绍 | 哪些组件最有用,配置复杂度 |
| 6 | RAG: Leading Practices | 53分钟 | RAG最佳实践、用例、关键功能 | 实际可用性,是否比LangChain简单 |
| 7 | Tailoring KPMG Workbench for Global: Feature Flags | 13分钟 | 功能定制、区域合规 | 日本KC能否自定义功能 |
| 8 | Designing AI Experiences with KPMG Workbench | 39分钟 | Design Systems、组件库、案例研究 | 组件质量,是否符合KPMG品牌 |
| 9 | Building Better, Faster: Guide to Developer Resources | 21分钟 | 开发者资源、文档、Troubleshooting | 文档完整度,是否足够实战 |

**总计:** 318分钟(~5.3小时纯视频) + 实践/作业/考试时间

### 你在第一周需要完成的具体任务:

#### Day 1-2: Prerequisites
- [ ] 确认GitHub EMU账号已开通
- [ ] 完成2个推荐认证(如果还没有):
  - [ ] Azure Fundamentals AZ-900 **或**
  - [ ] Azure AI Fundamentals AI-900 **或**
  - [ ] GitHub Foundations
- [ ] 提交至少1个Pull Request到GitHub EMU仓库(前置要求)

#### Day 3-5: Developer Learning Path
- [ ] 按顺序观看9个模块(318分钟)
- [ ] 每个模块结束后填写评估表:
  ```
  模块名称: [XXX]
  观看时间: [实际花费X小时,包括暂停、重看]
  实用性评分: [1-5分]
  关键收获: [3句话总结]
  是否可跳过: [是/否]
  ```
- [ ] 记录每个模块中的疑问点(用于后续测试验证)

#### Day 6: Assessment + 获得API Key
- [ ] 完成Knowledge Badge Assessment(考试)
- [ ] 获得Digital Badge
- [ ] **获得API Key**(最关键!)
- [ ] 测试API Key是否可用:
  - [ ] 尝试调用一次Completion API(发送"Hello World")
  - [ ] 记录配置过程和遇到的问题

#### Day 7: 总结第一周
- [ ] 汇总学习时间数据
- [ ] 撰写"第一周学习总结"(给上司的中期汇报):
  ```
  - 总学习时间: X小时(视频5.3h + 实践Yh + 考试Zh)
  - 最有价值模块: [模块5/6 RAG相关]
  - 可跳过模块: [模块1可能太理论]
  - API Key已获得,可进入第二周技术测试
  ```

---

**Good luck!** 记住:第一周的目标不是"学会",而是"评估学习成本"和"获得API Key"!💪
