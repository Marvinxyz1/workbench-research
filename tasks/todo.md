# 任务：HTML 幻灯片转换为 PowerPoint

## 目标
将 `KPMG_Workbench深度概要.md` 的 HTML 幻灯片内容转换为 `.pptx` 格式的 PowerPoint 文件。

## 任务清单

### 1. 环境准备
- [x] 安装 python-pptx 库
- [x] 确认 HTML 文件位置和内容结构

### 2. 开发 PPT 生成脚本
- [x] 创建 `scripts/generate_ppt.py` 脚本
- [x] 定义 KPMG 品牌颜色常量
- [x] 实现幻灯片 1：封面页
- [x] 实现幻灯片 2：执行摘要
- [x] 实现幻灯片 3：战略定位
- [x] 实现幻灯片 4：技术架构
- [x] 实现幻灯片 5：产品生态
- [x] 实现幻灯片 6：总结页

### 3. 测试与优化
- [x] 运行脚本生成 PPT 文件
- [x] 在 PowerPoint 中打开验证效果
- [ ] 调整字体、间距、布局（待用户反馈后调整）

### 4. 文档更新
- [ ] 更新 CLAUDE.md 添加 PPT 生成说明（可选）

## 技术方案

**使用库：** python-pptx

**输出路径：** `generated_docs/KPMG_Workbench深度概要.pptx`

**关键设计：**
- 幻灯片尺寸：16:9 (标准宽屏)
- KPMG 品牌色：
  - KPMG Blue: RGB(0, 51, 141)
  - KPMG Dark Blue: RGB(0, 51, 141)
  - Light Gray: RGB(240, 244, 248)
- 字体：
  - 中文：微软雅黑 (Microsoft YaHei)
  - 英文：Arial

## Review

### 已完成工作

✅ **成功生成 PowerPoint 文件**
- 文件路径: `generated_docs/KPMG_Workbench深度概要.pptx`
- 包含完整的 6 个幻灯片，完全基于原 HTML 内容

### 技术实现亮点

1. **忠实还原 HTML 内容**
   - 所有文本内容完整保留
   - 6 个幻灯片结构完整（封面、执行摘要、战略定位、技术架构、产品生态、总结）
   - 布局逻辑清晰，层次分明

2. **KPMG 品牌视觉规范**
   - 使用官方 KPMG Blue (#00338D) 作为主色调
   - 卡片式设计，包含白色卡片、浅灰背景、彩色强调卡片（红/绿/蓝）
   - 统一页脚样式，包含页码标注

3. **字体与排版**
   - 中文字体: 微软雅黑
   - 标题层级清晰 (44pt/36pt/16pt/14pt)
   - 合理的文本间距和布局

### 文件清单

**新增文件:**
- `scripts/generate_ppt.py` - PPT 生成主脚本 (约 550 行)
- `generated_docs/KPMG_Workbench深度概要.pptx` - 生成的 PowerPoint 文件

**修改文件:**
- `tasks/todo.md` - 任务追踪文档

### 使用说明

**再次生成 PPT:**
```bash
cd scripts
python generate_ppt.py
```

**自定义修改建议:**
- 如需调整颜色：修改 `generate_ppt.py` 中的颜色常量 (第 16-29 行)
- 如需调整布局：修改对应的 `create_slide_X()` 函数
- 如需更换字体：修改 `set_font()` 和 `add_text_box()` 函数中的 `font_name` 参数

### 注意事项

⚠️ **HTML 交互功能无法完全复刻到 PPT:**
- 原 HTML 的点击切换、键盘导航等交互效果在 PPT 中不适用
- PPT 采用静态页面布局，需手动翻页

✅ **保留的核心价值:**
- 所有文本内容
- 视觉层次和信息结构
- KPMG 品牌识别度

### 后续优化方向（可选）

1. 添加 PPT 动画效果
2. 插入图标/图表
3. 调整文本密度
4. 添加备注页

