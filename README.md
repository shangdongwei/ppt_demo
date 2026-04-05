# MD2PPT — Markdown 转 PowerPoint 生成器

> 将 LLM 输出或手写 Markdown 一键渲染为专业 PPT，支持模板继承、表格渲染、自动翻页。

---

## 功能特性

| 功能 | 说明 |
|------|------|
| ✅ Markdown 解析 | 支持标题、列表、表格、代码块、长文本段落 |
| ✅ PPT 模板继承 | 上传 .pptx 模板，自动提取主题色/字体 |
| ✅ 自动翻页 | 每页超过 6 条要点/12 行文本自动分页 |
| ✅ 多种布局 | 封面页、内容页、文本+表格混排、长文本页 |
| ✅ 富样式渲染 | 主题色、表格斑马纹、项目符号、代码字体 |
| ✅ POC Web 界面 | Gradio 驱动，支持文件上传/下载 |

---

## 快速开始

### 1. 安装依赖

```bash
cd md2ppt
pip install -r requirements.txt
```

### 2. 启动 Web 界面

```bash
python app.py
# 访问 http://localhost:7860
```

### 3. 纯 Python 调用

```python
from md_parser import parse_markdown
from ppt_builder import build_pptx

markdown = """
# 我的演示文稿
副标题内容

## 第一章
- 要点一
- 要点二

## 数据对比
| 指标 | Q1 | Q2 |
|------|----|----|
| 营收 | 100万 | 150万 |
"""

slides = parse_markdown(markdown)
pptx_bytes = build_pptx(slides, template_path=None)  # 或传入模板路径

with open("output.pptx", "wb") as f:
    f.write(pptx_bytes)
```

---

## Markdown 格式规范

```markdown
# 演示文稿标题          ← 封面页（H1）
副标题文字              ← 紧跟 H1 的首行文字作为副标题

## 内容章节             ← 新幻灯片（H2）

- 无序列表项           ← 项目符号
- 另一项

1. 有序列表            ← 编号列表
2. 第二项

普通文字段落...         ← 段落文本（自动换行）

| 列A | 列B | 列C |   ← Markdown 表格（与文本混排时自动左右分栏）
|-----|-----|-----|
| 数据 | 数据 | 数据 |

```python
# 代码块               ← 代码渲染（Consolas字体）
print("Hello")
```
```

---

## 项目结构

```
md2ppt/
├── app.py          # Gradio Web 界面
├── md_parser.py    # Markdown → 结构化数据
├── ppt_builder.py  # 结构化数据 → PPTX
├── requirements.txt
└── README.md
```

---

## 核心模块说明

### md_parser.py

- `parse_markdown(text)` → `list[SlideDict]`
- 按 H1/H2 标题自动分割幻灯片
- 解析 bullets、有序列表、表格、代码块、文本段落
- `_auto_paginate()` 实现自动翻页（超出阈值拆分为多页）

### ppt_builder.py

- `build_pptx(slides, template_path=None)` → `bytes`
- 支持 .pptx 模板加载，提取主题色
- `_add_title_slide()` — 深色背景封面，带装饰色条
- `_add_content_slide()` — 浅色背景内容页，顶部色带
- `_render_table()` — 带斑马纹、表头高亮的表格

---

## 进阶：接入 LLM

```python
import anthropic
from md_parser import parse_markdown
from ppt_builder import build_pptx

client = anthropic.Anthropic()

def generate_ppt_from_topic(topic: str) -> bytes:
    prompt = f"""请生成一份关于"{topic}"的演示文稿内容，使用 Markdown 格式：
    - 用 # 作为封面标题
    - 用 ## 分隔每张幻灯片
    - 适当使用列表、表格
    - 控制在 6-8 页以内
    只输出 Markdown，不要其他说明。"""

    message = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )
    
    md_text = message.content[0].text
    slides = parse_markdown(md_text)
    return build_pptx(slides)
```

---

## 默认主题（Ocean Gradient）

| 用途 | 颜色 |
|------|------|
| 深色背景（封面）| `#065A82` |
| 浅色背景（内容）| `#F0F7FF` |
| 强调色 | `#02C39A` |
| 表格表头 | `#065A82` |
| 正文文字 | `#1A1A2E` |

上传 .pptx 模板时，会自动从模板的色彩主题中提取 dk1（深色1）和 accent1 替换默认配色。
