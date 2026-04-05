"""
md_parser.py
将 Markdown 文本解析为结构化的幻灯片数据列表。

每张幻灯片结构:
{
    "type": "title" | "content" | "table" | "long_text",
    "title": str,
    "subtitle": str | None,       # 仅 title 类型
    "content": list[dict],         # 内容块列表
    "notes": str | None,
}

content 块结构:
{
    "kind": "bullets" | "table" | "text" | "code",
    "items": list[str],            # bullets
    "headers": list[str],          # table
    "rows": list[list[str]],       # table
    "text": str,                   # text / code
    "language": str,               # code
}
"""

import re
from typing import Any

# 每页最多建议的内容行数（超出则自动翻页）
MAX_BULLETS_PER_SLIDE = 6
MAX_TEXT_LINES_PER_SLIDE = 12


def parse_markdown(md_text: str) -> list[dict]:
    """主入口：将 markdown 字符串解析为 slide 列表。"""
    lines = md_text.splitlines()
    raw_slides = _split_into_raw_slides(lines)
    slides = []
    for raw in raw_slides:
        parsed = _parse_raw_slide(raw)
        # 自动翻页处理
        expanded = _auto_paginate(parsed)
        slides.extend(expanded)
    return slides


# ─────────────────────────────────────────────
# 第一步：按 H1 / H2 分割原始幻灯片块
# ─────────────────────────────────────────────

def _split_into_raw_slides(lines: list[str]) -> list[list[str]]:
    """按 H1 (# ) 或 H2 (## ) 分割为原始行组。"""
    slides = []
    current: list[str] = []

    for line in lines:
        if re.match(r'^#{1,2}\s+', line) and current:
            slides.append(current)
            current = [line]
        else:
            current.append(line)

    if current:
        slides.append(current)

    # 过滤空块
    return [s for s in slides if any(l.strip() for l in s)]


# ─────────────────────────────────────────────
# 第二步：解析单个原始幻灯片块
# ─────────────────────────────────────────────

def _parse_raw_slide(lines: list[str]) -> dict:
    """将一组行解析为 slide 字典。"""
    if not lines:
        return {}

    # 确定标题
    title_line = lines[0].strip()
    level = len(re.match(r'^(#+)', title_line).group(1)) if re.match(r'^(#+)', title_line) else 0
    title = re.sub(r'^#+\s*', '', title_line).strip()

    # 第一张 H1 可能是封面（带副标题）
    is_cover = level == 1
    subtitle = None
    body_start = 1

    if is_cover:
        # 副标题：紧跟的非空非标题行
        for i in range(1, len(lines)):
            l = lines[i].strip()
            if l and not l.startswith('#'):
                subtitle = l
                body_start = i + 1
                break
            elif l.startswith('#'):
                break

    body_lines = lines[body_start:]
    content_blocks = _parse_body(body_lines)

    slide_type = "title" if is_cover else "content"

    # 如果内容中只有大段文本则标记为 long_text
    if content_blocks and all(b["kind"] == "text" for b in content_blocks):
        slide_type = "long_text"

    return {
        "type": slide_type,
        "title": title,
        "subtitle": subtitle,
        "content": content_blocks,
        "notes": None,
    }


def _parse_body(lines: list[str]) -> list[dict]:
    """将正文行解析为 content 块列表（支持 bullets、table、text、code）。"""
    blocks = []
    i = 0
    while i < len(lines):
        line = lines[i]

        # 代码块
        if line.strip().startswith('```'):
            lang = line.strip()[3:].strip()
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            blocks.append({"kind": "code", "text": '\n'.join(code_lines), "language": lang})
            i += 1
            continue

        # Markdown 表格
        if '|' in line and i + 1 < len(lines) and re.match(r'^\s*\|?[\s\-|:]+\|', lines[i + 1]):
            table_lines = []
            while i < len(lines) and '|' in lines[i]:
                table_lines.append(lines[i])
                i += 1
            table_block = _parse_table(table_lines)
            if table_block:
                blocks.append(table_block)
            continue

        # 无序列表
        if re.match(r'^\s*[-*+]\s+', line):
            items = []
            while i < len(lines) and re.match(r'^\s*[-*+]\s+', lines[i]):
                item_text = re.sub(r'^\s*[-*+]\s+', '', lines[i]).strip()
                items.append(_clean_inline(item_text))
                i += 1
            blocks.append({"kind": "bullets", "items": items})
            continue

        # 有序列表
        if re.match(r'^\s*\d+\.\s+', line):
            items = []
            while i < len(lines) and re.match(r'^\s*\d+\.\s+', lines[i]):
                item_text = re.sub(r'^\s*\d+\.\s+', '', lines[i]).strip()
                items.append(_clean_inline(item_text))
                i += 1
            blocks.append({"kind": "bullets", "items": items, "ordered": True})
            continue

        # 普通文本段落（跳过空行和 H3+）
        stripped = line.strip()
        if stripped and not stripped.startswith('#'):
            para_lines = []
            while i < len(lines) and lines[i].strip() and not lines[i].strip().startswith('#'):
                if '|' in lines[i]:
                    break
                if re.match(r'^\s*[-*+]\s+', lines[i]):
                    break
                if re.match(r'^\s*\d+\.\s+', lines[i]):
                    break
                para_lines.append(_clean_inline(lines[i].strip()))
                i += 1
            if para_lines:
                blocks.append({"kind": "text", "text": ' '.join(para_lines)})
            continue

        # H3 子标题 → 当作 text 处理
        if stripped.startswith('###'):
            sub_title = re.sub(r'^#+\s*', '', stripped)
            blocks.append({"kind": "text", "text": f"▶ {sub_title}", "bold": True})
            i += 1
            continue

        i += 1

    return blocks


def _parse_table(table_lines: list[str]) -> dict | None:
    """解析 Markdown 表格为结构化 dict。"""
    rows = []
    for line in table_lines:
        if re.match(r'^\s*\|?[\s\-|:]+\|', line):
            continue  # 分隔行
        cells = [c.strip() for c in line.strip().strip('|').split('|')]
        rows.append(cells)
    if len(rows) < 1:
        return None
    return {
        "kind": "table",
        "headers": rows[0],
        "rows": rows[1:],
    }


def _clean_inline(text: str) -> str:
    """清理内联 Markdown 标记（粗体、斜体、代码等）保留文本。"""
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'`(.+?)`', r'\1', text)
    text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)
    return text.strip()


# ─────────────────────────────────────────────
# 第三步：自动翻页
# ─────────────────────────────────────────────

def _auto_paginate(slide: dict) -> list[dict]:
    """将超长幻灯片拆分为多页。"""
    if not slide:
        return []

    pages = [slide]
    result = []

    for page in pages:
        split_pages = _split_slide(page)
        result.extend(split_pages)

    return result


def _split_slide(slide: dict) -> list[dict]:
    """递归拆分单张超长幻灯片。"""
    content = slide.get("content", [])
    if not content:
        return [slide]

    # 统计总条目数
    total_items = 0
    for block in content:
        if block["kind"] == "bullets":
            total_items += len(block["items"])
        elif block["kind"] == "text":
            # 按换行估算
            total_items += max(1, len(block["text"]) // 80)
        elif block["kind"] == "table":
            total_items += len(block.get("rows", [])) + 2
        else:
            total_items += 3

    # 不超限则直接返回
    if total_items <= MAX_BULLETS_PER_SLIDE:
        return [slide]

    # 拆分：尽量按 block 边界分页
    pages = []
    current_blocks = []
    current_count = 0
    page_num = 1
    title = slide["title"]

    def flush(blocks, num):
        label = f" ({num})" if num > 1 else ""
        return {
            "type": slide["type"],
            "title": title + label,
            "subtitle": slide.get("subtitle") if num == 1 else None,
            "content": blocks,
            "notes": slide.get("notes"),
        }

    for block in content:
        block_size = 0
        if block["kind"] == "bullets":
            block_size = len(block["items"])
        elif block["kind"] == "text":
            block_size = max(1, len(block["text"]) // 80)
        elif block["kind"] == "table":
            block_size = len(block.get("rows", [])) + 2
        else:
            block_size = 3

        # 如果单个 bullets 块太长，拆分
        if block["kind"] == "bullets" and block_size > MAX_BULLETS_PER_SLIDE:
            if current_blocks:
                pages.append(flush(current_blocks, page_num))
                page_num += 1
                current_blocks = []
                current_count = 0
            items = block["items"]
            for chunk_start in range(0, len(items), MAX_BULLETS_PER_SLIDE):
                chunk = items[chunk_start:chunk_start + MAX_BULLETS_PER_SLIDE]
                chunk_block = {**block, "items": chunk}
                pages.append(flush([chunk_block], page_num))
                page_num += 1
            continue

        if current_count + block_size > MAX_BULLETS_PER_SLIDE and current_blocks:
            pages.append(flush(current_blocks, page_num))
            page_num += 1
            current_blocks = []
            current_count = 0

        current_blocks.append(block)
        current_count += block_size

    if current_blocks:
        pages.append(flush(current_blocks, page_num))

    return pages if pages else [slide]
