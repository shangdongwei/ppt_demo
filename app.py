"""
app.py
Gradio POC 演示界面
运行: python app.py
"""

import tempfile
import os
import gradio as gr

from md_parser import parse_markdown
from ppt_builder import build_pptx

# ─────────────────────────────────────────────
# 默认示例 Markdown
# ─────────────────────────────────────────────
DEFAULT_MD = """# AI 驱动的产品战略
2024 年度核心报告

## 执行摘要

本报告概述了公司在人工智能领域的战略布局、核心技术优势以及未来18个月的发展路线图。

## 核心业务指标

- **营收增长**: 同比提升 42%，达到 12.8 亿元
- **用户规模**: 活跃用户突破 500 万，留存率 87%
- **技术专利**: 新增 AI 相关专利 38 项
- **团队规模**: 研发人员扩充至 1200 人
- **客户满意度**: NPS 指数 72（行业平均 54）
- **市场份额**: 细分赛道占比从 18% 提升至 31%

## 技术架构对比

| 模块 | 当前方案 | 目标方案 | 预期收益 |
|------|---------|---------|---------|
| 推荐引擎 | 协同过滤 | 大模型增强 | CTR +25% |
| 搜索系统 | 倒排索引 | 向量检索 | 召回率 +40% |
| 风控模型 | 规则引擎 | 实时 ML | 误报 -60% |
| 数据管道 | 批处理 T+1 | 实时流处理 | 延迟 <100ms |

## 产品路线图

- Q1: 完成大模型微调基础设施建设
- Q2: 智能客服 2.0 上线，支持多轮对话
- Q3: 个性化推荐系统全量切换至向量召回
- Q4: 发布开放平台 API，拓展生态合作伙伴

## 竞争格局分析

当前市场处于快速整合阶段，头部效应逐步显现。我们在以下三个维度建立了差异化竞争壁垒：数据飞轮效应显著（用户数据积累超过竞争对手 3 倍）、算法团队核心成员来自顶级实验室、垂直场景的深度定制能力业界领先。

## 风险与应对

- **数据安全**: 已通过 ISO 27001 认证，启动等保三级建设
- **监管变化**: 专职合规团队 8 人，持续跟踪政策动向
- **人才流失**: 核心成员股权激励覆盖率 95%
- **技术替代**: 每季度开展技术雷达扫描，快速响应新范式

## 结论与展望

综合以上分析，公司已形成技术、数据、场景三位一体的护城护体系。建议董事会批准 2024 年 AI 基础设施追加投资 2 亿元，预计 24 个月内实现正向 ROI。
"""

# ─────────────────────────────────────────────
# 核心处理函数
# ─────────────────────────────────────────────

def generate_pptx(markdown_text: str, template_file):
    """Gradio 回调：解析 Markdown → 生成 PPTX → 返回下载路径。"""
    if not markdown_text or not markdown_text.strip():
        return None, "❌ 请输入 Markdown 内容"

    try:
        # 1. 解析 Markdown
        slides = parse_markdown(markdown_text)
        if not slides:
            return None, "❌ 未能从 Markdown 中解析出任何幻灯片，请检查格式"

        # 2. 模板路径
        template_path = template_file.name if template_file else None

        # 3. 生成 PPTX
        pptx_bytes = build_pptx(slides, template_path)

        # 4. 写入临时文件
        tmp = tempfile.NamedTemporaryFile(
            delete=False, suffix=".pptx", prefix="md2ppt_"
        )
        tmp.write(pptx_bytes)
        tmp.flush()
        tmp.close()

        # 5. 生成摘要信息
        slide_types = {}
        for s in slides:
            t = s.get("type", "content")
            slide_types[t] = slide_types.get(t, 0) + 1

        info_lines = [
            f"✅ **生成成功！共 {len(slides)} 张幻灯片**",
            "",
            "📊 幻灯片类型分布：",
        ]
        type_labels = {
            "title": "封面页",
            "content": "内容页",
            "long_text": "长文本页",
        }
        for t, cnt in slide_types.items():
            info_lines.append(f"  - {type_labels.get(t, t)}: {cnt} 张")

        info_lines += [
            "",
            "📋 幻灯片列表：",
        ]
        for i, s in enumerate(slides, 1):
            title = s.get("title", "(无标题)")
            blocks = s.get("content", [])
            block_desc = []
            for b in blocks:
                if b["kind"] == "bullets":
                    block_desc.append(f"{len(b['items'])}条要点")
                elif b["kind"] == "table":
                    block_desc.append(f"表格({len(b.get('rows',[]))}行)")
                elif b["kind"] == "text":
                    block_desc.append("文本段落")
            desc = "、".join(block_desc) if block_desc else "封面"
            info_lines.append(f"  **{i}.** {title} — {desc}")

        info_text = "\n".join(info_lines)
        return tmp.name, info_text

    except Exception as e:
        import traceback
        return None, f"❌ 生成失败：{str(e)}\n\n```\n{traceback.format_exc()}\n```"


# ─────────────────────────────────────────────
# Gradio UI
# ─────────────────────────────────────────────

CSS = """
:root {
    --primary: #065A82;
    --accent: #02C39A;
    --bg: #F0F7FF;
    --surface: #FFFFFF;
}

body {
    background: var(--bg) !important;
}

.gradio-container {
    max-width: 1400px !important;
    margin: 0 auto !important;
    font-family: 'Segoe UI', system-ui, sans-serif !important;
}

.header-block {
    background: linear-gradient(135deg, #065A82 0%, #021A2A 100%);
    border-radius: 16px;
    padding: 32px 40px;
    margin-bottom: 24px;
    color: white;
}

.header-title {
    font-size: 2rem;
    font-weight: 700;
    color: white !important;
    margin: 0;
    letter-spacing: -0.5px;
}

.header-subtitle {
    color: #A0D8EF;
    margin-top: 6px;
    font-size: 1rem;
}

.generate-btn {
    background: linear-gradient(135deg, #02C39A, #065A82) !important;
    color: white !important;
    font-size: 1.1rem !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    padding: 14px 0 !important;
    border: none !important;
    cursor: pointer !important;
    transition: opacity 0.2s !important;
}

.generate-btn:hover {
    opacity: 0.9 !important;
}

footer {
    display: none !important;
}
"""

HEADER_HTML = """
<div class="header-block">
    <h1 class="header-title">🎨 Markdown → PPT 生成器</h1>
    <p class="header-subtitle">
        输入 Markdown 格式内容，自动解析并生成专业 PowerPoint 演示文稿 · 支持上传 .pptx 模板
    </p>
</div>
"""

TIPS_MD = """
### 📝 Markdown 格式说明

| 标记 | 效果 |
|------|------|
| `# 标题` | 生成封面页，支持副标题 |
| `## 章节` | 生成内容页 |
| `- 要点` / `* 要点` | 项目符号列表 |
| `1. 条目` | 有序编号列表 |
| `\| 表头 \|` | Markdown 表格（自动美化） |
| 普通段落文字 | 长文本自动排版 |

**超出每页容量时自动翻页，无需手动分页！**
"""


def build_ui():
    with gr.Blocks(title="MD → PPT Generator") as demo:

        gr.HTML(HEADER_HTML)

        with gr.Row(equal_height=False):
            # ── 左栏：输入 ──────────────────────
            with gr.Column(scale=6):
                md_input = gr.Textbox(
                    label="📄 Markdown 内容",
                    placeholder="在此输入或粘贴 Markdown 文本...\n\n# 演示标题\n副标题文字\n\n## 第一章\n- 要点一\n- 要点二",
                    lines=22,
                    max_lines=40,
                    value=DEFAULT_MD,
                )

                with gr.Row():
                    template_file = gr.File(
                        label="📁 上传 PPT 模板（可选，.pptx）",
                        file_types=[".pptx"],
                        scale=4,
                    )
                    with gr.Column(scale=2):
                        gr.Markdown("**提示**：上传已有 .pptx 模板可继承其主题色")

                generate_btn = gr.Button(
                    "🚀 生成 PPT",
                    variant="primary",
                    elem_classes=["generate-btn"],
                )

            # ── 右栏：输出 ──────────────────────
            with gr.Column(scale=4):
                gr.Markdown(TIPS_MD)

                output_file = gr.File(
                    label="⬇️ 下载生成的 PPT",
                    interactive=False,
                )
                status_md = gr.Markdown(
                    value="👈 填写内容后点击【生成 PPT】按钮",
                    label="生成状态",
                )

        # 示例快速填充
        gr.Examples(
            examples=[
                ["# 项目启动报告\n初版 · 2024年\n\n## 项目背景\n\n- 市场需求持续增长，竞争格局加剧\n- 现有产品功能无法满足客户新需求\n- 技术债务积累，研发效率下降 30%\n\n## 解决方案\n\n- 全面重构核心引擎，采用微服务架构\n- 引入 AI 辅助功能，提升用户体验\n- 建立 DevOps 流水线，实现持续交付\n\n## 资源投入\n\n| 资源 | 数量 | 周期 |\n|------|------|------|\n| 研发工程师 | 12人 | 6个月 |\n| 产品经理 | 2人 | 6个月 |\n| 测试工程师 | 4人 | 4个月 |\n| 预算 | 300万 | 全周期 |"],
                ["# 季度业务复盘\nQ3 2024\n\n## 关键成果\n\n1. 成功上线新版用户系统\n2. 完成海外市场调研报告\n3. 获得 A 轮融资 5000 万\n4. 团队规模扩张至 80 人\n\n## 数据看板\n\n| 指标 | Q2 | Q3 | 环比 |\n|------|-----|-----|------|\n| DAU | 12万 | 18万 | +50% |\n| 收入 | 800万 | 1100万 | +37% |\n| 毛利率 | 62% | 68% | +6pp |\n\n## 下季计划\n\n- 国际版 Beta 测试启动\n- 完成 B 轮融资准备材料\n- 核心功能完成 AI 改造"],
            ],
            inputs=[md_input],
            label="💡 快速示例",
        )

        # 绑定按钮
        generate_btn.click(
            fn=generate_pptx,
            inputs=[md_input, template_file],
            outputs=[output_file, status_md],
            show_progress="full",
        )

    return demo


if __name__ == "__main__":
    app = build_ui()
    app.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=False,
        show_error=True,
        css=CSS,
    )