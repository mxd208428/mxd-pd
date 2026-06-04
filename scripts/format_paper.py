#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""论文排版脚本"""

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import copy
import re

# 排版规范配置
CONFIG = {
    "heading1": {
        "font_cn": "黑体",
        "font_en": "Times New Roman",
        "size": Pt(15),  # 小三
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.LEFT,
        "space_before": Pt(0),
        "space_after": Pt(0),
        "line_spacing": Pt(24),
    },
    "heading2": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": Pt(14),  # 四号
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.LEFT,
        "space_before": Pt(0),
        "space_after": Pt(0),
        "line_spacing": Pt(24),
    },
    "heading3": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": Pt(12),  # 小四
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.LEFT,
        "space_before": Pt(0),
        "space_after": Pt(0),
        "line_spacing": Pt(24),
    },
    "body": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": Pt(12),  # 小四
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "space_before": Pt(0),
        "space_after": Pt(0),
        "line_spacing": Pt(24),
        "first_line_indent": Cm(0.74),  # 首行缩进2字符
    },
    "table": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": Pt(10.5),  # 五号
        "alignment": WD_ALIGN_PARAGRAPH.CENTER,
        "line_spacing": Pt(24),
    },
    "figure_caption": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": Pt(10.5),  # 五号
        "alignment": WD_ALIGN_PARAGRAPH.CENTER,
        "line_spacing": Pt(24),
    },
    "abstract_title": {
        "font_cn": "黑体",
        "font_en": "Times New Roman",
        "size": Pt(14),  # 四号
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.CENTER,
        "space_before": Pt(0),
        "space_after": Pt(0),
        "line_spacing": Pt(24),
    },
    "abstract_body": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": Pt(12),  # 小四
        "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "space_before": Pt(0),
        "space_after": Pt(0),
        "line_spacing": Pt(24),
        "first_line_indent": Cm(0.74),
    },
    "reference": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": Pt(10.5),  # 五号
        "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "space_before": Pt(0),
        "space_after": Pt(0),
        "line_spacing": Pt(24),
    },
}


def set_run_font(run, font_cn, font_en, size, bold=False):
    """设置run的字体"""
    # 设置英文字体
    run.font.name = font_en
    run.font.size = size
    run.bold = bold

    # 设置中文字体
    r = run._element
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = r.makeelement(qn('w:rPr'), {})
        r.insert(0, rPr)

    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.insert(0, rFonts)

    rFonts.set(qn('w:eastAsia'), font_cn)
    rFonts.set(qn('w:ascii'), font_en)
    rFonts.set(qn('w:hAnsi'), font_en)


def set_paragraph_format(paragraph, config):
    """设置段落格式"""
    pf = paragraph.paragraph_format
    pf.space_before = config.get("space_before", Pt(0))
    pf.space_after = config.get("space_after", Pt(0))
    pf.line_spacing = config.get("line_spacing", Pt(24))
    pf.alignment = config.get("alignment", WD_ALIGN_PARAGRAPH.JUSTIFY)

    first_indent = config.get("first_line_indent")
    if first_indent:
        pf.first_line_indent = first_indent
    else:
        pf.first_line_indent = Cm(0)


def format_paragraph(paragraph, config):
    """格式化整个段落"""
    set_paragraph_format(paragraph, config)

    for run in paragraph.runs:
        set_run_font(
            run,
            config["font_cn"],
            config["font_en"],
            config["size"],
            config.get("bold", False)
        )


def format_table(table, config):
    """格式化表格"""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.line_spacing = config["line_spacing"]
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.alignment = config["alignment"]

                for run in paragraph.runs:
                    set_run_font(
                        run,
                        config["font_cn"],
                        config["font_en"],
                        config["size"],
                    )


def is_figure_caption(text):
    """判断是否是图注"""
    patterns = [
        r'^图\s*\d+[-\.]\d+',
        r'^Figure\s*\d+',
        r'^Fig\.\s*\d+',
    ]
    for pattern in patterns:
        if re.match(pattern, text.strip()):
            return True
    return False


def is_table_caption(text):
    """判断是否是表注"""
    patterns = [
        r'^表\s*\d+[-\.]\d+',
        r'^Table\s*\d+',
    ]
    for pattern in patterns:
        if re.match(pattern, text.strip()):
            return True
    return False


def is_reference(text):
    """判断是否是参考文献条目"""
    patterns = [
        r'^\[\d+\]',
        r'^\d+\.\s+[A-Z]',  # 1. Author...
    ]
    for pattern in patterns:
        if re.match(pattern, text.strip()):
            return True
    return False


def format_document(input_path, output_path):
    """主排版函数"""
    doc = Document(input_path)

    # 跟踪是否在参考文献部分
    in_references = False

    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()

        # 跳过空段落
        if not text:
            continue

        # 检查样式
        style_name = paragraph.style.name if paragraph.style else ""

        # 标题页元素（段落1-6左右）
        if i < 7 and style_name == "Normal":
            # 标题页内容，保持原样居中
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                if "深度学习" in text or "作业" in text:
                    set_run_font(run, "黑体", "Times New Roman", Pt(22), True)
                elif "提交时间" in text:
                    set_run_font(run, "宋体", "Times New Roman", Pt(14))
                else:
                    set_run_font(run, "宋体", "Times New Roman", Pt(14))
            continue

        # 一级标题
        if style_name == "Heading 1" or style_name.startswith("Heading 1"):
            format_paragraph(paragraph, CONFIG["heading1"])

            # 检查是否进入参考文献部分
            if "参考文献" in text:
                in_references = True

            continue

        # 二级标题
        if style_name == "Heading 2" or style_name.startswith("Heading 2"):
            format_paragraph(paragraph, CONFIG["heading2"])
            continue

        # 三级标题
        if style_name == "Heading 3" or style_name.startswith("Heading 3"):
            format_paragraph(paragraph, CONFIG["heading3"])
            continue

        # 参考文献条目
        if in_references and is_reference(text):
            format_paragraph(paragraph, CONFIG["reference"])
            paragraph.paragraph_format.first_line_indent = Cm(0)
            continue

        # 图注
        if is_figure_caption(text):
            format_paragraph(paragraph, CONFIG["figure_caption"])
            continue

        # 表注（通常在表格内，但以防万一）
        if is_table_caption(text):
            format_paragraph(paragraph, CONFIG["figure_caption"])
            continue

        # 摘要标题
        if "摘要" in text and len(text) < 10:
            format_paragraph(paragraph, CONFIG["abstract_title"])
            continue

        # 普通正文
        format_paragraph(paragraph, CONFIG["body"])

    # 格式化表格
    for table in doc.tables:
        format_table(table, CONFIG["table"])

    # 保存
    doc.save(output_path)
    print(f"排版完成！已保存到: {output_path}")


if __name__ == "__main__":
    input_file = "13-毛小东-深度学习-第二次实验作业-实验报告-排版.docx"
    output_file = "13-毛小东-深度学习-第二次实验作业-实验报告-已排版.docx"
    format_document(input_file, output_file)
