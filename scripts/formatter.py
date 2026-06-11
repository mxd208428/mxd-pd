#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""核心排版逻辑模块"""

import re
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_run_font(run, font_cn, font_en, size, bold=False, ascii_use_cn=False):
    """设置 run 的字体（中英文分别设置）

    Args:
        run: docx Run 对象
        font_cn: 中文字体名
        font_en: 英文字体名
        size: Pt 字号对象
        bold: 是否加粗
        ascii_use_cn: 标题模式，数字和英文也使用中文字体（避免标题数字突兀）
    """
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
    if ascii_use_cn:
        # 标题模式：数字和英文也跟中文字体，避免"1.1"、"1.2.1"等编号突兀
        rFonts.set(qn('w:ascii'), font_cn)
        rFonts.set(qn('w:hAnsi'), font_cn)
    else:
        rFonts.set(qn('w:ascii'), font_en)
        rFonts.set(qn('w:hAnsi'), font_en)


def _set_spacing(paragraph, attr, value):
    """设置段前/段后间距，支持 "1line" 格式（Word原生段前行数）"""
    pPr = paragraph._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = paragraph._element.makeelement(qn('w:pPr'), {})
        paragraph._element.insert(0, pPr)

    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = pPr.makeelement(qn('w:spacing'), {})
        pPr.append(spacing)

    if isinstance(value, str) and value.endswith('line'):
        lines = float(value.replace('line', ''))
        spacing.set(qn(f'w:{attr}Lines'), str(int(lines * 100)))
        # 清除旧的绝对间距，避免和 beforeLines 冲突
        old_attr = qn(f'w:{attr}')
        if old_attr in spacing.attrib:
            del spacing.attrib[old_attr]
    elif value:
        spacing.set(qn(f'w:{attr}'), str(Pt(value).emu))


def set_paragraph_format(paragraph, config):
    """设置段落格式

    Args:
        paragraph: docx Paragraph 对象
        config: 段落配置字典
    """
    pf = paragraph.paragraph_format

    # 段前/段后间距：支持 "1line" 格式
    sb = config.get("space_before", 0)
    sa = config.get("space_after", 0)
    if isinstance(sb, str) and sb.endswith('line'):
        _set_spacing(paragraph, 'before', sb)
    else:
        pf.space_before = Pt(sb) if sb else Pt(0)

    if isinstance(sa, str) and sa.endswith('line'):
        _set_spacing(paragraph, 'after', sa)
    else:
        pf.space_after = Pt(sa) if sa else Pt(0)

    pf.line_spacing = config.get("line_spacing", Pt(24))

    # 对齐方式：支持字符串和 WD_ALIGN_PARAGRAPH 枚举
    align = config.get("alignment", WD_ALIGN_PARAGRAPH.JUSTIFY)
    ALIGN_MAP = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    if isinstance(align, str):
        align = ALIGN_MAP.get(align.lower(), WD_ALIGN_PARAGRAPH.JUSTIFY)
    pf.alignment = align

    first_indent = config.get("first_line_indent")
    if first_indent is not None:
        # 支持 "2char" 格式（首行缩进2字符，相对于字号自动计算）
        if isinstance(first_indent, str) and first_indent.endswith("char"):
            chars = int(first_indent.replace("char", ""))
            target_val = str(chars * 100)
            pPr = paragraph._element.find(qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                paragraph._element.insert(0, pPr)
            ind = pPr.find(qn('w:ind'))
            # 检查当前值：已经是目标值则跳过，已经是空（无缩进）也跳过
            if ind is not None:
                current_chars = ind.get(qn('w:firstLineChars'))
                current_fixed = ind.get(qn('w:firstLine'))
                if current_chars == target_val:
                    return  # 已经是目标值
                if current_chars is None and (current_fixed is None or current_fixed == '0'):
                    return  # 已经是无缩进，不覆盖
            if ind is None:
                ind = OxmlElement('w:ind')
                pPr.append(ind)
            ind.set(qn('w:firstLineChars'), target_val)
            # 移除固定 firstLine 值，让 Word 按字号自动计算
            if ind.get(qn('w:firstLine')) is not None:
                del ind.attrib[qn('w:firstLine')]
        else:
            # 已经是目标值则跳过
            current = pf.first_line_indent
            if current is not None and current == first_indent:
                return
            pf.first_line_indent = first_indent
    else:
        pf.first_line_indent = Cm(0)


def format_paragraph(paragraph, config):
    """格式化整个段落（段落格式 + 所有 run 的字体）

    Args:
        paragraph: docx Paragraph 对象
        config: 段落配置字典
    """
    set_paragraph_format(paragraph, config)

    # 标题段落：数字也使用中文字体，避免编号突兀
    ascii_cn = config.get("ascii_use_cn", False)

    for run in paragraph.runs:
        set_run_font(
            run,
            config["font_cn"],
            config["font_en"],
            config["size"],
            config.get("bold", False),
            ascii_use_cn=ascii_cn,
        )


def set_text_color_black(paragraph):
    """将段落中所有文字颜色设为黑色"""
    for run in paragraph.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)


def set_all_text_black(doc):
    """将文档中所有文字颜色统一设置为黑色

    Args:
        doc: Document 对象
    """
    # 段落内文字
    for para in doc.paragraphs:
        set_text_color_black(para)

    # 表格内文字
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    set_text_color_black(para)


def has_image(paragraph):
    """检查段落是否包含图片"""
    return bool(
        paragraph._element.findall('.//' + qn('wp:inline'))
        + paragraph._element.findall('.//' + qn('wp:anchor'))
    )


def fix_image_paragraphs(doc):
    """将所有包含图片的段落设为单倍行距（防止固定行距导致图片显示不全）"""
    for para in doc.paragraphs:
        if has_image(para):
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE


def is_heading(text, level):
    """判断文本是否为指定级别的标题

    Args:
        text: 段落文本
        level: 标题级别（1-4）

    Returns:
        bool
    """
    text = text.strip()
    if not text:
        return False

    if level == 1:
        # 匹配：1 / 一、 / 第一章 / 1. / 第1章
        patterns = [
            r'^第[一二三四五六七八九十百]+[章部篇]',
            r'^[一二三四五六七八九十]+[、．.]',
            r'^\d+[\.、\s]',
            r'^第\d+[章部篇]',
        ]
    elif level == 2:
        # 匹配：3.1 / （一） / (一) / 1.1
        patterns = [
            r'^\d+\.\d+[\s\.、]',
            r'^（[一二三四五六七八九十]+）',
            r'^\([一二三四五六七八九十]+\)',
            r'^\d+\.\d+$',
        ]
    elif level == 3:
        # 匹配：3.1.1 / 1. / （1）
        patterns = [
            r'^\d+\.\d+\.\d+[\s\.、]',
            r'^\d+\.\d+\.\d+$',
            r'^\d+\.[\s]',
            r'^（\d+）',
        ]
    elif level == 4:
        # 匹配：3.1.1.1 / (1) / ①
        patterns = [
            r'^\d+\.\d+\.\d+\.\d+',
            r'^\(\d+\)',
            r'^①|②|③|④|⑤|⑥|⑦|⑧|⑨|⑩',
        ]
    else:
        return False

    for pattern in patterns:
        if re.match(pattern, text):
            return True
    return False


def is_figure_caption(text):
    """判断是否是图题

    图题特征：以"图X"开头 + 内容简短（通常不超过50字）。
    区分于"图X传递了三个关键信息..."这类分析段落。
    """
    text = text.strip()
    patterns = [
        r'^图\s*\d+',
        r'^Figure\s*\d+',
        r'^Fig\.\s*\d+',
    ]
    for pattern in patterns:
        if re.match(pattern, text):
            # 图题都很短，长段落是分析文字
            if len(text) > 60:
                return False
            # 排除以句号、逗号结尾的分析句（如图X传递了...）
            if any(kw in text for kw in ['传递', '展示', '给出', '显示', '汇总', '比较', '说明', '表明', '反映']):
                return False
            return True
    return False


def is_table_caption(text):
    """判断是否是表题

    表题特征：以"表X"开头 + 内容简短（通常不超过40字）。
    区分于"表X给出了..."这类分析段落。
    """
    text = text.strip()
    patterns = [
        r'^表\s*\d+',
        r'^Table\s*\d+',
    ]
    for pattern in patterns:
        if re.match(pattern, text):
            if len(text) > 50:
                return False
            if any(kw in text for kw in ['给出', '展示', '显示', '汇总', '说明', '可见', '表明', '仅提供']):
                return False
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


def detect_heading_by_style(paragraph):
    """通过 Word 样式名检测标题级别

    Returns:
        标题级别（1-4），0 表示不是标题
    """
    style_name = paragraph.style.name if paragraph.style else ""
    if style_name.startswith("Heading 1"):
        return 1
    elif style_name.startswith("Heading 2"):
        return 2
    elif style_name.startswith("Heading 3"):
        return 3
    elif style_name.startswith("Heading 4"):
        return 4
    return 0


def detect_section_type(paragraph, text, in_references):
    """检测段落类型

    Args:
        paragraph: Paragraph 对象
        text: 段落文本（已 strip）
        in_references: 是否在参考文献区域

    Returns:
        section_type 字符串
    """
    if not text:
        return "empty"

    # 获取段落样式名
    style_name = paragraph.style.name if paragraph.style else ""

    # List Bullet / List Number 样式直接判为正文（不走标题检测）
    if "Bullet" in style_name or "List" in style_name:
        return "body"

    # 通过 Word 样式检测标题（最可靠）
    style_level = detect_heading_by_style(paragraph)
    if style_level > 0:
        return f"heading{style_level}"

    # 回退：通过文本模式检测标题（适用于直接格式化的文档）
    for level in range(4, 0, -1):  # 从高级到低级，避免误判
        if is_heading(text, level):
            return f"heading{level}"

    # 参考文献区域
    if in_references and is_reference(text):
        return "reference"

    # 图题/表题
    if is_figure_caption(text):
        return "figure_caption"
    if is_table_caption(text):
        return "table_caption"

    # 摘要标题
    if "摘要" in text and len(text) < 10:
        return "abstract_title"

    return "body"


def set_page_margins(doc, config):
    """设置页边距

    Args:
        doc: Document 对象
        config: page 配置字典，包含 margin_top/bottom/left/right（单位 cm）
    """
    page_config = config.get("page", {})
    if not page_config:
        return

    for section in doc.sections:
        if "margin_top" in page_config:
            section.top_margin = Cm(page_config["margin_top"])
        if "margin_bottom" in page_config:
            section.bottom_margin = Cm(page_config["margin_bottom"])
        if "margin_left" in page_config:
            section.left_margin = Cm(page_config["margin_left"])
        if "margin_right" in page_config:
            section.right_margin = Cm(page_config["margin_right"])


def format_document(doc, config):
    """主排版函数

    Args:
        doc: Document 对象
        config: 已解析的配置字典

    Returns:
        (doc, summary) — 文档对象和修改摘要
    """
    summary = {
        "headings": 0,
        "body": 0,
        "abstract": 0,
        "reference": 0,
        "figure_caption": 0,
        "table_caption": 0,
        "tables": 0,
        "images_fixed": 0,
        "text_blackened": False,
    }

    in_references = False
    in_abstract = False
    first_heading_seen = False  # 标记是否已遇到第一个标题

    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()

        if not text:
            continue

        # 检测是否进入参考文献部分
        style_level = detect_heading_by_style(paragraph)
        # 回退：文本模式检测（适用于直接格式化的文档）
        if style_level == 0:
            for lv in range(4, 0, -1):
                if is_heading(text, lv):
                    style_level = lv
                    break
        if style_level == 1 and "参考文献" in text:
            in_references = True
            in_abstract = False
        elif style_level == 1:
            # 遇到其他一级标题时退出摘要区域
            in_abstract = False

        # 检测是否进入摘要
        if "摘要" in text and len(text) < 10:
            in_abstract = True

        # 标记是否已遇到第一个标题
        if style_level > 0:
            first_heading_seen = True

        # 确定段落类型
        section_type = detect_section_type(paragraph, text, in_references)

        # 标题页处理：第一个标题之前的 Normal 段落居中无缩进
        if not first_heading_seen and section_type == "body":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Cm(0)
            # 清除可能残留的 firstLineChars
            pPr = paragraph._element.find(qn('w:pPr'))
            if pPr is not None:
                ind = pPr.find(qn('w:ind'))
                if ind is not None:
                    if ind.get(qn('w:firstLineChars')) is not None:
                        del ind.attrib[qn('w:firstLineChars')]
            # 保持原字体大小，只统一字体
            for run in paragraph.runs:
                cn = "宋体"
                sz = run.font.size if run.font.size else Pt(14)
                set_run_font(run, cn, "Times New Roman", sz, run.bold or False)
            summary["body"] += 1
            continue

        if section_type.startswith("heading"):
            level = int(section_type[-1])
            cfg_key = f"heading{level}"
            if cfg_key in config:
                format_paragraph(paragraph, config[cfg_key])
            summary["headings"] += 1

        elif section_type == "reference":
            format_paragraph(paragraph, config["reference"])
            summary["reference"] += 1

        elif section_type == "figure_caption":
            format_paragraph(paragraph, config["figure_caption"])
            summary["figure_caption"] += 1

        elif section_type == "table_caption":
            format_paragraph(paragraph, config["table_caption"])
            summary["table_caption"] += 1

        elif section_type == "abstract_title":
            format_paragraph(paragraph, config["abstract_title"])
            in_abstract = True
            summary["abstract"] += 1

        elif in_abstract:
            format_paragraph(paragraph, config.get("abstract_body", config["body"]))
            summary["abstract"] += 1

        else:
            style_name = paragraph.style.name if paragraph.style else ""

            # 检查是否是 List 段落（样式名或 XML 中有 ilvl/numPr 残留）
            pPr = paragraph._element.find(qn('w:pPr'))
            is_list = ("Bullet" in style_name or "List" in style_name
                       or (pPr is not None and (
                           pPr.find(qn('w:numPr')) is not None
                           or (pPr.find(qn('w:ilvl')) is not None))))

            if is_list:
                # List 段落：转为 Normal 样式，消除项目符号
                paragraph.style = doc.styles['Normal']

                # 用正文格式但不缩进
                body_cfg = dict(config["body"])
                body_cfg["first_line_indent"] = Pt(0)
                format_paragraph(paragraph, body_cfg)

                # 彻底清除缩进（Pt(0) 不会清除 firstLineChars）
                pPr = paragraph._element.find(qn('w:pPr'))
                if pPr is not None:
                    ind = pPr.find(qn('w:ind'))
                    if ind is not None:
                        if ind.get(qn('w:firstLineChars')) is not None:
                            del ind.attrib[qn('w:firstLineChars')]
                        if ind.get(qn('w:firstLine')) is not None:
                            del ind.attrib[qn('w:firstLine')]
                    numPr = pPr.find(qn('w:numPr'))
                    if numPr is not None:
                        pPr.remove(numPr)
                summary["body"] += 1
                continue  # 跳过后续处理，避免正文分支覆盖缩进
            else:
                # 检查是否曾被清除过缩进（无 firstLineChars 且无 firstLine）
                # 如果是，说明是已处理过的 List Bullet，不加缩进
                pPr_check = paragraph._element.find(qn('w:pPr'))
                ind_check = pPr_check.find(qn('w:ind')) if pPr_check is not None else None
                already_cleared = (ind_check is not None
                                   and ind_check.get(qn('w:firstLineChars')) is None
                                   and (ind_check.get(qn('w:firstLine')) is None
                                        or ind_check.get(qn('w:firstLine')) == '0'))

                if not already_cleared:
                    format_paragraph(paragraph, config["body"])
            summary["body"] += 1

    # 格式化表格
    for table in doc.tables:
        table_cfg = config.get("table", config["body"])
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.paragraph_format.line_spacing = table_cfg.get("line_spacing", Pt(24))
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                    para.alignment = table_cfg.get("alignment", WD_ALIGN_PARAGRAPH.CENTER)
                    for run in para.runs:
                        set_run_font(
                            run,
                            table_cfg["font_cn"],
                            table_cfg["font_en"],
                            table_cfg["size"],
                            table_cfg.get("bold", False),
                        )
        summary["tables"] += 1

    # 图片段落行距修复
    img_count = 0
    for para in doc.paragraphs:
        if has_image(para):
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            img_count += 1
    summary["images_fixed"] = img_count

    # 统一文字颜色为黑色
    set_all_text_black(doc)
    summary["text_blackened"] = True

    # 设置页边距
    set_page_margins(doc, config)

    return doc, summary


def format_summary_text(summary):
    """生成排版摘要文本

    Args:
        summary: format_document 返回的摘要字典

    Returns:
        格式化的摘要文本
    """
    lines = [
        "排版完成！",
        f"  标题: {summary['headings']} 个",
        f"  正文: {summary['body']} 个段落",
        f"  摘要: {summary['abstract']} 个段落",
        f"  参考文献: {summary['reference']} 条",
        f"  图题: {summary['figure_caption']} 个",
        f"  表题: {summary['table_caption']} 个",
        f"  表格: {summary['tables']} 个",
        f"  图片行距修复: {summary['images_fixed']} 个",
        f"  文字颜色统一: {'是' if summary['text_blackened'] else '否'}",
    ]
    return "\n".join(lines)
