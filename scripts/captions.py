#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""图表序号检测、补全与重编号模块"""

import re
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def has_image(paragraph):
    """检查段落是否包含图片"""
    return bool(
        paragraph._element.findall('.//' + qn('wp:inline'))
        + paragraph._element.findall('.//' + qn('wp:anchor'))
    )


def get_paragraph_text(para_elem):
    """从 w:p 元素中正确提取文本（文本在 w:r/w:t 中）"""
    texts = []
    for r_elem in para_elem.findall('.//' + qn('w:r')):
        for t_elem in r_elem.findall(qn('w:t')):
            if t_elem.text:
                texts.append(t_elem.text)
    return ''.join(texts)


def find_image_paragraphs(doc):
    """找到所有包含图片的段落

    Returns:
        list of (para_index, paragraph) 元组
    """
    result = []
    for i, para in enumerate(doc.paragraphs):
        if has_image(para):
            result.append((i, para))
    return result


def find_table_positions(doc):
    """找到所有表格在 body 元素中的位置

    Returns:
        list of body_element_index
    """
    positions = []
    body = doc.element.body
    for i, child in enumerate(body):
        if child.tag == qn('w:tbl'):
            positions.append(i)
    return positions


def check_figure_caption(doc):
    """检查图片是否有图序图题

    Returns:
        (missing_captions, figure_nums)
        - missing_captions: 缺失图题的图片信息列表
        - figure_nums: 已有的图序号列表
    """
    image_paras = find_image_paragraphs(doc)
    missing_captions = []
    figure_nums = []

    for para_idx, para in image_paras:
        has_caption = False
        # 检查图片下方是否有图题
        if para_idx + 1 < len(doc.paragraphs):
            next_text = doc.paragraphs[para_idx + 1].text.strip()
            match = re.match(r'^图\s*(\d+)', next_text)
            if match:
                has_caption = True
                figure_nums.append(int(match.group(1)))

        if not has_caption:
            missing_captions.append(para_idx)

    return missing_captions, figure_nums


def check_table_caption(doc):
    """检查表格是否有表序表题

    Returns:
        (missing_captions, table_nums)
    """
    table_positions = find_table_positions(doc)
    missing_captions = []
    table_nums = []
    body = doc.element.body

    for i, tbl_pos in enumerate(table_positions):
        has_caption = False
        # 检查表格上方是否有表题
        if tbl_pos > 0:
            prev_elem = body[tbl_pos - 1]
            if prev_elem.tag == qn('w:p'):
                prev_text = get_paragraph_text(prev_elem)
                match = re.match(r'^表\s*(\d+)', prev_text.strip())
                if match:
                    has_caption = True
                    table_nums.append(int(match.group(1)))

        if not has_caption:
            missing_captions.append(i)

    return missing_captions, table_nums


def check_sequence(nums):
    """检查序号是否连续

    Args:
        nums: 序号列表

    Returns:
        错误信息列表（空表示无错误）
    """
    if not nums:
        return []
    errors = []
    expected = list(range(1, len(nums) + 1))
    if nums != expected:
        for i, (actual, exp) in enumerate(zip(nums, expected)):
            if actual != exp:
                errors.append(f"第{i+1}个应为{exp}，实际为{actual}")
    return errors


def get_context_text(doc, para_index, direction='before', count=3):
    """获取指定段落的上下文文本（跳过图题/表题引用）

    Args:
        doc: Document 对象
        para_index: 段落索引
        direction: 'before' 或 'after'
        count: 获取几个相邻段落

    Returns:
        文本列表
    """
    texts = []
    if direction == 'before':
        start = max(0, para_index - count)
        for i in range(start, para_index):
            t = doc.paragraphs[i].text.strip()
            if t and not re.match(r'^图\s*\d+', t) and not re.match(r'^表\s*\d+', t):
                texts.append(t)
    elif direction == 'after':
        end = min(len(doc.paragraphs), para_index + count + 1)
        for i in range(para_index + 1, end):
            t = doc.paragraphs[i].text.strip()
            if t and not re.match(r'^图\s*\d+', t) and not re.match(r'^表\s*\d+', t):
                texts.append(t)
    return texts


def generate_caption_text(context_texts, is_figure=True):
    """根据上下文生成图题/表题内容

    Args:
        context_texts: 上下文文本列表
        is_figure: True 生成图题，False 生成表题

    Returns:
        生成的题注文本
    """
    prefix = "图" if is_figure else "表"
    if not context_texts:
        return f"{prefix}题"

    text = context_texts[-1].strip()

    # 去掉无意义的词
    stop_words = ['的', '了', '在', '是', '有', '和', '与', '及', '等', '先', '然后', '接着']
    words = [w for w in text if w not in stop_words]

    # 提取关键信息
    if len(words) > 25:
        caption = ''.join(words[:25])
    elif len(words) > 10:
        caption = ''.join(words)
    else:
        caption = text[:20]

    return caption


def _create_caption_paragraph(text, font_name='宋体', font_size=Pt(10.5)):
    """创建图题/表题段落的 XML 元素

    Args:
        text: 段落文本
        font_name: 中文字体
        font_size: 字号

    Returns:
        OxmlElement (w:p)
    """
    new_para = OxmlElement('w:p')
    new_run = OxmlElement('w:r')
    new_rpr = OxmlElement('w:rPr')

    # 字体
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), font_name)
    new_rpr.append(rFonts)

    # 字号
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(font_size.pt * 2)))
    new_rpr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(int(font_size.pt * 2)))
    new_rpr.append(szCs)

    # 颜色黑色
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '000000')
    new_rpr.append(color)

    new_run.append(new_rpr)
    new_t = OxmlElement('w:t')
    new_t.text = text
    new_run.append(new_t)
    new_para.append(new_run)

    # 居中
    new_ppr = OxmlElement('w:pPr')
    new_jc = OxmlElement('w:jc')
    new_jc.set(qn('w:val'), 'center')
    new_ppr.append(new_jc)
    new_para.insert(0, new_ppr)

    return new_para


def insert_paragraph_before(doc, element_index, text, font_name='宋体', font_size=Pt(10.5)):
    """在指定 body 元素之前插入新段落（用于表题）"""
    body = doc.element.body
    new_para = _create_caption_paragraph(text, font_name, font_size)
    target_elem = body[element_index]
    target_elem.addprevious(new_para)
    return new_para


def insert_paragraph_after(doc, body_element_index, text, font_name='宋体', font_size=Pt(10.5)):
    """在指定 body 元素之后插入新段落（用于图题）"""
    body = doc.element.body
    new_para = _create_caption_paragraph(text, font_name, font_size)
    body.insert(body_element_index + 1, new_para)
    return new_para


def fix_figure_captions(doc, font_name='宋体', font_size=Pt(10.5)):
    """自动补全图序图题（幂等处理，不会重复添加）

    处理逻辑：
    1. 找到所有图片的 body 元素索引
    2. 检查图片紧后面（body 索引+1）是否有"图N"段落
    3. 只为没有紧邻图题的图片插入新图题
    4. 从后往前插入避免索引偏移

    Returns:
        新插入的图题数量
    """
    body = doc.element.body

    # 找到所有图片的 body 索引
    image_body_indices = []
    for i, child in enumerate(body):
        if child.tag == qn('w:p'):
            if child.findall('.//' + qn('wp:inline')) or child.findall('.//' + qn('wp:anchor')):
                image_body_indices.append(i)

    if not image_body_indices:
        return 0

    # 检查每个图片紧后面是否有图题
    missing = []
    for img_idx in image_body_indices:
        has_caption = False
        if img_idx + 1 < len(body):
            next_elem = body[img_idx + 1]
            if next_elem.tag == qn('w:p'):
                next_text = get_paragraph_text(next_elem)
                if re.match(r'^图\s*\d+', next_text.strip()):
                    has_caption = True
        if not has_caption:
            missing.append(img_idx)

    # 从后往前插入（避免索引偏移）
    count = 0
    for img_idx in reversed(missing):
        fig_num = image_body_indices.index(img_idx) + 1
        insert_paragraph_after(doc, img_idx, f'图{fig_num}', font_name, font_size)
        count += 1

    return count


def fix_table_captions(doc, font_name='宋体', font_size=Pt(10.5)):
    """自动补全表序表题（幂等处理，不会重复添加）

    处理逻辑：
    1. 找到所有表格
    2. 通过 body 元素直接检查表格前面是否有"表N"段落
    3. 从后往前插入，避免位置偏移问题

    Returns:
        新插入的表题数量
    """
    body = doc.element.body

    # 找到所有表格及其 body 索引
    table_indices = []
    for i, child in enumerate(body):
        if child.tag == qn('w:tbl'):
            table_indices.append(i)

    if not table_indices:
        return 0

    # 检查每个表格前面是否有表题，记录需要补全的
    missing = []
    for tbl_idx in table_indices:
        has_caption = False
        for check in [1, 2]:
            if tbl_idx - check >= 0:
                prev = body[tbl_idx - check]
                if prev.tag == qn('w:p'):
                    prev_text = get_paragraph_text(prev)
                    if re.match(r'^表\s*\d+', prev_text.strip()):
                        has_caption = True
                        break
        if not has_caption:
            missing.append(tbl_idx)

    # 从后往前插入（避免索引偏移）
    count = 0
    for tbl_idx in reversed(missing):
        table_num = table_indices.index(tbl_idx) + 1
        insert_paragraph_before(doc, tbl_idx, f'表{table_num}', font_name, font_size)
        count += 1

    return count


def fix_incomplete_captions(doc):
    """修复格式不规范的图表题（如"图1：xxx"改为"图1 xxx"）

    处理逻辑：
    1. 找到所有图题/表题
    2. 统一格式为"图N 内容"（去掉冒号，加空格）
    3. 不会生成新内容，只修正格式

    Returns:
        修正的图表题数量
    """
    count = 0

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()

        # 匹配"图N：内容"或"图N:内容"（需要去掉冒号）
        match = re.match(r'^(图|表)\s*(\d+)\s*[：:]\s*(.+)$', text)
        if match:
            prefix = match.group(1)
            num = match.group(2)
            content = match.group(3).strip()
            new_text = f'{prefix}{num} {content}'

            # 更新文本
            for run in para.runs:
                run.text = ''
            if para.runs:
                para.runs[0].text = new_text

            count += 1

    return count


def check_all_captions(doc):
    """检查所有图表序号的状态（只报告，不修改）

    Returns:
        dict: 检查结果
    """
    missing_figures, figure_nums = check_figure_caption(doc)
    missing_tables, table_nums = check_table_caption(doc)
    fig_seq_errors = check_sequence(figure_nums)
    tbl_seq_errors = check_sequence(table_nums)

    return {
        'missing_figures': missing_figures,
        'missing_tables': missing_tables,
        'figure_nums': figure_nums,
        'table_nums': table_nums,
        'figure_seq_errors': fig_seq_errors,
        'table_seq_errors': tbl_seq_errors,
    }


def format_caption_report(result):
    """将图表序号检查结果格式化为可读文本

    Args:
        result: check_all_captions 返回的字典

    Returns:
        格式化的报告文本
    """
    lines = []
    lines.append("【图表序号检查】")

    # 图序号
    fig_nums = result['figure_nums']
    missing_fig = result['missing_figures']
    fig_errors = result['figure_seq_errors']

    if fig_nums:
        lines.append(f"  图序号: {fig_nums}")
    else:
        lines.append("  图序号: 无")

    if missing_fig:
        lines.append(f"  ⚠ {len(missing_fig)} 个图片缺少图题")
    if fig_errors:
        lines.append(f"  ⚠ 图序号不连续: {', '.join(fig_errors)}")

    # 表序号
    tbl_nums = result['table_nums']
    missing_tbl = result['missing_tables']
    tbl_errors = result['table_seq_errors']

    if tbl_nums:
        lines.append(f"  表序号: {tbl_nums}")
    else:
        lines.append("  表序号: 无")

    if missing_tbl:
        lines.append(f"  ⚠ {len(missing_tbl)} 个表格缺少表题")
    if tbl_errors:
        lines.append(f"  ⚠ 表序号不连续: {', '.join(tbl_errors)}")

    if not missing_fig and not missing_tbl and not fig_errors and not tbl_errors:
        lines.append("  ✅ 图表序号正常")

    return "\n".join(lines)
