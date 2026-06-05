#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""AI痕迹清除模块"""

import re


# 项目符号（全角和半角）
BULLET_CHARS = [
    '•', '·', '■', '▪', '◆', '◇', '○', '●', '★', '☆',
    '▶', '▷', '►', '‣', '⁃', '⁌', '⁍',
    '◎', '◉', '◐', '◑', '◒', '◓',
    '▫', '▬', '▭', '▮', '▯', '▰', '▱',
    '▸', '▹', '▻', '▼', '▽', '▾', '▿',
    '◈', '◊', '◌', '◍',
    '↖', '↗', '↘', '↙', '↺', '↻',
    '✦', '✧', '✩', '✪', '✫', '✬',
    '✠', '✡', '✢', '✣', '✤', '✥',
]

# 列表符号
LIST_CHARS = [
    '①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩',
    '❶', '❷', '❸', '❹', '❺', '❻', '❼', '❽', '❾', '❿',
    '⑴', '⑵', '⑶', '⑷', '⑸', '⑹', '⑺', '⑻', '⑼', '⑽',
    'Ⅰ', 'Ⅱ', 'Ⅲ', 'Ⅳ', 'Ⅴ', 'Ⅵ', 'Ⅶ', 'Ⅷ', 'Ⅸ', 'Ⅹ',
    'ⅰ', 'ⅱ', 'ⅲ', 'ⅳ', 'ⅴ', 'ⅵ', 'ⅶ', 'ⅷ', 'ⅸ', 'ⅹ',
]

# AI 常用表达替换表（保守替换，避免误伤）
AI_PHRASES = {
    '本文旨在': '本文主要',
    '本研究旨在': '本研究主要',
    '研究表明': '结果显示',
    '研究发现': '结果发现',
    '具有重要意义': '很有意义',
    '具有重要价值': '很有价值',
    '为提供参考': '可作参考',
    '为提供依据': '可作依据',
}


def _update_runs_text(paragraph, new_text):
    """更新段落文本，保留原有格式

    Args:
        paragraph: Paragraph 对象
        new_text: 新文本
    """
    if not paragraph.runs:
        return

    total_len = sum(len(run.text) for run in paragraph.runs)
    if total_len == 0:
        return

    remaining = new_text
    for i, run in enumerate(paragraph.runs):
        if i == len(paragraph.runs) - 1:
            run.text = remaining
        else:
            ratio = len(run.text) / total_len
            chars = int(len(new_text) * ratio)
            run.text = remaining[:chars]
            remaining = remaining[chars:]


def remove_bullet_symbols(text):
    """删除文本中的项目符号

    Args:
        text: 原始文本

    Returns:
        清理后的文本
    """
    for char in BULLET_CHARS:
        # 先处理行首：删除符号及紧跟的空格
        while text.startswith(char):
            text = text[len(char):].lstrip()

        # 再处理行内：符号替换为逗号（避免删除后两个词粘连）
        # 情况1：符号前后都有内容，如"方法•对比"
        # 情况2：符号前面有空格，如" 方法•对比"
        # 情况3：符号后面有空格，如"方法• 对比"
        text = text.replace(' ' + char + ' ', '，')
        text = text.replace(char + ' ', '，')
        text = text.replace(' ' + char, '，')
        text = text.replace(char, '，')

    # 清理多余逗号和首尾逗号
    text = re.sub(r'，{2,}', '，', text)
    text = text.strip('，').strip()
    return text


def remove_list_markers(text):
    """删除文本中的列表符号

    Args:
        text: 原始文本

    Returns:
        清理后的文本
    """
    for char in LIST_CHARS:
        # 删除行首列表符号
        while text.startswith(char):
            text = text[len(char):].lstrip()
        # 行内列表符号替换为逗号
        text = text.replace(char + ' ', '，')
        text = text.replace(' ' + char, '，')
        text = text.replace(char, '，')

    # 删除"一是...二是..."格式
    text = re.sub(r'^[一二三四五六七八九十]+是[：:，,]?\s*', '', text)

    # 清理多余逗号
    text = re.sub(r'，{2,}', '，', text)
    text = text.strip('，').strip()
    return text


def replace_ai_phrases(text):
    """替换AI常用表达

    Args:
        text: 原始文本

    Returns:
        替换后的文本
    """
    for old, new in AI_PHRASES.items():
        text = text.replace(old, new)
    return text


def remove_ai_traces(doc):
    """消除文档中的AI痕迹（保留格式）

    Args:
        doc: Document 对象

    Returns:
        (doc, count) — 文档对象和修改的段落数
    """
    count = 0

    for para in doc.paragraphs:
        if not para.runs:
            continue

        full_text = para.text
        new_text = full_text

        # 删除项目符号
        new_text = remove_bullet_symbols(new_text)

        # 删除列表符号
        new_text = remove_list_markers(new_text)

        # 替换AI常用表达
        new_text = replace_ai_phrases(new_text)

        # 如果文本有变化，更新 runs（保留格式）
        if new_text != full_text:
            _update_runs_text(para, new_text)
            count += 1

    return doc, count
