#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""论文排版配置模块"""

from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import json
import os

# ========== 字号对照表 ==========
FONT_SIZE_MAP = {
    "初号": Pt(42),
    "小初": Pt(36),
    "一号": Pt(26),
    "小一": Pt(24),
    "二号": Pt(22),
    "小二": Pt(18),
    "三号": Pt(16),
    "小三": Pt(15),
    "四号": Pt(14),
    "小四": Pt(12),
    "五号": Pt(10.5),
    "小五": Pt(9),
}


def get_size(size_spec):
    """将字号名称或磅值转为 Pt 对象

    Args:
        size_spec: 字号名称（如"小四"）或磅值数字（如12）
    """
    if isinstance(size_spec, str):
        if size_spec in FONT_SIZE_MAP:
            return FONT_SIZE_MAP[size_spec]
        # 尝试解析为数字
        try:
            return Pt(float(size_spec))
        except ValueError:
            raise ValueError(f"无法识别的字号: {size_spec}")
    elif isinstance(size_spec, (int, float)):
        return Pt(size_spec)
    else:
        raise ValueError(f"无效的字号规格: {size_spec}")


# ========== 默认排版规范 ==========
DEFAULT_CONFIG = {
    "heading1": {
        "font_cn": "黑体",
        "font_en": "Times New Roman",
        "size": "小三",
        "bold": False,
        "alignment": "left",
        "space_before": "0.5line",
        "space_after": "0.5line",
        "line_spacing": 24,
        "first_line_indent": 0,
    },
    "heading2": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": "四号",
        "bold": False,
        "alignment": "left",
        "space_before": "0.5line",
        "space_after": "0.5line",
        "line_spacing": 24,
        "first_line_indent": 0,
    },
    "heading3": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": "小四",
        "bold": False,
        "alignment": "left",
        "space_before": 0,
        "space_after": 0,
        "line_spacing": 24,
        "first_line_indent": 0,
    },
    "heading4": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": "小四",
        "bold": False,
        "alignment": "left",
        "space_before": 0,
        "space_after": 0,
        "line_spacing": 24,
        "first_line_indent": 0,
    },
    "body": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": "小四",
        "bold": False,
        "alignment": "justify",
        "space_before": 0,
        "space_after": 0,
        "line_spacing": 24,
        "first_line_indent": "2char",
    },
    "abstract_title": {
        "font_cn": "黑体",
        "font_en": "Times New Roman",
        "size": "四号",
        "bold": False,
        "alignment": "center",
        "space_before": 0,
        "space_after": 0,
        "line_spacing": 24,
    },
    "abstract_body": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": "小四",
        "bold": False,
        "alignment": "justify",
        "space_before": 0,
        "space_after": 0,
        "line_spacing": 24,
        "first_line_indent": "2char",
    },
    "reference": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": "五号",
        "bold": False,
        "alignment": "justify",
        "space_before": 0,
        "space_after": 0,
        "line_spacing": 24,
        "first_line_indent": 0,
    },
    "figure_caption": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": "五号",
        "bold": False,
        "alignment": "center",
        "space_before": 0,
        "space_after": 0,
        "line_spacing": 24,
    },
    "table_caption": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": "五号",
        "bold": False,
        "alignment": "center",
        "space_before": 0,
        "space_after": 0,
        "line_spacing": 24,
    },
    "table": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": "五号",
        "bold": False,
        "alignment": "center",
        "line_spacing": 24,
    },
    "page": {
        "margin_top": 2.54,     # cm
        "margin_bottom": 2.54,
        "margin_left": 3.17,
        "margin_right": 3.17,
    },
}

# ========== 理工类默认规范 ==========
SCIENCE_CONFIG = DEFAULT_CONFIG.copy()

# ========== 文史类默认规范 ==========
ARTS_CONFIG = {
    **DEFAULT_CONFIG,
    "heading1": {
        **DEFAULT_CONFIG["heading1"],
        "font_cn": "黑体",
        "size": "三号",
        "bold": True,
    },
    "heading2": {
        **DEFAULT_CONFIG["heading2"],
        "font_cn": "黑体",
        "size": "小三",
        "bold": True,
    },
    "heading3": {
        **DEFAULT_CONFIG["heading3"],
        "font_cn": "仿宋",
        "size": "四号",
        "bold": True,
    },
    "body": {
        **DEFAULT_CONFIG["body"],
        "font_cn": "仿宋",
        "size": "四号",
        "first_line_indent": "2char",
    },
}

# 预设规范
PRESET_CONFIGS = {
    "science": SCIENCE_CONFIG,
    "arts": ARTS_CONFIG,
}


# ========== 对齐方式映射 ==========
ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def resolve_config(raw_config):
    """将原始配置（字符串字号、字符串对齐等）转为 python-docx 可用的格式

    Args:
        raw_config: 原始配置字典

    Returns:
        解析后的配置字典，所有值都已转为 python-docx 对象
    """
    resolved = {}
    for section, values in raw_config.items():
        if section == "page":
            resolved[section] = values
            continue
        if not isinstance(values, dict):
            resolved[section] = values
            continue

        section_resolved = {}
        for key, val in values.items():
            if key == "size":
                section_resolved[key] = get_size(val)
            elif key == "alignment":
                section_resolved[key] = ALIGNMENT_MAP.get(val, WD_ALIGN_PARAGRAPH.JUSTIFY)
            elif key in ("line_spacing", "space_before", "space_after", "first_line_indent"):
                # 支持 "2char" / "1line" 格式（保持字符串，由 formatter 处理）
                if isinstance(val, str) and (val.endswith("char") or val.endswith("line")):
                    section_resolved[key] = val
                else:
                    section_resolved[key] = Pt(val) if val else Pt(0)
            else:
                section_resolved[key] = val
        resolved[section] = section_resolved
    return resolved


def load_config(config_path=None, preset=None):
    """加载排版配置

    Args:
        config_path: JSON 配置文件路径（可选）
        preset: 预设名称，"science" 或 "arts"（可选）

    Returns:
        解析后的配置字典
    """
    if config_path and os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        # 如果文件只包含部分配置，用默认值补全
        config = {**DEFAULT_CONFIG, **raw}
    elif preset and preset in PRESET_CONFIGS:
        config = PRESET_CONFIGS[preset]
    else:
        config = DEFAULT_CONFIG

    return resolve_config(config)


def save_config(config, output_path):
    """将配置保存为 JSON 文件"""
    # 反向解析：将 Pt 对象转回数字
    serializable = {}
    for section, values in config.items():
        if section == "page":
            serializable[section] = values
            continue
        if not isinstance(values, dict):
            serializable[section] = values
            continue

        section_out = {}
        for key, val in values.items():
            if hasattr(val, 'pt'):
                section_out[key] = val.pt
            elif hasattr(val, 'value'):
                # WD_ALIGN_PARAGRAPH
                reverse_map = {v: k for k, v in ALIGNMENT_MAP.items()}
                section_out[key] = reverse_map.get(val, str(val))
            else:
                section_out[key] = val
        serializable[section] = section_out

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(serializable, f, ensure_ascii=False, indent=2)
