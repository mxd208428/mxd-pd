#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""论文排版主入口脚本

用法:
    python format_paper.py 论文.docx                    # 排版模式
    python format_paper.py 论文.docx -o 输出.docx       # 指定输出路径
    python format_paper.py 论文.docx --check             # 检查模式
    python format_paper.py 论文.docx --preset arts       # 使用文史类规范
    python format_paper.py 论文.docx --config my.json    # 使用自定义配置
    python format_paper.py 论文.docx --three-line        # 转换三线表
    python format_paper.py 论文.docx --remove-ai         # 清除AI痕迹
    python format_paper.py 论文.docx --fix-captions      # 补全图表序号
"""

import argparse
import sys
import os

# 确保 scripts 目录在 path 中（支持直接运行）
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from scripts.config import load_config
from scripts.validator import validate_docx_file, safe_save_document, generate_output_path
from scripts.formatter import format_document, format_summary_text
from scripts.captions import check_all_captions, format_caption_report
from scripts.ai_traces import remove_ai_traces
from scripts.three_line_table import convert_all_tables_to_three_line
from scripts.checker import check_format, format_check_report


def main():
    parser = argparse.ArgumentParser(
        description='论文排版工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument('input', help='输入 .docx 文件路径')
    parser.add_argument('-o', '--output', help='输出文件路径（默认: 原文件名_已排版.docx）')
    parser.add_argument('--check', action='store_true', help='检查模式（只报告问题，不修改）')
    parser.add_argument('--preset', choices=['science', 'arts'], help='使用预设规范（science=理工类, arts=文史类）')
    parser.add_argument('--config', help='自定义配置文件路径（JSON 格式）')
    parser.add_argument('--three-line', action='store_true', help='将所有表格转换为三线表')
    parser.add_argument('--remove-ai', action='store_true', help='清除AI痕迹')
    parser.add_argument('--fix-captions', action='store_true', help='检查图表序号（只报告缺失，不自动插入）')
    parser.add_argument('--overwrite', action='store_true', help='覆盖已有文件')

    args = parser.parse_args()

    # 1. 验证输入文件
    is_valid, errors = validate_docx_file(args.input)
    if not is_valid:
        print("❌ 文件验证失败:")
        for error in errors:
            print(f"   - {error}")
        sys.exit(1)

    # 2. 加载配置
    config = load_config(config_path=args.config, preset=args.preset)

    # 3. 打开文档
    try:
        doc = Document(args.input)
    except Exception as e:
        print(f"❌ 无法打开文档: {e}")
        sys.exit(1)

    # 4. 检查模式
    if args.check:
        result = check_format(doc, config)
        report = format_check_report(result)
        print(report)
        sys.exit(0)

    # 5. 排版模式 — 按需执行各功能
    actions_performed = []

    # 图表序号检查
    if args.fix_captions:
        caption_result = check_all_captions(doc)
        caption_report = format_caption_report(caption_result)
        print(caption_report)

    # AI痕迹清除
    if args.remove_ai:
        doc, ai_count = remove_ai_traces(doc)
        actions_performed.append(f"AI痕迹清除: 修改{ai_count}个段落")

    # 三线表转换
    if args.three_line:
        tbl_count = convert_all_tables_to_three_line(doc)
        actions_performed.append(f"三线表转换: {tbl_count}个表格")

    # 主排版（始终执行）
    doc, summary = format_document(doc, config)
    actions_performed.append("格式排版完成")

    # 6. 保存
    output_path = args.output or generate_output_path(args.input)
    success, result = safe_save_document(doc, output_path, overwrite=args.overwrite)

    if success:
        print("✅ 排版完成！")
        print(f"   输出文件: {result}")
        print()
        print(format_summary_text(summary))
        if actions_performed:
            print()
            print("执行的操作:")
            for action in actions_performed:
                print(f"  - {action}")
    else:
        print(f"❌ 保存失败: {result}")
        sys.exit(1)


if __name__ == "__main__":
    main()
