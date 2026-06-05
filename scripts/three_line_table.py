#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""三线表转换模块"""

from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def convert_to_three_line_table(table, font_name='宋体', font_size=Pt(10.5)):
    """将表格转换为三线表格式

    三线表规范：
    - 顶线：1.5磅
    - 表头下线：0.75磅
    - 底线：1.5磅
    - 无竖线
    - 表头居中，数据默认左对齐

    Args:
        table: Table 对象
        font_name: 表内字体
        font_size: 表内字号

    Returns:
        table 对象
    """
    # 1. 设置表格宽度为100%
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:type'), 'pct')
    tblW.set(qn('w:w'), '5000')
    existing_tblW = tblPr.find(qn('w:tblW'))
    if existing_tblW is not None:
        tblPr.remove(existing_tblW)
    tblPr.append(tblW)

    # 2. 删除现有边框
    existing_borders = tblPr.find(qn('w:tblBorders'))
    if existing_borders is not None:
        tblPr.remove(existing_borders)

    # 3. 创建三线表边框
    tblBorders = OxmlElement('w:tblBorders')

    # 顶线：1.5磅
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '12')  # 1.5磅 = 12半磅
    top.set(qn('w:color'), '000000')
    top.set(qn('w:space'), '0')
    tblBorders.append(top)

    # 底线：1.5磅
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:color'), '000000')
    bottom.set(qn('w:space'), '0')
    tblBorders.append(bottom)

    # 左边框：无
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'none')
    left.set(qn('w:sz'), '0')
    left.set(qn('w:color'), 'auto')
    left.set(qn('w:space'), '0')
    tblBorders.append(left)

    # 右边框：无
    right = OxmlElement('w:right')
    right.set(qn('w:val'), 'none')
    right.set(qn('w:sz'), '0')
    right.set(qn('w:color'), 'auto')
    right.set(qn('w:space'), '0')
    tblBorders.append(right)

    # 内部水平线：0.75磅
    # 内部水平线：无（三线表只在表头下方有一条线，数据行之间无线）
    insideH = OxmlElement('w:insideH')
    insideH.set(qn('w:val'), 'none')
    insideH.set(qn('w:sz'), '0')
    insideH.set(qn('w:color'), 'auto')
    insideH.set(qn('w:space'), '0')
    tblBorders.append(insideH)

    # 内部垂直线：无
    insideV = OxmlElement('w:insideV')
    insideV.set(qn('w:val'), 'none')
    insideV.set(qn('w:sz'), '0')
    insideV.set(qn('w:color'), 'auto')
    insideV.set(qn('w:space'), '0')
    tblBorders.append(insideV)

    tblPr.append(tblBorders)

    # 3.5 清除所有单元格的独立边框（否则单元格级边框会覆盖表格级设置）
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is not None:
                existing_tcBorders = tcPr.find(qn('w:tcBorders'))
                if existing_tcBorders is not None:
                    tcPr.remove(existing_tcBorders)

    # 4. 设置表头行（第一行）
    if table.rows:
        first_row = table.rows[0]
        for cell in first_row.cells:
            # 表头居中
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 设置表头下边框0.75磅
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            bottom_border = OxmlElement('w:bottom')
            bottom_border.set(qn('w:val'), 'single')
            bottom_border.set(qn('w:sz'), '6')
            bottom_border.set(qn('w:color'), '000000')
            bottom_border.set(qn('w:space'), '0')
            tcBorders.append(bottom_border)
            tcPr.append(tcBorders)

        # 设置表头跨页重复
        first_tr = first_row._tr
        trPr = first_tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        trPr.append(tblHeader)

    # 5. 统一设置表内文字格式
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.name = 'Times New Roman'
                    rPr = run._element.find(qn('w:rPr'))
                    if rPr is None:
                        rPr = OxmlElement('w:rPr')
                        run._element.insert(0, rPr)
                    rFonts = rPr.find(qn('w:rFonts'))
                    if rFonts is None:
                        rFonts = OxmlElement('w:rFonts')
                        rPr.insert(0, rFonts)
                    rFonts.set(qn('w:eastAsia'), font_name)
                    run.font.size = font_size
                    run.font.color.rgb = RGBColor(0, 0, 0)

    return table


def convert_all_tables_to_three_line(doc, font_name='宋体', font_size=Pt(10.5)):
    """将文档中所有表格转换为三线表

    Returns:
        转换的表格数量
    """
    count = 0
    for table in doc.tables:
        convert_to_three_line_table(table, font_name, font_size)
        count += 1
    return count


def set_table_allow_break(table):
    """允许表格跨页断行（防止表格被推到下一页导致大片空白）"""
    for row in table.rows:
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        cantSplit = trPr.find(qn('w:cantSplit'))
        if cantSplit is not None:
            trPr.remove(cantSplit)


def set_table_header_repeat(table):
    """设置表头在跨页时重复显示"""
    if table.rows:
        tr = table.rows[0]._tr
        trPr = tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        trPr.append(tblHeader)
