---
name: mxd-pd
description: "论文排版skill。当用户需要对论文进行排版、检查格式、生成符合规范的Word文档时使用。触发词：论文排版、排版、格式规范、毕业论文、mxd-pd。"
---

# 论文排版助手

## 核心工作流程

**始终使用交互式引导，不要假设任何信息。**

```
用户输入 /mxd-pd
    ↓
询问：要做什么？（排版 / 检查格式）
    ↓
询问：论文文件路径？
    ↓
询问：有排版规范文件吗？
    ├─ 有 → 读取规范文件
    └─ 没有 → 逐步询问格式要求
    ↓
确认规范 → 执行排版/检查
```

## 第一步：确认任务

使用AskUserQuestion询问：

```
你要做什么？
1. 排版论文（修改格式）
2. 检查格式（只报告问题）
```

## 第二步：确认论文文件

```
请提供论文文件路径：
- 可以是 .docx 文件
- 也可以是 .doc 文件（会先转换）
```

如果用户在调用时已经提供了文件路径（如 `/mxd-pd 论文.docx`），则跳过此步。

## 第三步：获取排版规范

### 询问是否有规范文件

```
你有排版规范文件吗？
1. 有，我提供文件
2. 没有，我口述要求
3. 没有，用常见学校标准
```

### 情况A：用户有规范文件

```
请提供排版规范文件路径：
```

→ 读取文件，提取格式要求

→ 提取完成后，展示给用户确认：
```
我从规范文件中提取到以下格式要求：
- 一级标题：xxx
- 二级标题：xxx
- 正文：xxx
...
是否正确？需要修改吗？
```

### 情况B：用户口述要求

分4组询问，每组用AskUserQuestion一次性展示4个问题，字体/字号/行距分开选：

**第1组：标题编号 + 一级标题**（AskUserQuestion，4个问题）

| 问题 | 选项 |
|------|------|
| 标题编号格式 | 1理工类（1.→3.1→3.1.1） 2文史类（一、→（一）→1.→（1）） |
| 一级标题字体 | 1黑体 2宋体 3仿宋 4楷体 |
| 一级标题字号/加粗 | 1三号加粗 2三号不加粗 3小三加粗 4小三不加粗 |
| 一级标题行距 | 1固定值24磅 2固定值28.8磅 3固定值20磅 4其他 |

**第2组：二级+三级+四级标题**（AskUserQuestion，4个问题）

| 问题 | 选项 |
|------|------|
| 二级标题字体/字号/加粗 | 1仿宋小三加粗 2仿宋四号加粗 3宋体四号加粗 4其他 |
| 三级标题字体/字号/加粗 | 1仿宋四号加粗 2宋体小四加粗 3黑体四号加粗 4其他 |
| 四级标题字体/字号/加粗 | 1仿宋小四加粗 2宋体小四加粗 3黑体小四加粗 4与正文相同 |
| 首行缩进 | 1缩进2字符 2缩进1字符 3不缩进 4其他 |

**第3组：正文格式**（AskUserQuestion，4个问题）

| 问题 | 选项 |
|------|------|
| 正文字体 | 1仿宋 2宋体 3楷体 |
| 正文字号 | 1四号 2小四 3五号 |
| 正文行距 | 1固定值24磅 2固定值28.8磅 3固定值20磅 41.5倍行距 |
| 英文/数字字体 | 1Times New Roman 2Arial 3与正文字体相同 |

**第4组：摘要+关键词**（AskUserQuestion，4个问题）

| 问题 | 选项 |
|------|------|
| 中文摘要标题字体/字号 | 1黑体三号 2黑体小三 3黑体四号 4宋体三号 |
| 中文摘要内容字体/字号/行距 | 1仿宋四号28.8磅 2仿宋四号24磅 3宋体小四28.8磅 4与正文相同 |
| 英文摘要标题字体/字号 | 1Times New Roman加粗三号 2Times New Roman加粗小三 3Arial加粗三号 4与中文摘要相同 |
| 关键词格式 | 1"关键词"黑体加粗+内容仿宋 2"关键词"宋体加粗+内容宋体 3与正文相同 4其他 |

**第5组：参考文献+图表+页边距**（AskUserQuestion，4个问题）

| 问题 | 选项 |
|------|------|
| 参考文献编码方式 | 1顺序编码[1][2][3] 2作者-年份制 |
| 参考文献正文字号 | 1五号仿宋 2小五仿宋 3五号宋体 4与正文相同 |
| 图题/表题字体/字号 | 1五号宋体 2五号仿宋 3小五宋体 4与正文相同 |
| 页边距 | 1上下2.54cm左右3.17cm 2上下2.5cm左右2.5cm 3自定义 |

### 情况C：用常见标准

```
使用哪种标准？
1. 理工类（1. → 3.1 → 3.1.1）
2. 文史体艺类（一、 → （一） → 1. → （1））
```

→ 应用对应的默认规范，展示给用户确认

## 第四步：确认并执行

展示最终规范摘要：

```
排版规范确认：
━━━━━━━━━━━━━━
论文文件：xxx.docx
操作类型：排版 / 检查
标题格式：理工类 / 文史类
一级标题：三号黑体
二级标题：小三仿宋加粗
正文：四号仿宋，行距24磅
...
━━━━━━━━━━━━━━
确认无误？开始执行？
```

用户确认后执行。

## 第五步：执行

### 排版模式
- 读取论文文件
- 按规范逐段修改格式
- **自动检测图表序号并补全图序图题/表序表题**
- 生成新文件：`论文_已排版.docx`
- 输出修改摘要

### 检查模式
- 读取论文文件
- 对照规范逐项检查
- **检测图表序号是否连续、是否有图序图题**
- 输出问题清单（带行号和具体问题）

### 图表序号机制（自动处理）

**功能：**
1. **检测缺失**：图片没有图序图题、表格没有表序表题时，自动补全
2. **序号检查**：检测图1、图2、图3...是否连续，表1、表2、表3...是否连续
3. **自动修复**：序号不连续时重新编号，缺失标题时根据上下文生成

**处理流程：**
```
1. 遍历文档，找到所有图片和表格
2. 检查每个图片/表格是否有对应的图序图题/表序表题
3. 检查序号是否连续（不跳号、不重复、不乱序）
4. 对于缺失标题的图表：
   - 读取上下文段落内容
   - 分析图表主题
   - 生成简洁准确的图题/表题
5. 对于序号错误的图表：
   - 按出现顺序重新编号
```

**图序图题格式：**
- 位置：图片正下方居中
- 格式：`图1 图题内容`（图序与图题之间空一格）
- 字体：按用户选择的图题格式

**表序表题格式：**
- 位置：表格正上方居中
- 格式：`表1 表题内容`（表序与表题之间空一格）
- 字体：按用户选择的表题格式

## 常见字号对照

| 字号 | 磅值 |
|------|------|
| 初号 | 42pt |
| 小初 | 36pt |
| 一号 | 26pt |
| 小一 | 24pt |
| 二号 | 22pt |
| 小二 | 18pt |
| 三号 | 16pt |
| 小三 | 15pt |
| 四号 | 14pt |
| 小四 | 12pt |
| 五号 | 10.5pt |
| 小五 | 9pt |

## Python工具参考

```python
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn

doc = Document('论文.docx')

# 设置中文字体（需同时设置ascii和eastAsia）
run.font.name = 'Times New Roman'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')

# 设置字号
run.font.size = Pt(14)

# 设置行距（固定值）
paragraph.paragraph_format.line_spacing = Pt(24)

# 设置首行缩进
paragraph.paragraph_format.first_line_indent = Cm(0.74)

# 设置加粗
run.bold = True

# 设置对齐
paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER  # 居中
paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY  # 两端对齐

# ========== 图片段落行距处理（重要！） ==========
# 固定行距会导致嵌入式图片显示不全，图片段落需设为单倍行距
from docx.oxml.ns import qn

def has_image(paragraph):
    """检查段落是否包含图片"""
    return bool(paragraph._element.findall('.//' + qn('wp:inline')) +
                paragraph._element.findall('.//' + qn('wp:anchor')))

def fix_image_paragraphs(doc):
    """将所有包含图片的段落设为单倍行距"""
    for para in doc.paragraphs:
        if has_image(para):
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            # 或者设为固定值，值略大于图片高度
            # para.paragraph_format.line_spacing = Pt(图片高度+10)

# 使用示例
fix_image_paragraphs(doc)

# ========== 表格跨页断行处理 ==========
# 长表格跨页时，可能需要设置跨页断行属性
from docx.oxml import OxmlElement

def set_table_allow_break(table):
    """允许表格跨页断行（防止表格被推到下一页导致大片空白）"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    # 设置 cantSplit 为 false（允许断行）
    for tr in tbl.tr_elements:
        trPr = tr.get_or_add_trPr()
        cantSplit = trPr.find(qn('w:cantSplit'))
        if cantSplit is not None:
            trPr.remove(cantSplit)

def set_table_header_repeat(table):
    """设置表头在跨页时重复显示（三线表常用）"""
    # 只对第一行设置
    if table.rows:
        tr = table.rows[0]._tr
        trPr = tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        trPr.append(tblHeader)

# 使用示例
for table in doc.tables:
    set_table_allow_break(table)
    set_table_header_repeat(table)  # 可选：表头跨页重复

# ========== 图表序号检测与自动补全 ==========
import re
from docx.shared import Pt

def find_image_paragraphs(doc):
    """找到所有包含图片的段落及其索引"""
    image_paras = []
    for i, para in enumerate(doc.paragraphs):
        if has_image(para):
            image_paras.append((i, para))
    return image_paras

def find_table_positions(doc):
    """找到所有表格的位置（在段落列表中的插入点）"""
    table_positions = []
    # 表格在XML中与段落同级，需要通过body元素遍历
    body = doc.element.body
    for i, child in enumerate(body):
        if child.tag == qn('w:tbl'):
            table_positions.append(i)
    return table_positions

def get_context_text(doc, para_index, direction='before', count=3):
    """获取指定段落的上下文文本"""
    texts = []
    if direction == 'before':
        start = max(0, para_index - count)
        for i in range(start, para_index):
            if doc.paragraphs[i].text.strip():
                texts.append(doc.paragraphs[i].text.strip())
    elif direction == 'after':
        end = min(len(doc.paragraphs), para_index + count + 1)
        for i in range(para_index + 1, end):
            if doc.paragraphs[i].text.strip():
                texts.append(doc.paragraphs[i].text.strip())
    return texts

def check_figure_caption(doc):
    """检查图片是否有图序图题，返回缺失列表"""
    image_paras = find_image_paragraphs(doc)
    missing_captions = []
    figure_nums = []

    for i, (para_idx, para) in enumerate(image_paras):
        # 检查图片下方是否有图序图题（通常是下一个段落）
        has_caption = False
        if para_idx + 1 < len(doc.paragraphs):
            next_para = doc.paragraphs[para_idx + 1].text.strip()
            # 检查是否匹配"图X"或"图 X"格式
            match = re.match(r'^图\s*(\d+)\s*', next_para)
            if match:
                has_caption = True
                figure_nums.append(int(match.group(1)))

        if not has_caption:
            context = get_context_text(doc, para_idx, 'before', 2)
            missing_captions.append({
                'para_index': para_idx,
                'context': context,
                'has_caption': False
            })

    return missing_captions, figure_nums

def get_paragraph_text(para_elem):
    """从w:p元素中正确提取文本（文本在w:r/w:t中）"""
    texts = []
    for r_elem in para_elem.findall('.//' + qn('w:r')):
        for t_elem in r_elem.findall(qn('w:t')):
            if t_elem.text:
                texts.append(t_elem.text)
    return ''.join(texts)

def check_table_caption(doc):
    """检查表格是否有表序表题，返回缺失列表"""
    table_positions = find_table_positions(doc)
    missing_captions = []
    table_nums = []

    for i, tbl_pos in enumerate(table_positions):
        # 检查表格上方是否有表序表题（通常是前一个段落）
        has_caption = False
        # 找到表格前的段落
        body = doc.element.body
        if tbl_pos > 0:
            prev_elem = body[tbl_pos - 1]
            if prev_elem.tag == qn('w:p'):
                # 正确获取段落文本（文本在w:r/w:t中，不在p.text上）
                prev_text = get_paragraph_text(prev_elem)
                # 检查是否匹配"表X"或"表 X"格式
                match = re.match(r'^表\s*(\d+)\s*', prev_text.strip())
                if match:
                    has_caption = True
                    table_nums.append(int(match.group(1)))

        if not has_caption:
            missing_captions.append({
                'table_index': i,
                'has_caption': False
            })

    return missing_captions, table_nums

def check_sequence(nums):
    """检查序号是否连续，返回错误信息列表（空列表表示无错误）"""
    if not nums:
        return []
    errors = []
    expected = list(range(1, len(nums) + 1))
    if nums != expected:
        for i, (actual, exp) in enumerate(zip(nums, expected)):
            if actual != exp:
                errors.append(f"第{i+1}个应为{exp}，实际为{actual}")
    return errors

def generate_caption(context_texts, is_figure=True):
    """根据上下文生成图题/表题"""
    prefix = "图" if is_figure else "表"
    # 简单策略：取上下文中的关键词
    if context_texts:
        # 取最近的一段非空文本，截取前20字作为标题
        text = context_texts[-1][:20]
        return f"{text}示意图" if is_figure else f"{text}情况表"
    return f"{prefix}题待补充"

def insert_paragraph_after(doc, para_index, text, font_name='宋体', font_size=Pt(10.5), bold=False, alignment=WD_ALIGN_PARAGRAPH.CENTER):
    """在指定段落之后插入新段落"""
    body = doc.element.body
    # 创建新段落元素
    new_para = OxmlElement('w:p')
    # 创建run
    new_run = OxmlElement('w:r')
    new_rpr = OxmlElement('w:rPr')
    # 设置字体
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    new_rpr.append(rFonts)
    # 设置字号
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(font_size.pt * 2)))  # half-points
    new_rpr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(int(font_size.pt * 2)))
    new_rpr.append(szCs)
    # 设置加粗
    if bold:
        b = OxmlElement('w:b')
        new_rpr.append(b)
    new_run.append(new_rpr)
    # 设置文本
    new_t = OxmlElement('w:t')
    new_t.text = text
    new_run.append(new_t)
    new_para.append(new_run)
    # 设置对齐
    new_ppr = OxmlElement('w:pPr')
    new_jc = OxmlElement('w:jc')
    jc_val = 'center' if alignment == WD_ALIGN_PARAGRAPH.CENTER else 'left'
    new_jc.set(qn('w:val'), jc_val)
    new_ppr.append(new_jc)
    new_para.insert(0, new_ppr)
    # 找到目标段落并插入
    target_elem = doc.paragraphs[para_index]._element
    target_elem.addnext(new_para)
    return new_para

def fix_figure_table_captions(doc, figure_format, table_format):
    """
    自动补全图序图题、表序表题，并修复序号

    参数:
        doc: Document对象
        figure_format: 图题格式 {'font': '宋体', 'size': Pt(10.5), 'bold': False}
        table_format: 表题格式 {'font': '宋体', 'size': Pt(10.5), 'bold': False}
    """
    # 1. 检查并修复图序图题
    missing_figures, figure_nums = check_figure_caption(doc)
    seq_errors_fig = check_sequence(figure_nums)

    if missing_figures or seq_errors_fig:
        print(f"发现 {len(missing_figures)} 个图片缺失图序图题")
        if seq_errors_fig:
            print(f"图序错误: {seq_errors_fig}")

        # 重新编号所有图片
        image_paras = find_image_paragraphs(doc)
        for i, (para_idx, para) in enumerate(image_paras):
            fig_num = i + 1
            # 检查下方是否已有图题
            if para_idx + 1 < len(doc.paragraphs):
                next_para = doc.paragraphs[para_idx + 1]
                if re.match(r'^图\s*\d+\s*', next_para.text.strip()):
                    # 更新序号
                    new_text = re.sub(r'^图\s*\d+', f'图{fig_num}', next_para.text)
                    for run in next_para.runs:
                        run.text = ''
                    next_para.runs[0].text = new_text
                else:
                    # 在图片段落之后插入新图题
                    context = get_context_text(doc, para_idx, 'before', 2)
                    caption_text = generate_caption(context, is_figure=True)
                    insert_paragraph_after(
                        doc, para_idx,
                        f'图{fig_num} {caption_text}',
                        font_name=figure_format.get('font', '宋体'),
                        font_size=figure_format.get('size', Pt(10.5)),
                        bold=figure_format.get('bold', False)
                    )

    # 2. 检查并修复表序表题
    missing_tables, table_nums = check_table_caption(doc)
    seq_errors_tbl = check_sequence(table_nums)

    if missing_tables or seq_errors_tbl:
        print(f"发现 {len(missing_tables)} 个表格缺失表序表题")
        if seq_errors_tbl:
            print(f"表序错误: {seq_errors_tbl}")

    return {
        'missing_figures': len(missing_figures),
        'missing_tables': len(missing_tables),
        'figure_seq_errors': seq_errors_fig,
        'table_seq_errors': seq_errors_tbl
    }

# 使用示例
figure_format = {'font': '宋体', 'size': Pt(10.5), 'bold': False}
table_format = {'font': '宋体', 'size': Pt(10.5), 'bold': False}
result = fix_figure_table_captions(doc, figure_format, table_format)
print(f"缺失图题: {result['missing_figures']}个")
print(f"缺失表题: {result['missing_tables']}个")

doc.save('论文_已排版.docx')
```
