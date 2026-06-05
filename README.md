# mxd-pd — 论文排版 Claude Code Skill

一个用于论文排版的 Claude Code Skill，支持自动排版、格式检查、图表序号处理、AI痕迹清除、三线表转换等功能。

## 功能特性

- 📝 **交互式排版引导** — 分步询问格式要求，支持数字快速选择
- 📊 **图表序号自动处理** — 检测缺失的图序图题/表序表题，根据上下文自动生成
- 🔢 **序号连续性检查** — 自动检测图/表序号是否跳号或乱序
- 🖼️ **图片行距修复** — 自动处理固定行距下图片显示不全的问题
- 📋 **三线表转换** — 一键将普通表格转为学术三线表格式
- 🤖 **AI痕迹清除** — 删除项目符号、列表符号，替换AI常用表达
- ✅ **格式检查模式** — 只检查不修改，输出带位置的问题清单
- 🎨 **排版模式** — 自动修改格式，生成新文件

## 安装

```bash
# 将 mxd-pd 文件夹复制到 Claude Code 的 skills 目录
# Windows
copy mxd-pd %USERPROFILE%\.claude\skills\

# macOS/Linux
cp -r mxd-pd ~/.claude/skills/

# 安装依赖
pip install -r ~/.claude/skills/mxd-pd/requirements.txt
```

## 使用方法

### 通过 Claude Code 交互式使用

```
/mxd-pd
```

然后按照引导操作：选择任务 → 提供文件 → 选择规范 → 确认执行。

### 通过命令行直接使用

```bash
# 排版模式
python scripts/format_paper.py 论文.docx

# 检查模式
python scripts/format_paper.py 论文.docx --check

# 使用文史类规范
python scripts/format_paper.py 论文.docx --preset arts

# 启用所有自动功能
python scripts/format_paper.py 论文.docx --fix-captions --remove-ai --three-line

# 指定输出路径
python scripts/format_paper.py 论文.docx -o 输出.docx
```

## 项目结构

```
mxd-pd/
├── SKILL.md                # Claude Code 指令文件
├── README.md               # 本文件
├── LICENSE                 # MIT 协议
├── requirements.txt        # Python 依赖
└── scripts/
    ├── __init__.py
    ├── format_paper.py     # 主入口（CLI + 排版流程）
    ├── config.py           # 配置模块（字号表、默认规范、预设）
    ├── validator.py        # 文件验证与安全保存
    ├── formatter.py        # 核心排版逻辑（段落、标题、正文、表格）
    ├── captions.py         # 图表序号检测、补全、重编号
    ├── ai_traces.py        # AI痕迹清除（项目符号、列表符号、AI表达）
    ├── three_line_table.py # 三线表转换、跨页断行
    └── checker.py          # 格式检查（只报告不修改）
```

## 支持的格式选项

### 标题格式
- 字体：黑体、宋体、仿宋、楷体
- 字号：三号、小三、四号、小四、五号
- 编号：理工类（1.→3.1→3.1.1）、文史类（一、→（一）→1.→（1））

### 正文格式
- 字体：仿宋、宋体、楷体
- 字号：四号、小四、五号
- 行距：固定值24磅、28.8磅、20磅、1.5倍行距

### 预设规范
- `science` — 理工类（默认）
- `arts` — 文史体艺类

## 常见字号对照

| 字号 | 磅值 |
|------|------|
| 三号 | 16pt |
| 小三 | 15pt |
| 四号 | 14pt |
| 小四 | 12pt |
| 五号 | 10.5pt |
| 小五 | 9pt |

## 开源协议

MIT License — 详见 [LICENSE](LICENSE)
