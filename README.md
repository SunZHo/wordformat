# Wordformat

**中文公文与技术报告版式技能** —— 让 AI 助手直接产出符合规范的 `.docx`，而不是把内容和格式一起糊出来。

<img src="assets/cover-preview.png" width="620" alt="由本技能生成的封面效果">

## 为什么需要它

让 AI 写一份可研报告或需求分析报告，内容通常没问题，卡住的是版式：封面无页码、目录用罗马数字、正文从 1 重排；正文行距是**固定 29 磅**而不是 1.5 倍；首行缩进按**字符**算而不是磅；表格无文字环绕、表头跨页重复、合计行加粗；表题在上、图题在下，编号跟着章号走。这些在 Word 里手调要几十次点击，换一份新文档又得重来一遍。

Wordformat 把这些规则固化成一份 Python 排版库：**内容归内容，版式归版式**。调用方只管往章节里填字，其余的页边距、页码、字体、行距、题注编号由代码保证。

## 规范覆盖

| 项目 | 落地值 |
|---|---|
| 页面 | A4，页边距上 35 / 下 32 / 左 28 / 右 26 mm |
| 页码 | 封面单独成节、无页码；目录罗马数字 `Ⅰ Ⅱ Ⅲ`；正文阿拉伯数字从 `1` 起；均底部居中 |
| 封面 | 首行标题方正小标宋一号、1.5 倍行距、居中，标题与页面顶部隔一个小四字号的 1.5 倍行距；下方空四行三号；项目名称 / 申报单位 / 编制单位方正仿宋三号、首行缩进 2 字符；日期汉字写法 |
| 目录 | 标题宋体三号；条目宋体五号、固定 18 磅、段前段后 0；只到三级，缩进 2 / 3.5 / 5 字符，前两级加粗 |
| 标题 | 三号；一级方正黑体、二级方正楷体、三级与四级方正仿宋；自动挂大纲级别 |
| 正文 | 方正仿宋三号、首行缩进 2 字符（字符单位）、**固定 29 磅**、段前段后 0 |
| 表格 | 宋体五号、无缩进、固定 18 磅；无文字环绕、允许跨页断行、表头跨页重复、单元格垂直居中、表头与合计加粗 |
| 题注 | 表题在上、图题在下；编号「章号-序号」，编号与题名空一个字符；五号黑体、1.5 倍行距、居中；子项分层时编为 `表3-1-1` |
| 文件名 | `项目名称_文件名称.docx`；封面标题按「单位前缀 + 文件类型」自动拼 |

## 安装

### 方式一：克隆到技能目录（推荐，便于 `git pull` 更新）

```bash
# WorkBuddy / CodeBuddy
git clone https://github.com/SunZHo/wordformat.git ~/.workbuddy/skills/wordformat

# Claude Code
git clone https://github.com/SunZHo/wordformat.git ~/.claude/skills/wordformat
```

### 方式二：下载 ZIP

不想装 git 的，直接下发布包：[**wordformat-1.0.0.zip**](https://github.com/SunZHo/wordformat/releases/latest)，解压后把 `wordformat/` 整个目录放进技能目录即可。

也可以在仓库页点 **Code → Download ZIP**，效果相同。

### 依赖

```bash
pip install python-docx Pillow
```

`python-docx` 用于写文档，`Pillow` 只在 `demo.py` 生成示例配图时用到。

### 其他助手

本仓库遵循通用的 `SKILL.md` 约定（YAML frontmatter + Markdown 说明 + 可执行脚本）。不支持技能目录的助手，可以把 `SKILL.md` 全文塞进系统提示，脚本路径按实际位置改一下即可。

## 快速上手

```python
import sys; sys.path.insert(0, "scripts")
import docformat as G

# 文件名下划线之后那段即"文件类型"；封面标题 = 前缀 + 文件类型
d = G.DocFormat(G.doc_filename("某某项目", "需求分析报告"),
                project="某某项目建设",
                applicant="××单位",
                compiler="××单位",
                date_cn=G.cn_date(2026, 11),
                cover_org="某某项目建设")   # 封面首行标题前缀，可省

d.cover()        # 封面（独立成节，无页码）
d.toc()          # 目录页（罗马数字页码）
d.start_body()   # 正文（阿拉伯数字从 1 起）

d.h1("第一章 项目概述")
d.p("本项目面向……")
d.h2("1.1 建设背景")
d.p("……")
d.table("建设内容与规模一览表", ["序号", "建设内容", "规模"],
        [["1", "排版能力库", "1 套"], ["合计", "—", "1 套"]],
        total=True, widths_cm=[2, 8, 3])
d.figure("chart.png", "总体架构图", width_cm=14)
d.save()
```

表题、图题、章节序号全部自动编号，不用手写 `表3-1` 这类前缀；`h1` 会解析标题里的「第X章」来决定章号，从中间章节起排也不会错号。

### 重排已有文档

```bash
# 纯文本重排（重建结构、重编号，图片不迁移）
python scripts/reformat.py 源.docx 某某项目_需求分析报告.docx \
  --project 某某项目建设 --applicant 申报单位 --compiler 编制单位 --date 二〇二六年十一月

# 图文手册重排（保留截图，按版心宽重排后重新插图）
python scripts/manual_reformat.py 源手册.docx 输出.docx \
  --software "某某平台V1.0" --applicant 申报单位 --compiler 编制单位 --date 二〇二六年十一月
```

`manual_reformat.py` 专为带截图的操作手册准备：自动识别封面、目录、正文起点，重建三节页码结构，把内嵌图片导出后按 15.6 cm 版心重新插入，**图号沿用原稿**，正文里「如图 3-2 所示」的交叉引用不会失效。

## 核心 API

| 调用 | 作用 |
|---|---|
| `DocFormat(path, project, applicant, compiler, date_cn, cover_title=None, doc_name=None, cover_org=COVER_ORG)` | 建文档并设 A4、页边距与正文样式；封面标题缺省时按「前缀 + 文件类型」拼，两者都拿不到会直接抛错而不是留空 |
| `.cover()` | 封面：小标宋一号标题 + 四行三号空距 + 信息行 + 汉字日期 |
| `.toc()` | 目录页：宋体三号标题、TOC1~3 样式、罗马页码 |
| `.start_body()` | 开正文节：阿拉伯页码从 1 起、页脚居中 |
| `.h1 / .h2 / .h3 / .h4(text)` | 三号标题，黑体 / 楷体 / 仿宋 / 仿宋，自动挂大纲级别 |
| `.p(text)` | 正文：仿宋三号、首行缩进 2 字符、固定 29 磅、两端对齐 |
| `.table(caption, header, rows, total=False, widths_cm=None, sub=None)` | 表题自动编号；`sub=1` 出 `表3-1-1`；表头跨页重复、合计行加粗 |
| `.figure(path, caption, width_cm=15)` | 图题自动编号，置于图下 |
| `.save()` | 落盘 |

字体名与字号以常量形式暴露（`FONT_XIAOBIAOSONG`、`FONT_FANGSONG`、`FONT_HEITI`、`FONT_KAITI`、`FONT_SONG`，`SIZE_1=26`、`SIZE_3=16`、`SIZE_4S=12`、`SIZE_5=10.5`），目标机器字体名不同时改常量即可。

## 自检

交付前跑一遍，比肉眼翻页可靠：

```bash
python scripts/verify_docx.py 输出.docx
```

覆盖九类可机检项：页边距、三节页码格式与起始值、目录标题与条目样式、封面字体字号、标题层级字体、正文缩进与固定行距、表格属性（环绕 / 跨页 / 垂直居中 / 字号字体）、题注编号连续性、文件名格式。全部通过输出 `OK`，否则逐条列出不符合项。

## 目录结构

```
wordformat/
├── SKILL.md                    技能说明：触发条件、工作流、硬性约束
├── references/format-spec.md   规范逐条口径 + 交付检查清单 + 常见坑
├── scripts/
│   ├── docformat.py            排版库，所有格式规则都在这里
│   ├── verify_docx.py          格式自检
│   ├── reformat.py             纯文本重排
│   ├── manual_reformat.py      图文手册重排（保留截图）
│   └── demo.py                 生成样例文档
└── assets/cover-preview.png
```

## 已知限制

- **目录是 Word 域**，页码在打开时计算。`updateFields` 会在打开时询问是否更新域，选「是」即可；想让别人打开即见页码，需在有 Word 的机器上打开保存一次。
- **方正字库需自行安装**。未安装时系统会回退显示（多为宋体），但文件里的字体名已经写对，在装有方正字库的机器上打开即为规范效果。
- 无 Word / LibreOffice 的机器上只能做 XML 级校验，`verify_docx.py` 已覆盖全部可机检项，但版式观感仍需人工看一眼。
- `reformat.py` 会丢弃图片（结构重建所致），带截图的手册请用 `manual_reformat.py`。

## 说明

版式取自中文公文与技术报告的通用排版要求：三节页码制度、固定行距、字符缩进、题注随章编号。字体与页边距按同一套口径实现，换单位时改 `COVER_ORG` 或传 `cover_org`、`cover_title` 即可。

## License

[MIT](LICENSE)
