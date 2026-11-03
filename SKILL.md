---
name: wordformat
description: Wordformat —— 按中文公文与技术报告版式规范生成或重排 Word 文档，采用方正字库、35/32/28/26mm 页边距、方正小标宋封面、宋体目录、方正仿宋三号固定 29 磅正文、宋体五号表格与表/图题注自动编号。当需要撰写或排版可研报告、需求分析报告、项目建议书、技术方案等规范格式 docx，或要求"按公文格式""封面无页码/目录罗马数字/正文阿拉伯数字/固定 29 磅行距"时使用；提到 Wordformat 或要求"Word 排版规范化"时同样适用。
agent_created: true
display_name: Wordformat
---

# Wordformat · 公文与技术报告版式

## 用途

本技能标识 `wordformat`，名称 Wordformat。把"写内容"与"套格式"分开：内容交给调用方，版式全部由 `scripts/docformat.py` 落地。产出可直接送审的 `.docx`——页边距、页码制度、字体字号、缩进行距、表格属性、表图题注编号全部符合规范。

完整口径与逐条核对清单见 `references/format-spec.md`，交付前必须按该清单收口。

## 何时使用

- 生成可研报告、需求分析报告、项目建议书、技术方案、投资估算书等规范格式 docx。
- 把一份已有 docx 重排为规范格式。
- 用户提到"公文格式""方正小标宋""固定 29 磅""目录罗马数字页码"等要求。

## 工作流

1. **确认四要素**：项目名称、项目申报单位、编制单位、日期（汉字数字，如"二〇二六年十一月"）。缺失时按已有信息推断并说明假设，不留空字段。
2. **定封面标题**：首行标题不是固定文案，由"前缀 + 文件类型"组成（如"××项目建设需求分析报告"）。前缀默认取常量 `COVER_ORG`（为空则只用文件类型），文件类型取输出文件名 `项目名称_文件名称.docx` 下划线之后的部分；前缀或标题另有约定时，分别传 `cover_org`、`cover_title` 覆盖。
3. **准备内容**：章节层级不超过四层。一级写"第一章 ×××"，二级"1.1 ×××"，三级"1.1.1 ×××"。
4. **调库生成**：导入 `scripts/docformat.py`，按"封面 → 目录 → 正文"顺序构建，表格与图按出现顺序调用即自动编号。
5. **自检**：运行 `python scripts/verify_docx.py 输出.docx`，有任何"不符合"就修，直到输出 `OK`。
6. **交付**：用 `present_files` 打开文档，并提醒两件事：目标机需安装方正字库；打开时选择"是"以刷新目录域页码。

```python
import sys; sys.path.insert(0, "<skill>/scripts")
import docformat as G

# 文件名第二段即文件类型，配合 cover_org 拼成“××项目建设需求分析报告”
d = G.DocFormat(G.doc_filename("项目名称", "需求分析报告"),
                project="××项目建设", applicant="××单位", compiler="××单位",
                date_cn=G.cn_date(2026, 11), cover_org="××项目建设")
# 标题另有写法时整串覆盖：cover_title="××项目建设可行性研究报告"
d.cover()          # 封面（独立成节，无页码）
d.toc()            # 目录（罗马数字页码，仅三级）
d.start_body()     # 正文（阿拉伯数字，从 1 起）
d.h1("第一章 项目概述"); d.p("……")
d.h2("1.1 项目背景"); d.p("……")
d.table("建设内容一览表", ["序号", "内容"], [["1", "……"]], total=True, widths_cm=[2, 10])
d.figure("chart.png", "总体架构图", width_cm=14)
d.save()
```

已有文档重排：

```bash
python scripts/reformat.py 源.docx 项目名称_需求分析报告.docx \
  --project 项目名称 --applicant 申报单位 --compiler 编制单位 --date 二〇二六年十一月
# 输出文件名不含下划线时补 --doc-name 需求分析报告，或用 --cover-title 直接指定整串标题
```

重排会重建结构并重新编号表题，图片不迁移（脚本给出数量提示），需要时用 `figure()` 重新插入。

带截图的图文手册（操作手册、使用说明）用另一个脚本，图片会导出后按规范重新插入：

```bash
python scripts/manual_reformat.py 源手册.docx 输出.docx \
  --software "××平台V1.0" --applicant 申报单位 --compiler 编制单位 --date 二〇二六年十一月
```

- 自动识别封面（软件全称、编制单位）、目录、正文起始位置，重建三节结构与页码。
- 内嵌截图导出后按版心宽 15.6 cm 重排，图题沿用原稿文字，**图号与原稿一致**，正文里"如图 3-2 所示"的交叉引用不会失效。
- 表格重建为规范样式并自动补表题：小节内的表取小节名，直接挂在章下的表取首列表头（如"操作角色一览表"）。
- 运行后逐项核对：图题数=图片数、表题数=表格数，再跑 `verify_docx.py`。

## 核心 API

| 调用 | 作用 |
|---|---|
| `DocFormat(path, project, applicant, compiler, date_cn, cover_title=None, doc_name=None, cover_org=COVER_ORG)` | 建文档并设 A4、35/32/28/26mm、正文样式；封面标题缺省时按"前缀+文件名第二段"自动拼 |
| `.cover()` | 封面：小标宋一号标题 + 4 行三号空距 + 信息行 + 汉字日期 |
| `.toc()` | 目录页：宋体三号标题、TOC1~3 样式、罗马页码 |
| `.start_body()` | 开正文节：阿拉伯页码从 1 起、页脚居中 |
| `.h1/.h2/.h3/.h4(text)` | 三号标题，黑体/楷体/仿宋/仿宋，自动挂大纲级别 |
| `.p(text)` | 正文：仿宋三号、首行缩进 2 字符、固定 29 磅、两端对齐 |
| `.table(caption, header, rows, total=False, widths_cm=None, sub=None)` | 表题自动编号（表X-Y，`sub=1` 得表X-1-1），表头跨页重复、合计加粗 |
| `.figure(path, caption, width_cm=15)` | 图题自动编号（图X-Y），置于图下 |
| `.save()` | 落盘 |

字号常量 `SIZE_1=26 / SIZE_3=16 / SIZE_4S=12 / SIZE_5=10.5`；字体常量 `FONT_XIAOBIAOSONG / FONT_FANGSONG / FONT_HEITI / FONT_KAITI / FONT_SONG`，目标机字体名不同时改这些常量即可。

## 硬性约束（违反即返工）

- 页边距上 35、下 32、左 28、右 26 mm；每个新增 `section` 都要重设。
- 封面首行标题随文件类型走，不写死成某一种报告名；标题恒定一号字，过长按规范自然折行，不缩字号。
- 封面单独成节且页脚为空；目录罗马数字；正文阿拉伯数字且 `start=1`；均底部居中。
- 正文固定行距 **29 磅**（不是 1.5 倍）；表格与目录固定 **18 磅**。
- 首行缩进必须用字符单位（`w:firstLineChars="200"`），不能用磅近似。
- 表格必须无文字环绕、允许跨页断行、表头跨页重复、单元格垂直居中、表头与合计加粗。
- 表题在上、图题在下，编号"章号-序号"，编号与题名之间空一个字符，同章内连续不跳号。
- 每幅图在正文必须被明确提及；图宽 ≤ 15.6 cm。
- 文件名"项目名称\_文件名称.docx"。

## 环境

依赖 `python-docx`（写 docx）与 `Pillow`（`demo.py` 生成示例配图用）：

```bash
pip install python-docx Pillow
```

生成样例验证环境：

```bash
python scripts/demo.py 示例项目_可行性研究报告.docx
```

`scripts/demo.py` 生成含封面、目录、两章正文、两张表、一张图的样例，可直接验证规范落地效果。

脚本分工：`docformat.py` 排版库、`demo.py` 样例、`reformat.py` 纯文本重排（丢图）、`manual_reformat.py` 图文手册重排（保图）、`verify_docx.py` 自检。

## 已知限制

- 目录是 Word 域，页码由 Word 打开时计算；`updateFields=true` 会在打开时询问是否更新域，属正常现象。若要求"打开即见页码"，需在装有 Word 的机器上打开保存一次。
- 方正字库未安装时按系统回退显示（多为宋体）。文件内字体名已写对，在装有对应字库的机器上打开即为规范效果，交付时提醒用户核对。
- 本机无 Word/LibreOffice 时无法渲染截图核对版式，只能做 XML 级校验；`verify_docx.py` 已覆盖全部可机检项。
