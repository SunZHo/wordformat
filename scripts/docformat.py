#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中文公文与技术报告排版库（python-docx 封装）。

实现一套可直接送审的公文版式：
页边距上35/下32/左28/右26mm；封面方正小标宋一号；正文方正仿宋三号、
首行缩进2字符、行距固定29磅；标题三号（黑体/楷体/仿宋）；表格宋体五号、
行距固定18磅、表头跨页重复；封面不编页码、目录罗马数字、正文阿拉伯数字。

用法：
    from docformat import DocFormat
    d = DocFormat("项目名称_需求分析报告.docx", project="...", applicant="...",
                  compiler="...", date_cn="二〇二六年十一月")
    d.cover()
    d.toc()
    d.start_body()
    d.h1("第一章 项目概述")
    d.p("正文……")
    d.table(caption="建设内容一览表", header=[...], rows=[[...]], total=True)
    d.figure("chart.png", caption="总体架构图", width_cm=15)
    d.save()
"""
from __future__ import annotations

import os
import re
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt

# ---------------------------------------------------------------- 常量

FONT_XIAOBIAOSONG = "方正小标宋简体"
FONT_FANGSONG = "方正仿宋_GBK"
FONT_HEITI = "方正黑体_GBK"
FONT_KAITI = "方正楷体_GBK"
FONT_SONG = "宋体"
FONT_HEI = "黑体"

SIZE_1 = 26.0    # 一号
SIZE_3 = 16.0    # 三号
SIZE_4S = 12.0   # 小四
SIZE_5 = 10.5    # 五号

BODY_LINE_PT = 29.0        # 正文行距固定 29 磅
TABLE_LINE_PT = 18.0       # 表格、目录行距固定 18 磅
MARGIN_TOP = Mm(35)
MARGIN_BOTTOM = Mm(32)
MARGIN_LEFT = Mm(28)
MARGIN_RIGHT = Mm(26)

_CN_DIGITS = "〇一二三四五六七八九"


def cn_number(num: int) -> str:
    """阿拉伯数字 -> 汉字数字（用于日期年份，如 2026 -> 二〇二六）。"""
    return "".join(_CN_DIGITS[int(c)] for c in str(num))


def cn_date(year: int, month: int) -> str:
    """生成“二〇二六年十一月”式日期。"""
    return f"{cn_number(year)}年{cn_number(month)}月"


_CN_VAL = {"〇": 0, "零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
           "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def cn_to_int(text: str) -> int | None:
    """汉字数字 -> 整数，支持“三”“十二”“二十一”等写法。"""
    if not text:
        return None
    if "十" in text:
        head, _, tail = text.partition("十")
        tens = _CN_VAL.get(head, 1) if head else 1
        ones = _CN_VAL.get(tail, 0) if tail else 0
        return tens * 10 + ones
    if len(text) == 1:
        return _CN_VAL.get(text)
    return None


def doc_filename(project: str, doc_name: str, ext: str = ".docx") -> str:
    """按“项目名称_文件名称”命名。"""
    return f"{project}_{doc_name}{ext}"


# 封面首行标题的默认前缀；为空则标题只写文件类型，需要时传 cover_org
COVER_ORG = ""


def doc_name_of(path: str) -> str:
    """从“项目名称_文件名称.docx”中取文件名称（最后一个下划线之后的部分）。"""
    stem = os.path.splitext(os.path.basename(path))[0]
    return stem.rsplit("_", 1)[1].strip() if "_" in stem else ""


def compose_cover_title(org: str, doc_name: str) -> str:
    """拼封面首行标题：前缀 + 文件类型，如“××项目建设需求分析报告”。"""
    org = (org or "").strip()
    doc_name = (doc_name or "").strip()
    if not doc_name:
        return org
    if not org:
        return doc_name
    return org + doc_name


# ---------------------------------------------------------------- XML 工具

def _set_run_fonts(run, cn_font: str, en_font: str | None = None) -> None:
    en_font = en_font or cn_font
    run.font.name = en_font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), en_font)
    rfonts.set(qn("w:hAnsi"), en_font)
    rfonts.set(qn("w:eastAsia"), cn_font)
    rfonts.set(qn("w:cs"), en_font)


def _set_char_indent(par, first_line_chars: float | None = None,
                     left_chars: float | None = None,
                     right_chars: float | None = None,
                     font_pt: float = SIZE_3) -> None:
    """按“字符”设置缩进（w:firstLineChars / w:leftChars），同时给出磅值兜底。"""
    ppr = par._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    if first_line_chars is not None:
        ind.set(qn("w:firstLineChars"), str(int(round(first_line_chars * 100))))
        ind.set(qn("w:firstLine"), str(int(round(first_line_chars * font_pt * 20))))
    if left_chars is not None:
        ind.set(qn("w:leftChars"), str(int(round(left_chars * 100))))
        ind.set(qn("w:left"), str(int(round(left_chars * font_pt * 20))))
    if right_chars is not None:
        ind.set(qn("w:rightChars"), str(int(round(right_chars * 100))))
        ind.set(qn("w:right"), str(int(round(right_chars * font_pt * 20))))


def _set_outline_level(par, level: int) -> None:
    """level=0..8 为标题层级，9 表示正文。"""
    ppr = par._p.get_or_add_pPr()
    lvl = ppr.find(qn("w:outlineLvl"))
    if lvl is None:
        lvl = OxmlElement("w:outlineLvl")
        ppr.append(lvl)
    lvl.set(qn("w:val"), str(level))


def _fixed_line(par, pt: float) -> None:
    pf = par.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(pt)


def _no_space(par) -> None:
    par.paragraph_format.space_before = Pt(0)
    par.paragraph_format.space_after = Pt(0)


def _field(par, instr: str, placeholder: str = "") -> None:
    """在段落中插入 Word 域（PAGE / TOC 等）。"""
    run = par.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run._element.append(begin)

    run2 = par.add_run()
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    run2._element.append(it)

    run3 = par.add_run()
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    run3._element.append(sep)

    if placeholder:
        par.add_run(placeholder)

    run5 = par.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run5._element.append(end)


def _set_pgnum(section, fmt: str | None, start: int | None) -> None:
    sectpr = section._sectPr
    pg = sectpr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType")
        sectpr.append(pg)
    if fmt:
        pg.set(qn("w:fmt"), fmt)
    if start is not None:
        pg.set(qn("w:start"), str(start))


def _cell_valign_center(cell) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    va = tcpr.find(qn("w:vAlign"))
    if va is None:
        va = OxmlElement("w:vAlign")
        tcpr.append(va)
    va.set(qn("w:val"), "center")


def _table_borders(table, sz: int = 4) -> None:
    """全表细边框（sz 为 1/8 磅，4 ≈ 0.5 磅）。"""
    tblpr = table._tbl.tblPr
    borders = tblpr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblpr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)


def _mark_header_row(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    if trpr.find(qn("w:tblHeader")) is None:
        trpr.append(OxmlElement("w:tblHeader"))


# ---------------------------------------------------------------- 主类

class DocFormat:
    """按公文版式规范逐段构建 Word 文档。"""

    def __init__(self, path: str, project: str = "", applicant: str = "",
                 compiler: str = "", date_cn: str = "",
                 cover_title: str | None = None,
                 doc_name: str | None = None,
                 cover_org: str = COVER_ORG,
                 toc_fmt: str = "upperRoman"):
        self.path = path
        self.project = project
        self.applicant = applicant
        self.compiler = compiler
        self.date_cn = date_cn
        # 封面首行标题：优先用 cover_title；否则由“前缀 + 文件类型”拼，
        # 文件类型默认取文件名“项目名称_文件名称.docx”中下划线之后的部分
        self.cover_org = (cover_org or "").strip()
        self.doc_name = (doc_name or doc_name_of(path)).strip()
        self.cover_title = (cover_title or compose_cover_title(self.cover_org, self.doc_name)).strip()
        if not cover_title and not self.doc_name:
            raise ValueError(
                "封面标题缺少文件类型：请把文件名写成“项目名称_文件名称.docx”"
                "（如“某某项目_需求分析报告.docx”），或传 doc_name=“需求分析报告”，"
                "或直接传 cover_title 指定整串标题。"
            )
        self.toc_fmt = toc_fmt

        self.doc = Document()
        self._setup_page(self.doc.sections[0])
        self._setup_styles()
        self._chapter = 0
        self._tbl_seq: dict[tuple[int, int], int] = {}   # (章, 子项) -> 序号
        self._fig_seq: dict[int, int] = {}
        self._in_body = False

    # ------------------------------------------------------------ 基础设置

    @staticmethod
    def _setup_page(section) -> None:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = MARGIN_TOP
        section.bottom_margin = MARGIN_BOTTOM
        section.left_margin = MARGIN_LEFT
        section.right_margin = MARGIN_RIGHT
        section.header_distance = Mm(15)
        section.footer_distance = Mm(20)

    def _setup_styles(self) -> None:
        st = self.doc.styles

        normal = st["Normal"]
        normal.font.size = Pt(SIZE_3)
        normal.font.name = FONT_FANGSONG
        rpr = normal.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rfonts.set(qn(attr), FONT_FANGSONG)

        specs = [
            ("Heading 1", FONT_HEITI, 0),
            ("Heading 2", FONT_KAITI, 1),
            ("Heading 3", FONT_FANGSONG, 2),
            ("Heading 4", FONT_FANGSONG, 3),
        ]
        for name, font, lvl in specs:
            style = st[name]
            style.font.size = Pt(SIZE_3)
            style.font.bold = False
            style.font.italic = False
            style.font.color.rgb = None
            style.font.name = font
            srpr = style.element.get_or_add_rPr()
            srfonts = srpr.find(qn("w:rFonts"))
            if srfonts is None:
                srfonts = OxmlElement("w:rFonts")
                srpr.insert(0, srfonts)
            for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                srfonts.set(qn(attr), font)
            pf = style.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.line_spacing = Pt(BODY_LINE_PT)
            pf.keep_with_next = True
            ppr = style.element.get_or_add_pPr()
            el = ppr.find(qn("w:outlineLvl"))
            if el is None:
                el = OxmlElement("w:outlineLvl")
                ppr.append(el)
            el.set(qn("w:val"), str(lvl))

        self._setup_toc_styles()

    def _setup_toc_styles(self) -> None:
        """目录 1~3 级：宋体五号、行距固定18磅、缩进2/3.5/5字符、前两级加粗。"""
        styles_el = self.doc.styles.element
        specs = [("TOC1", "toc 1", 2.0, True),
                 ("TOC2", "toc 2", 3.5, True),
                 ("TOC3", "toc 3", 5.0, False)]
        for sid, name, indent, bold in specs:
            if styles_el.find(f'.//{qn("w:style")}[@{qn("w:styleId")}="{sid}"]') is not None:
                continue
            style = OxmlElement("w:style")
            style.set(qn("w:type"), "paragraph")
            style.set(qn("w:styleId"), sid)
            nm = OxmlElement("w:name")
            nm.set(qn("w:val"), name)
            style.append(nm)
            based = OxmlElement("w:basedOn")
            based.set(qn("w:val"), "Normal")
            style.append(based)
            nxt = OxmlElement("w:next")
            nxt.set(qn("w:val"), "Normal")
            style.append(nxt)
            style.append(OxmlElement("w:qFormat"))

            ppr = OxmlElement("w:pPr")
            spacing = OxmlElement("w:spacing")
            spacing.set(qn("w:before"), "0")
            spacing.set(qn("w:after"), "0")
            spacing.set(qn("w:line"), str(int(TABLE_LINE_PT * 20)))
            spacing.set(qn("w:lineRule"), "exact")
            ppr.append(spacing)
            ind = OxmlElement("w:ind")
            ind.set(qn("w:leftChars"), str(int(indent * 100)))
            ind.set(qn("w:left"), str(int(indent * SIZE_5 * 20)))
            ind.set(qn("w:firstLineChars"), "0")
            ind.set(qn("w:firstLine"), "0")
            ppr.append(ind)
            style.append(ppr)

            rpr = OxmlElement("w:rPr")
            rfonts = OxmlElement("w:rFonts")
            for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                rfonts.set(qn(attr), FONT_SONG)
            rpr.append(rfonts)
            if bold:
                rpr.append(OxmlElement("w:b"))
            sz = OxmlElement("w:sz")
            sz.set(qn("w:val"), str(int(SIZE_5 * 2)))
            rpr.append(sz)
            szcs = OxmlElement("w:szCs")
            szcs.set(qn("w:val"), str(int(SIZE_5 * 2)))
            rpr.append(szcs)
            style.append(rpr)
            styles_el.append(style)

    # ------------------------------------------------------------ 封面

    def cover(self) -> "DocFormat":
        sec = self.doc.sections[0]
        sec.footer.is_linked_to_previous = False
        _set_pgnum(sec, None, 1)
        # 封面无页码：页脚留空
        sec.footer.paragraphs[0].text = ""

        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(SIZE_4S * 1.5)  # 与小四1.5倍行距等高
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        r = p.add_run(self.cover_title)
        _set_run_fonts(r, FONT_XIAOBIAOSONG)
        r.font.size = Pt(SIZE_1)

        # 与项目名称间隔“三号字4行1.5倍行距”
        for _ in range(4):
            blank = self.doc.add_paragraph()
            blank.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            blank.paragraph_format.line_spacing = 1.5
            _no_space(blank)
            br = blank.add_run("")
            _set_run_fonts(br, FONT_FANGSONG)
            br.font.size = Pt(SIZE_3)

        for text in (self.project, self.applicant, self.compiler):
            if not text:
                continue
            par = self.doc.add_paragraph()
            par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            par.paragraph_format.line_spacing = 1.5
            _no_space(par)
            _set_char_indent(par, first_line_chars=2, font_pt=SIZE_3)
            rr = par.add_run(text)
            _set_run_fonts(rr, FONT_FANGSONG)
            rr.font.size = Pt(SIZE_3)

        if self.date_cn:
            date_par = self.doc.add_paragraph()
            date_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            date_par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            date_par.paragraph_format.line_spacing = 1.5
            _no_space(date_par)
            _set_char_indent(date_par, first_line_chars=1.49, font_pt=SIZE_3)
            dr = date_par.add_run(self.date_cn)
            _set_run_fonts(dr, FONT_FANGSONG)
            dr.font.size = Pt(SIZE_3)
        return self

    # ------------------------------------------------------------ 目录

    def toc(self, title: str = "目　　录") -> "DocFormat":
        sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
        self._setup_page(sec)
        sec.footer.is_linked_to_previous = False
        _set_pgnum(sec, self.toc_fmt, 1)
        self._footer_page_number(sec)

        tp = self.doc.add_paragraph()
        tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        tp.paragraph_format.line_spacing = Pt(BODY_LINE_PT)
        _no_space(tp)
        tr = tp.add_run(title)
        _set_run_fonts(tr, FONT_SONG)
        tr.font.size = Pt(SIZE_3)

        holder = self.doc.add_paragraph()
        _no_space(holder)
        _field(holder, r'TOC \o "1-3" \h \z \u', "右键此处选择“更新域”生成目录")

        # 打开文档时提示更新域，保证目录页码正确
        settings = self.doc.settings.element
        if settings.find(qn("w:updateFields")) is None:
            uf = OxmlElement("w:updateFields")
            uf.set(qn("w:val"), "true")
            settings.append(uf)
        return self

    def start_body(self) -> "DocFormat":
        sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
        self._setup_page(sec)
        sec.footer.is_linked_to_previous = False
        _set_pgnum(sec, "decimal", 1)
        self._footer_page_number(sec)
        self._in_body = True
        return self

    @staticmethod
    def _footer_page_number(section) -> None:
        footer = section.footer
        par = footer.paragraphs[0]
        par.text = ""
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _no_space(par)
        _fixed_line(par, TABLE_LINE_PT)
        _field(par, "PAGE", "1")
        for run in par.runs:
            _set_run_fonts(run, FONT_SONG)
            run.font.size = Pt(SIZE_5)

    # ------------------------------------------------------------ 标题与正文

    def h1(self, text: str, page_break: bool = False) -> "DocFormat":
        # 章号优先取标题中的“第X章”，便于从中间章节起排时编号仍对齐
        m = re.match(r"^\s*第([一二三四五六七八九十〇零]+)章", text)
        parsed = cn_to_int(m.group(1)) if m else None
        self._chapter = parsed if parsed else self._chapter + 1
        par = self.doc.add_paragraph(style="Heading 1")
        if page_break:
            par.paragraph_format.page_break_before = True
        _fixed_line(par, BODY_LINE_PT)
        _set_char_indent(par, first_line_chars=0, font_pt=SIZE_3)
        r = par.add_run(text)
        _set_run_fonts(r, FONT_HEITI)
        r.font.size = Pt(SIZE_3)
        _set_outline_level(par, 0)
        return self

    def h2(self, text: str) -> "DocFormat":
        par = self.doc.add_paragraph(style="Heading 2")
        _fixed_line(par, BODY_LINE_PT)
        _set_char_indent(par, first_line_chars=0, font_pt=SIZE_3)
        r = par.add_run(text)
        _set_run_fonts(r, FONT_KAITI)
        r.font.size = Pt(SIZE_3)
        _set_outline_level(par, 1)
        return self

    def h3(self, text: str) -> "DocFormat":
        return self._heading_fangsong(text, 2)

    def h4(self, text: str) -> "DocFormat":
        return self._heading_fangsong(text, 3)

    def _heading_fangsong(self, text: str, level: int) -> "DocFormat":
        par = self.doc.add_paragraph(style=f"Heading {level + 1}")
        _fixed_line(par, BODY_LINE_PT)
        _set_char_indent(par, first_line_chars=0, font_pt=SIZE_3)
        r = par.add_run(text)
        _set_run_fonts(r, FONT_FANGSONG)
        r.font.size = Pt(SIZE_3)
        _set_outline_level(par, level)
        return self

    def p(self, text: str, bold: bool = False, align=None,
          indent_chars: float = 2.0) -> "DocFormat":
        par = self.doc.add_paragraph()
        if align is not None:
            par.alignment = align
        else:
            par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _no_space(par)
        _fixed_line(par, BODY_LINE_PT)
        _set_char_indent(par, first_line_chars=indent_chars, font_pt=SIZE_3)
        _set_outline_level(par, 9)
        r = par.add_run(text)
        _set_run_fonts(r, FONT_FANGSONG)
        r.font.size = Pt(SIZE_3)
        r.font.bold = bold
        return self

    # ------------------------------------------------------------ 表格

    def table(self, caption: str, header: list[str], rows: list[list[str]],
              total: bool = False, widths_cm: list[float] | None = None,
              font_pt: float = SIZE_5, sub: int | None = None) -> "DocFormat":
        """插入带表题的表格。

        caption 只写表题文字，编号自动生成：
          章内顺排 -> 表X-1、表X-2……
          有子项时传入 sub（子项序号）-> 表X-1-1、表X-2-1……
          估算书章节按规范不做子项分层，直接顺排。
        """
        key = (self._chapter, sub or 0)
        seq = self._tbl_seq.get(key, 0) + 1
        self._tbl_seq[key] = seq
        if sub:
            label = f"表{self._chapter}-{sub}-{seq} {caption}"
        else:
            label = f"表{self._chapter}-{seq} {caption}"

        cap = self.doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.keep_with_next = True
        _no_space(cap)
        cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        cap.paragraph_format.line_spacing = 1.5
        cr = cap.add_run(label)
        _set_run_fonts(cr, FONT_HEI)
        cr.font.size = Pt(SIZE_5)

        table = self.doc.add_table(rows=1, cols=len(header))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True
        _table_borders(table)

        hdr = table.rows[0]
        _mark_header_row(hdr)
        for i, text in enumerate(header):
            cell = hdr.cells[i]
            par = cell.paragraphs[0]
            for old in list(par.runs):
                old._element.getparent().remove(old._element)
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _no_space(par)
            _fixed_line(par, TABLE_LINE_PT)
            _set_char_indent(par, first_line_chars=0, font_pt=font_pt)
            r = par.add_run(str(text))
            _set_run_fonts(r, FONT_SONG)
            r.font.size = Pt(font_pt)
            r.font.bold = True
            _cell_valign_center(cell)

        for idx, row_data in enumerate(rows):
            row = table.add_row()
            is_total = total and idx == len(rows) - 1
            for i, text in enumerate(row_data):
                cell = row.cells[i]
                par = cell.paragraphs[0]
                for old in list(par.runs):
                    old._element.getparent().remove(old._element)
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER if (is_total or i == 0) else WD_ALIGN_PARAGRAPH.LEFT
                _no_space(par)
                _fixed_line(par, TABLE_LINE_PT)
                _set_char_indent(par, first_line_chars=0, font_pt=font_pt)
                r = par.add_run(str(text))
                _set_run_fonts(r, FONT_SONG)
                r.font.size = Pt(font_pt)
                r.font.bold = is_total
                _cell_valign_center(cell)

        if widths_cm:
            for i, w in enumerate(widths_cm):
                for row in table.rows:
                    row.cells[i].width = Cm(w)
        return self

    # ------------------------------------------------------------ 图

    def figure(self, path: str, caption: str, width_cm: float = 15.0) -> "DocFormat":
        seq = self._fig_seq.get(self._chapter, 0) + 1
        self._fig_seq[self._chapter] = seq

        par = self.doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _no_space(par)
        par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        par.paragraph_format.line_spacing = 1.5
        par.add_run().add_picture(path, width=Cm(width_cm))

        cap = self.doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _no_space(cap)
        cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        cap.paragraph_format.line_spacing = 1.5
        cr = cap.add_run(f"图{self._chapter}-{seq} {caption}")
        _set_run_fonts(cr, FONT_HEI)
        cr.font.size = Pt(SIZE_5)
        return self

    # ------------------------------------------------------------ 保存

    def save(self) -> str:
        os.makedirs(os.path.dirname(os.path.abspath(self.path)) or ".", exist_ok=True)
        self.doc.save(self.path)
        return self.path
