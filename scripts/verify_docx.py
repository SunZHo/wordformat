#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""公文版式自检：核对页边距、页码域、字体字号、缩进、行距、表图题。

用法：
    python verify_docx.py 文件.docx
仅在“全部通过”时打印 OK，其他情况列出可定位的问题行，便于修正后重跑。
"""
from __future__ import annotations

import sys
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Mm

EXPECT = {
    "top": (34.0, 36.0),
    "bottom": (31.0, 33.0),
    "left": (27.0, 29.0),
    "right": (25.0, 27.0),
}

problems: list[str] = []


def check(cond: bool, msg: str) -> None:
    if not cond:
        problems.append(msg)


def mm(value) -> float:
    return round(value.mm, 1) if value is not None else -1


def run_font(run):
    rpr = run._element.find(qn("w:rPr"))
    if rpr is None:
        return None
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        return None
    return rfonts.get(qn("w:eastAsia"))


def first_run(par):
    """取首个有文字的 run（避免空 run 导致字号读不到）。"""
    for r in par.runs:
        if r.text.strip():
            return r
    return par.runs[0] if par.runs else None


def para_indent(par):
    ppr = par._p.find(qn("w:pPr"))
    if ppr is None:
        return {}
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        return {}
    return {k.split("}")[-1]: v for k, v in ind.attrib.items()}


def section_break_indexes(doc):
    idx = []
    for i, par in enumerate(doc.paragraphs):
        ppr = par._p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
            idx.append(i)
    return idx


def para_spacing(par):
    ppr = par._p.find(qn("w:pPr"))
    if ppr is None:
        return {}
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        return {}
    return {k.split("}")[-1]: v for k, v in sp.attrib.items()}


def main(path: str) -> int:
    doc = Document(path)
    body_text = "\n".join(p.text for p in doc.paragraphs)
    brk = section_break_indexes(doc)
    body_start = (brk[1] + 1) if len(brk) >= 2 else 0

    # 1 页面设置
    for i, sec in enumerate(doc.sections):
        for side, (lo, hi) in EXPECT.items():
            v = mm(getattr(sec, f"{side}_margin"))
            check(lo <= v <= hi, f"节{i + 1} 页边距{side}={v}mm，应在{lo}~{hi}mm")
        pg = sec._sectPr.find(qn("w:pgNumType"))
        if i == 0:
            check(not sec.footer.paragraphs[0].text.strip(),
                  "封面页脚应为空（封面不编页码）")
        else:
            check(pg is not None, f"节{i + 1} 缺少 pgNumType")
            if pg is not None:
                fmt = pg.get(qn("w:fmt"))
                start = pg.get(qn("w:start"))
                if i == 1:
                    check(fmt in ("upperRoman", "lowerRoman"), f"目录页码格式={fmt}，应为罗马数字")
                else:
                    check(fmt == "decimal", f"正文页码格式={fmt}，应为阿拉伯数字")
                check(start == "1", f"节{i + 1} 页码起始={start}，应为1")

    # 2 目录样式
    styles = {s.style_id: s for s in doc.styles}
    for sid, chars, bold in (("TOC1", 200, True), ("TOC2", 350, True), ("TOC3", 500, False)):
        st = styles.get(sid)
        check(st is not None, f"缺少目录样式 {sid}")
        if st is None:
            continue
        ppr = st.element.find(qn("w:pPr"))
        ind = ppr.find(qn("w:ind")) if ppr is not None else None
        sp = ppr.find(qn("w:spacing")) if ppr is not None else None
        check(ind is not None and ind.get(qn("w:leftChars")) == str(chars),
              f"{sid} 左缩进应为{chars / 100}字符")
        check(sp is not None and sp.get(qn("w:line")) == "360",
              f"{sid} 行距应为固定18磅(360twips)")
        check(sp is not None and sp.get(qn("w:before")) == "0" and sp.get(qn("w:after")) == "0",
              f"{sid} 段前段后应为0")

    # 2.5 目录标题：宋体三号居中
    if len(brk) >= 2:
        toc_pars = [p for p in doc.paragraphs[brk[0] + 1:brk[1]] if p.text.strip()]
        if toc_pars:
            tp = toc_pars[0]
            tr0 = first_run(tp)
            check(tr0 is not None and run_font(tr0) == "宋体", f"目录标题字体应为宋体，实为{run_font(tr0) if tr0 else None}")
            check(tr0 is not None and tr0.font.size is not None and abs(tr0.font.size.pt - 16) < 0.1,
                  "目录标题字号应为三号(16磅)")
            check(tp.alignment == WD_ALIGN_PARAGRAPH.CENTER, "目录标题应居中")

    # 3 标题层级与字体
    heading_font = {}
    for par in doc.paragraphs:
        style_id = par.style.style_id if par.style is not None else ""
        if style_id.startswith("Heading") and par.text.strip():
            lvl = int(style_id.replace("Heading", "") or 1)
            r0 = first_run(par)
            font = run_font(r0) if r0 is not None else None
            heading_font.setdefault(lvl, set()).add(font)
            hr = first_run(par)
            check(hr is not None and hr.font.size is not None and abs(hr.font.size.pt - 16) < 0.1,
                  f"标题“{par.text[:14]}”字号应为三号(16磅)")
    for lvl, fonts in heading_font.items():
        want = {1: "方正黑体_GBK", 2: "方正楷体_GBK"}.get(lvl, "方正仿宋_GBK")
        check(fonts == {want}, f"{lvl}级标题字体={fonts}，应为{want}")

    # 4 正文段落
    # 封面：标题小标宋一号、信息行 1.5 倍行距 + 首行缩进 2 字符
    cover_pars = doc.paragraphs[:brk[0]] if brk else doc.paragraphs
    if cover_pars:
        t = cover_pars[0]
        tr0 = first_run(t)
        check(tr0 is not None and run_font(tr0) == "方正小标宋简体",
              f"封面标题字体={run_font(tr0) if tr0 is not None else None}，应为方正小标宋简体")
        check(tr0 is not None and tr0.font.size is not None and abs(tr0.font.size.pt - 26) < 0.1,
              "封面标题字号应为一号(26磅)")
        check(t.alignment is not None and t.alignment == WD_ALIGN_PARAGRAPH.CENTER,
              "封面标题应居中对齐")
        for par in cover_pars[1:]:
            if not par.text.strip():
                continue
            r0 = first_run(par)
            check(r0 is not None and run_font(r0) == "方正仿宋_GBK",
                  f"封面“{par.text[:10]}”应为方正仿宋_GBK")
            check(para_spacing(par).get("lineRule") == "auto",
                  f"封面“{par.text[:10]}”应为1.5倍行距")

    body_checked = 0
    for i, par in enumerate(doc.paragraphs):
        if i < body_start:
            continue
        sid = par.style.style_id if par.style is not None else ""
        if sid.startswith("Heading") or sid.startswith("TOC") or not par.text.strip():
            continue
        if par.text.startswith(("表", "图")) and len(par.text) > 3:
            continue
        ind = para_indent(par)
        sp = para_spacing(par)
        if ind.get("firstLineChars") == "200":
            body_checked += 1
            check(sp.get("line") == "580", f"正文“{par.text[:12]}”行距{sp.get('line')}，应为固定29磅(580twips)")
            check(sp.get("before") in ("0", None) and sp.get("after") in ("0", None),
                  f"正文“{par.text[:12]}”段前段后应为0")
            r0 = first_run(par)
            font = run_font(r0) if r0 is not None else None
            check(font == "方正仿宋_GBK", f"正文“{par.text[:12]}”字体={font}，应为方正仿宋_GBK")
    check(body_checked > 0, "未发现符合首行缩进2字符的正文段落")

    # 5 表格
    for ti, table in enumerate(doc.tables, 1):
        tblpr = table._tbl.tblPr
        check(tblpr.find(qn("w:tblpPr")) is None, f"表{ti} 应为无文字环绕（不应有 tblpPr）")
        hdr = table.rows[0]._tr
        check(hdr.find(qn("w:trPr")) is not None and
              hdr.find(qn("w:trPr")).find(qn("w:tblHeader")) is not None,
              f"表{ti} 表头未设置跨页重复")
        cell = table.rows[0].cells[0]
        tcpr = cell._tc.find(qn("w:tcPr"))
        check(tcpr is not None and tcpr.find(qn("w:vAlign")) is not None and
              tcpr.find(qn("w:vAlign")).get(qn("w:val")) == "center",
              f"表{ti} 单元格垂直对齐应为居中")
        p0 = cell.paragraphs[0]
        sp = para_spacing(p0)
        check(sp.get("line") == "360", f"表{ti} 表格行距{sp.get('line')}，应为固定18磅")
        c0 = first_run(p0)
        check(c0 is not None and c0.font.size is not None and abs(c0.font.size.pt - 10.5) < 0.1, f"表{ti} 表格字号应为五号(10.5磅)")
        check(c0 is not None and run_font(c0) == "宋体", f"表{ti} 表格字体应为宋体")

    # 6 表题/图题编号
    labels = [p.text.strip() for p in doc.paragraphs
              if p.text.strip().startswith(("表", "图")) and len(p.text.strip()) > 3]
    for lb in labels:
        head = lb.split(" ")[0]
        check("-" in head, f"题注“{lb[:16]}”缺少“章号-序号”编号")
        check(lb.count(" ") >= 1, f"题注“{lb[:16]}”编号与题名之间应空一个字符")

    # 7 命名
    import os
    check("_" in os.path.basename(path), "文件名应为“项目名称_文件名称.docx”")

    print(f"文件：{path}")
    print(f"段落 {len(doc.paragraphs)}，表格 {len(doc.tables)}，节 {len(doc.sections)}，题注 {len(labels)}")
    if problems:
        print(f"发现 {len(problems)} 处不符合：")
        for p in problems:
            print("  -", p)
        return 1
    print("OK：页边距、页码、字体字号、缩进行距、表格属性、题注编号均符合规范")
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1]))
