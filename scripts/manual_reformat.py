#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把既有「图文操作手册」重排为本规范格式。

与 reformat.py 的分工：reformat.py 面向纯文本文档，图片不迁移；本脚本面向
带截图的软件操作手册，把内嵌截图导出后按规范重新插入，表格重建为规范样式，
封面、目录、页码、标题层级、正文行距全部按规范落地。

用法：
    python manual_reformat.py 源.docx 输出.docx --software 软件全称 \
        --applicant 申报单位 --compiler 编制单位 --date 二〇二六年十一月
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

import docformat as G

BLIP = ".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"

RE_H1 = re.compile(r"^第[一二三四五六七八九十〇零]+章")
RE_H2 = re.compile(r"^\d+\.\d+(?!\.)\s*")
RE_H3 = re.compile(r"^\d+\.\d+\.\d+\s*")
RE_FIGCAP = re.compile(r"^图\s*\d+\s*[-–—]\s*\d+\s*(.*)$")
RE_TOC = re.compile(r"^目\s*录$")
RE_COMPILER = re.compile(r"^编制单位[：:]\s*(.+)$")
RE_TOTAL = re.compile(r"合计|总计|小计")
CN_SUB = ["二", "三", "四", "五", "六", "七", "八", "九"]

FIG_WIDTH_CM = 15.6      # 版心宽 156mm，图不得超出版心


def iter_blocks(doc):
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def has_image(par: Paragraph) -> bool:
    return bool(par._p.findall(BLIP))


def export_images(par: Paragraph, outdir: str, idx: int) -> list[str]:
    """导出一个段落里的全部内嵌图片，返回文件路径。"""
    paths: list[str] = []
    for blip in par._p.findall(BLIP):
        rid = blip.get(qn("r:embed"))
        if not rid:
            continue
        part = par.part.related_parts[rid]
        ext = os.path.splitext(str(part.partname))[1] or ".png"
        path = os.path.join(outdir, f"img{idx:04d}_{len(paths)}{ext}")
        with open(path, "wb") as fh:
            fh.write(part.blob)
        paths.append(path)
    return paths


def parse_cover(doc) -> tuple[str, str]:
    """从原封面取出软件全称与编制单位。"""
    software, compiler = "", ""
    for par in doc.paragraphs:
        text = par.text.strip()
        if not text:
            continue
        if text.startswith("目"):
            break
        m = RE_COMPILER.match(text)
        if m:
            compiler = m.group(1).strip()
            continue
        if not software and "手册" not in text:
            software = text
    return software, compiler


def strip_number(title: str) -> str:
    """去掉标题前编号，用作表题。"""
    return re.sub(r"^(第[一二三四五六七八九十〇零]+章|\d+(\.\d+)*)\s*", "", title).strip()


def convert(src: str, out: str, software: str = "", applicant: str = "",
            compiler: str = "", date_cn: str = "", doc_name: str = "操作手册",
            fig_width: float = FIG_WIDTH_CM, verbose: bool = True) -> dict:
    origin = Document(src)
    auto_software, auto_compiler = parse_cover(origin)
    software = software or auto_software
    compiler = compiler or auto_compiler
    if not software:
        raise ValueError(f"未能从 {os.path.basename(src)} 判断软件名称，请用 --software 指定。")

    blocks = list(iter_blocks(origin))
    toc_idx = next((i for i, b in enumerate(blocks)
                    if isinstance(b, Paragraph) and RE_TOC.match(b.text.strip())), -1)
    # 目录里也逐条列了“第一章…”，正文起始取“第一章”的第二次出现
    chapter_idx = [i for i, b in enumerate(blocks)
                   if isinstance(b, Paragraph) and RE_H1.match(b.text.strip())
                   and b.text.strip().startswith("第一章")]
    body_start = chapter_idx[1] if len(chapter_idx) >= 2 else (chapter_idx[0] if chapter_idx else -1)
    if body_start < 0 or (toc_idx >= 0 and body_start < toc_idx):
        raise ValueError("未找到正文起始的“第一章”标题，无法确定重排范围。")

    tmpdir = tempfile.mkdtemp(prefix="manual_imgs_")
    doc = G.DocFormat(out, project=software, applicant=applicant, compiler=compiler,
                    date_cn=date_cn, cover_title=f"{software} {doc_name}")
    doc.cover()
    doc.toc()
    doc.start_body()

    sec_title = ""
    last_level = 0
    tbl_count: dict[str, int] = {}
    stats = {"h1": 0, "h2": 0, "h3": 0, "p": 0, "fig": 0, "tbl": 0, "skip": 0}

    i = body_start
    while i < len(blocks):
        blk = blocks[i]

        if isinstance(blk, Table):
            header = [c.text.strip() for c in blk.rows[0].cells]
            rows = [[c.text.strip() for c in r.cells] for r in blk.rows[1:]]
            # 表题按结构取：小节名优先；直接挂在章下的表用首列表头，避免“表4-1 功能操作说明”
            if last_level >= 2 and strip_number(sec_title):
                caption = strip_number(sec_title)
            elif header and header[0]:
                caption = f"{header[0]}一览表"
            else:
                caption = "数据一览表"
            count = tbl_count.get(caption, 0) + 1
            tbl_count[caption] = count
            if count > 1:
                caption = f"{caption}（{CN_SUB[min(count - 2, len(CN_SUB) - 1)]}）"
            total = bool(rows) and bool(RE_TOTAL.search(rows[-1][0] or ""))
            doc.table(caption, header, rows, total=total)
            stats["tbl"] += 1
            i += 1
            continue

        text = blk.text.strip()

        if has_image(blk):
            # 图题在图片下方的一段，形如“图 3-1 平台登录页”
            caption = ""
            if i + 1 < len(blocks) and isinstance(blocks[i + 1], Paragraph):
                m = RE_FIGCAP.match(blocks[i + 1].text.strip())
                if m:
                    caption = m.group(1).strip()
                    stats["skip"] += 1
                    i += 1
            for path in export_images(blk, tmpdir, i):
                doc.figure(path, caption or "界面截图", width_cm=fig_width)
                stats["fig"] += 1
            i += 1
            continue

        if not text:
            i += 1
            continue

        if RE_H1.match(text):
            doc.h1(text, page_break=stats["h1"] > 0)
            sec_title, last_level = text, 1
            stats["h1"] += 1
        elif RE_H3.match(text):
            doc.h3(text)
            sec_title, last_level = text, 3
            stats["h3"] += 1
        elif RE_H2.match(text):
            doc.h2(text)
            sec_title, last_level = text, 2
            stats["h2"] += 1
        elif RE_FIGCAP.match(text):
            stats["skip"] += 1        # 无对应图片的孤立图题，丢弃避免重号
        else:
            doc.p(text)
            stats["p"] += 1
        i += 1

    doc.save()
    if verbose:
        print(f"已生成：{out}")
        print(f"  章{stats['h1']} 节{stats['h2']} 三级{stats['h3']} 正文{stats['p']} "
              f"图{stats['fig']} 表{stats['tbl']} 丢弃图题{stats['skip']}")
    return stats


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--software", default="", help="软件全称（含版本），缺省从封面首行取")
    ap.add_argument("--applicant", default="", help="封面项目申报单位")
    ap.add_argument("--compiler", default="", help="封面编制单位，缺省从原封面取")
    ap.add_argument("--date", default="二〇二六年十一月", help="封面日期，汉字数字")
    ap.add_argument("--doc-name", default="操作手册", help="封面标题后缀")
    ap.add_argument("--fig-width", type=float, default=FIG_WIDTH_CM)
    args = ap.parse_args()

    convert(args.src, args.out, software=args.software, applicant=args.applicant,
            compiler=args.compiler, date_cn=args.date, doc_name=args.doc_name,
            fig_width=args.fig_width)
    return 0


if __name__ == "__main__":
    sys.exit(main())
