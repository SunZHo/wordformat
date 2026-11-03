#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把一份既有 .docx 重排为本规范格式（结构重建，不做版面微调）。

用法：
    python reformat.py 源文件.docx 输出.docx --project 项目名称 \
        --applicant 申报单位 --compiler 编制单位 --date 二〇二六年十一月

处理范围：标题层级（按 Heading 样式或“第X章/1.1/1.1.1”编号推断）、正文段落、
表格（自动补表题编号）、表图题注复用原文字。图片与页眉页脚不迁移，遇到时给出提示。
"""
from __future__ import annotations

import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

import docformat as G

CAP_TBL = re.compile(r"^表\s*\d+(-\d+)*\s*[ 　]*(.+)$")
CAP_FIG = re.compile(r"^图\s*\d+(-\d+)*\s*[ 　]*(.+)$")
H_CHAPTER = re.compile(r"^第[一二三四五六七八九十]+章")


def iter_block_items(doc):
    from docx.oxml.ns import qn
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def heading_level(par: Paragraph) -> int:
    sid = par.style.style_id if par.style is not None else ""
    if sid.startswith("Heading"):
        try:
            return int(sid.replace("Heading", "") or 1)
        except ValueError:
            return 1
    text = par.text.strip()
    if H_CHAPTER.match(text):
        return 1
    if re.match(r"^\d+\.\d+\.\d+\.\d+\s", text):
        return 4
    if re.match(r"^\d+\.\d+\.\d+\s", text):
        return 3
    if re.match(r"^\d+\.\d+\s", text):
        return 2
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--project", default="")
    ap.add_argument("--applicant", default="")
    ap.add_argument("--compiler", default="")
    ap.add_argument("--date", default="")
    ap.add_argument("--cover-title", default="", help="封面首行标题，留空则按“前缀+文件类型”自动拼")
    ap.add_argument("--doc-name", default="", help="文件类型（如“需求分析报告”），留空则从输出文件名推断")
    ap.add_argument("--cover-org", default=G.COVER_ORG, help="封面标题前缀")
    ap.add_argument("--skip-cover", action="store_true")
    args = ap.parse_args()

    src = Document(args.src)
    doc = G.DocFormat(args.out, project=args.project, applicant=args.applicant,
                    compiler=args.compiler, date_cn=args.date,
                    cover_title=args.cover_title or None,
                    doc_name=args.doc_name or None,
                    cover_org=args.cover_org)
    if not args.skip_cover:
        doc.cover()
    doc.toc()
    doc.start_body()

    pending_caption: str | None = None
    dropped_images = 0

    for block in iter_block_items(src):
        if isinstance(block, Table):
            header = [c.text.strip() for c in block.rows[0].cells]
            rows = [[c.text.strip() for c in r.cells] for r in block.rows[1:]]
            caption = pending_caption or "数据表"
            pending_caption = None
            doc.table(caption, header, rows)
            continue

        text = block.text.strip()
        if not text:
            continue
        if block._p.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
            dropped_images += 1
            continue
        m_tbl = CAP_TBL.match(text)
        if m_tbl:
            pending_caption = m_tbl.group(2).strip()
            continue
        if CAP_FIG.match(text):
            continue  # 图题随图丢弃，重新插图时再补

        lvl = heading_level(block)
        if lvl == 1:
            doc.h1(text)
        elif lvl == 2:
            doc.h2(text)
        elif lvl == 3:
            doc.h3(text)
        elif lvl == 4:
            doc.h4(text)
        else:
            doc.p(text)

    doc.save()
    print(f"已生成：{args.out}")
    if dropped_images:
        print(f"提示：源文档含 {dropped_images} 处图片，未迁移，请用 doc.figure(path, caption) 重新插入。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
