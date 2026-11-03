#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成一份符合本规范的示例文档，用于验证排版库与自检脚本。

用法：
    python demo.py            # 输出到 ./示例项目_可行性研究报告.docx
    python demo.py 输出路径
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import docformat as G
from PIL import Image, ImageDraw


def make_chart(path: str) -> str:
    """画一张简单的架构示意图（PIL 生成，避免示例依赖外部素材）。"""
    w, h = 1600, 900
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    boxes = [(120, 320, "数据接入"), (620, 320, "测评执行"), (1120, 320, "报告归档")]
    for x, y, label in boxes:
        d.rounded_rectangle([x, y, x + 340, y + 220], radius=18, outline="#1f4e79", width=4)
        d.text((x + 120, y + 100), label, fill="#1f4e79")
        if x < 1100:
            d.line([x + 340, y + 110, x + 500, y + 110], fill="#1f4e79", width=4)
    d.text((620, 180), "总体架构", fill="#333333")
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    img.save(path)
    return path


def main(out: str | None = None) -> str:
    out = out or os.path.join(os.getcwd(), G.doc_filename("示例项目", "可行性研究报告"))
    fig = make_chart(os.path.join(os.path.dirname(os.path.abspath(out)), "架构示意.png"))

    doc = G.DocFormat(
        out,
        project="××信息化平台建设",
        applicant="××单位",
        compiler="××单位",
        cover_org="××项目建设",
        date_cn=G.cn_date(2026, 11),
    )
    doc.cover()
    doc.toc()
    doc.start_body()

    doc.h1("第一章 项目概述")
    doc.h2("1.1 项目背景")
    doc.p("随着人工智能技术在调度、运检、服务等业务域的规模化应用，"
          "模型上线前的安全性与可靠性测评成为必要条件。目前测评环境依赖人工搭建，"
          "存在环境不一致、用例复用率低、过程数据难以追溯等问题，亟需建设统一的仿真测评环境。")
    doc.h2("1.2 建设目标")
    doc.p("构建覆盖环境编排、用例执行、过程监控与报告输出的测评环境仿真平台，"
          "实现测评环境一键构建、测评任务并行调度、测评过程全程留痕，如图1-1所示。")
    doc.figure(fig, "平台总体架构", width_cm=14)
    doc.h3("1.2.1 建设内容")
    doc.table("建设内容与规模一览表",
              ["序号", "建设内容", "规模", "说明"],
              [["1", "仿真环境编排", "12 类环境模板", "覆盖语义、视觉、多模态模型测评场景"],
               ["2", "测评工具与算子", "48 个算子", "支持对抗样本生成与鲁棒性评估"],
               ["3", "过程监控与留痕", "全流程", "任务、数据、操作三类记录可追溯"],
               ["合计", "—", "—", "三项建设内容"]],
              total=True, widths_cm=[1.6, 4.0, 3.0, 7.4])

    doc.h1("第二章 需求分析")
    doc.h2("2.1 功能需求")
    doc.p("平台功能需求按业务角色划分，管理员负责环境与账号管理，测评人员负责任务执行，"
          "专家负责结论复核。各角色功能需求见表2-1。")
    doc.table("功能需求清单",
              ["序号", "功能项", "使用角色", "需求描述"],
              [["1", "环境模板管理", "管理员", "环境模板的创建、版本发布与下线"],
               ["2", "测评任务调度", "测评人员", "任务编排、并行执行与状态跟踪"],
               ["3", "结论复核", "专家", "测评报告复核与结论审定"]],
              widths_cm=[1.6, 4.0, 3.0, 7.4])
    doc.h2("2.2 非功能需求")
    doc.p("平台应满足测评隔离区部署要求，单任务并发不低于 20 个，"
          "关键操作响应时间不超过 3 秒，全部操作留痕并可按时间与角色检索。")

    doc.save()
    return out


if __name__ == "__main__":
    path = main(sys.argv[1] if len(sys.argv) > 1 else None)
    print(f"已生成：{path}")
