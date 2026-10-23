#!/usr/bin/env python3
"""Build the Chinese competition introduction PDF from verified project output."""

from __future__ import annotations

from pathlib import Path

from PIL import Image as PillowImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[1]
WORKSPACE = ROOT.parents[1]
OUTPUT = WORKSPACE / "outputs" / "多源资料一致性核验助手-作品介绍.pdf"
SCREENSHOT = ROOT / "docs" / "assets" / "report-screenshot.png"
ARCHITECTURE = ROOT / "docs" / "assets" / "architecture.png"
FONT = (
    ROOT
    / "multi-source-consistency-auditor"
    / "scripts"
    / "multi_source_consistency_auditor"
    / "assets"
    / "NotoSansSC-VF.ttf"
)

pdfmetrics.registerFont(TTFont("MSCA", str(FONT)))
INK = colors.HexColor("#17212B")
MUTED = colors.HexColor("#607080")
TEAL = colors.HexColor("#156F73")
DARK = colors.HexColor("#12384B")
LINE = colors.HexColor("#D6E0E7")
PALE = colors.HexColor("#F3F7F8")

styles = getSampleStyleSheet()
TITLE = ParagraphStyle(
    "TitleCN",
    parent=styles["Title"],
    fontName="MSCA",
    fontSize=27,
    leading=38,
    textColor=DARK,
    alignment=TA_CENTER,
)
SUBTITLE = ParagraphStyle(
    "SubtitleCN",
    parent=styles["Normal"],
    fontName="MSCA",
    fontSize=13,
    leading=21,
    textColor=TEAL,
    alignment=TA_CENTER,
)
H1 = ParagraphStyle(
    "H1CN",
    parent=styles["Heading1"],
    fontName="MSCA",
    fontSize=17,
    leading=24,
    textColor=DARK,
    spaceBefore=3,
    spaceAfter=7,
)
H2 = ParagraphStyle(
    "H2CN",
    parent=styles["Heading2"],
    fontName="MSCA",
    fontSize=12,
    leading=18,
    textColor=TEAL,
    spaceBefore=8,
    spaceAfter=4,
)
BODY = ParagraphStyle(
    "BodyCN",
    parent=styles["BodyText"],
    fontName="MSCA",
    fontSize=9.1,
    leading=16,
    textColor=INK,
    spaceAfter=5,
    wordWrap="CJK",
)
SMALL = ParagraphStyle(
    "SmallCN",
    parent=BODY,
    fontSize=7.8,
    leading=12,
    textColor=MUTED,
)
CELL = ParagraphStyle(
    "CellCN",
    parent=BODY,
    fontSize=8.1,
    leading=12,
    spaceAfter=0,
)
CELL_HEAD = ParagraphStyle(
    "CellHeadCN",
    parent=CELL,
    textColor=colors.white,
)


def p(text: str, style: ParagraphStyle = BODY) -> Paragraph:
    return Paragraph(escape(text), style)


def table(data: list[list[object]], widths: list[float]) -> Table:
    wrapped = []
    for row_index, row in enumerate(data):
        style = CELL_HEAD if row_index == 0 else CELL
        wrapped.append(
            [
                cell if not isinstance(cell, str) else p(cell, style)
                for cell in row
            ]
        )
    item = Table(wrapped, colWidths=widths, repeatRows=1)
    item.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), DARK),
                ("TEXTCOLOR", (0, 1), (-1, -1), INK),
                ("GRID", (0, 0), (-1, -1), 0.4, LINE),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, PALE]),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return item


def cover(canvas, doc) -> None:
    canvas.saveState()
    canvas.setFillColor(colors.HexColor("#12384B"))
    canvas.rect(0, A4[1] - 48 * mm, A4[0], 48 * mm, fill=1, stroke=0)
    canvas.setFillColor(TEAL)
    canvas.rect(0, A4[1] - 51 * mm, A4[0], 3 * mm, fill=1, stroke=0)
    canvas.restoreState()


def page_decoration(canvas, doc) -> None:
    canvas.saveState()
    canvas.setStrokeColor(LINE)
    canvas.line(18 * mm, 14 * mm, A4[0] - 18 * mm, 14 * mm)
    canvas.setFont("MSCA", 7.2)
    canvas.setFillColor(MUTED)
    canvas.drawString(18 * mm, 9 * mm, "多源资料一致性核验助手 · 作品介绍")
    canvas.drawRightString(A4[0] - 18 * mm, 9 * mm, f"第 {doc.page} 页")
    canvas.restoreState()


def build() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    screenshot_top = ROOT / "docs" / "assets" / "report-screenshot-top.png"
    with PillowImage.open(SCREENSHOT) as image:
        width, height = image.size
        image.crop((0, 0, width, min(1850, height))).save(screenshot_top)

    document = SimpleDocTemplate(
        str(OUTPUT),
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=20 * mm,
        title="多源资料一致性核验助手 - 作品介绍",
        author="项目参赛团队",
        subject="第四届开放原子大赛·2026 开源行业解决方案创新赛 赛道一",
    )
    story: list[object] = [
        Spacer(1, 58 * mm),
        p("多源资料一致性核验助手", TITLE),
        Spacer(1, 4 * mm),
        p("一个只读、可复现、带原文证据的多文件交叉核验工具", SUBTITLE),
        Spacer(1, 12 * mm),
        p("第四届开放原子大赛 · 2026 开源行业解决方案创新赛", SUBTITLE),
        p("赛道一：开源基础软件与解决方案", SUBTITLE),
        Spacer(1, 30 * mm),
        table(
            [
                ["项目形态", "Agent Skill + Python 确定性核验引擎"],
                ["输入", "多个表格与文档来源，按业务编号关联"],
                ["输出", "HTML / Excel / Word / PDF / JSON / Markdown / CSV 差异汇总"],
                ["安全边界", "源文件只读、临时副本解析、全程离线、不自动发布"],
            ],
            [35 * mm, 135 * mm],
        ),
        Spacer(1, 25 * mm),
        p("2026 年 10 月", ParagraphStyle("Date", parent=SMALL, alignment=TA_CENTER)),
        PageBreak(),
        p("一、项目概述", H1),
        p(
            "多源资料一致性核验助手面向“同一业务实体分散在多份文件、多个表格和不同资料格式中”的核验问题。"
            "用户通过 JSON 配置定义业务编号、字段别名和核验规则，工具将多个来源转换为统一事实记录，"
            "按编号归组后检查来源缺失、重复编号、字段冲突、必要字段缺失和日期顺序。"
        ),
        p(
            "项目以 Agent Skill 作为智能体入口，以 Python 确定性程序作为核心执行引擎。"
            "工作流支持手动配置，也支持 Skill 根据任务组织输入、调用命令并解释结果。"
            "报告生成不依赖在线大模型，相同输入和规则可得到可复现结果。"
        ),
        p("项目核心价值", H2),
        table(
            [
                ["价值点", "实际效果"],
                ["多来源", "不再限制两份表格，可加入任意数量的表格、文档和 Excel Sheet"],
                ["可追溯", "每条差异保留来源文件、Sheet、行、页码或段落及原文摘录"],
                ["可复核", "确定性规则与标准化结果避免随机解释，人工可对照原证据判断"],
                ["可选择", "用户按需要生成 HTML、Excel、Word、PDF、JSON、Markdown 或 CSV"],
                ["守边界", "原文件只读，解析器仅处理自动清理的临时副本，不联网上传"],
            ],
            [34 * mm, 136 * mm],
        ),
        p("二、项目背景", H1),
        p(
            "采购、项目执行、行政归档和资料审查中，同一编号往往同时出现在表格、合同、申请单、验收材料和附件里。"
            "人工核对需要逐个打开文件、查找字段、抄录位置，再比较不同来源。资料数量增加后，"
            "漏项、金额不一致、孤立文件和重复编号容易被忽略，发现问题后也难以快速定位原文。"
        ),
        p(
            "普通对话式 AI 可以概括文本，但缺少稳定的解析流程、明确规则和可验证证据，无法保证同一资料每次得到同一结论。"
            "本项目的技术重点是把“找出不同”拆解成可测试的解析、标准化、关联和规则计算，"
            "让 AI 负责流程编排与结果解释，让程序负责事实计算。"
        ),
        PageBreak(),
        p("三、技术架构", H1),
        Image(str(ARCHITECTURE), width=170 * mm, height=72 * mm),
        Spacer(1, 5 * mm),
        table(
            [
                ["模块", "实现内容"],
                ["来源适配", "CSV、TSV、XLSX、XLSM、JSON、DOCX、文本型 PDF、TXT、Markdown、HTML"],
                ["只读隔离", "SHA-256 前后校验；第三方解析器只读取临时字节副本"],
                ["统一事实模型", "字段原值、标准化值、来源、文件位置和原文证据"],
                ["实体关联", "根据业务编号分组，支持一个实体对应多条记录和多份附件"],
                ["规则引擎", "来源覆盖、重复编号、字段一致性、数值容差、必填、日期顺序"],
                ["报告层", "七种可选格式，每种均包含汇总、实体状态、问题与证据"],
            ],
            [38 * mm, 132 * mm],
        ),
        p("工作流程", H2),
        p(
            "读取配置 → 定位多个来源 → 建立只读快照 → 解析表格和文档 → 字段标准化 → 按编号关联 → "
            "执行规则 → 输出差异汇总 → 人工结合原始资料复核。"
        ),
        p("配置化简例", H2),
        p(
            "配置文件声明“采购台账”“采购申请”“合同文件”三个来源，并设置供应商、金额、申请日期和签订日期字段。"
            "金额按 0.01 容差比较，日期检查申请是否晚于签署。来源名称与字段别名均可按业务场景替换。",
            SMALL,
        ),
        PageBreak(),
        p("四、核心功能与应用场景", H1),
        table(
            [
                ["功能", "输入与输出"],
                ["多文件读取", "一次解析多份表格和多份文档；一个 Excel 工作簿的每个 Sheet 独立读取"],
                ["字段核验", "比较名称、编号、金额、数量和日期；金额可设置容差"],
                ["资料完整性", "发现某来源缺少实体、文件无编号、同一来源编号重复"],
                ["日期逻辑", "检查申请晚于签署等时序矛盾"],
                ["证据追溯", "显示文件名、位置、原始值和标准化值"],
                ["格式选择", "HTML、Excel、Word、PDF、JSON、Markdown、CSV"],
            ],
            [37 * mm, 133 * mm],
        ),
        p("可应用场景", H2),
        p(
            "采购资料交叉审查；合同与申请材料核验；项目附件完整性检查；活动报名、审批、签到和汇总资料核对；"
            "行政档案目录核对；内部审计前的资料一致性预检。用户通过更换字段和来源配置适配不同流程。"
        ),
        p("五、创新点", H1),
        table(
            [
                ["对比对象", "本项目的实际特点"],
                ["人工逐份检查", "自动建立实体关联和来源覆盖检查，同时保留原始证据供人工复核"],
                ["普通文件汇总", "不只是拼接内容，而是标准化后执行明确规则并区分错误和警告"],
                ["简单提示词封装", "包含真实解析代码、配置协议、规则引擎、报告层和自动化测试"],
                ["传统对账工具", "同时处理表格与文档，支持多来源、多附件和七种输出格式"],
                ["数据处理安全", "解析器不直接处理原文件；临时副本自动清理；运行前后哈希核验"],
            ],
            [39 * mm, 131 * mm],
        ),
        PageBreak(),
        p("六、实际运行效果", H1),
        p(
            "样例使用虚构采购资料：1 个 Excel、1 个 CSV、2 个 Word 和 1 个文本型 PDF。"
            "系统关联出 5 个采购项目，识别 P-2026-002 的 3,200 元与 3,500 元金额冲突，"
            "并识别 P-2026-003、P-2026-004、P-2026-005 的来源缺失。"
        ),
        Image(str(screenshot_top), width=165 * mm, height=212 * mm),
        PageBreak(),
        p("七、测试与运行结果", H1),
        table(
            [
                ["测试项", "验证结果"],
                ["自动化测试", "11 项测试全部通过"],
                ["正常输入", "Excel、CSV、Word、PDF 均成功解析并提取字段"],
                ["多 Sheet", "同一工作簿两个 Sheet 均成功识别"],
                ["非法配置", "返回明确配置错误，不执行半成品核验"],
                ["重复编号", "产生 duplicate_identifier 错误"],
                ["字段冲突", "P-2026-002 金额冲突被定位，证据覆盖三个来源"],
                ["来源缺失", "3 个实体共 6 条来源缺失被标记为警告"],
                ["运行安全", "运行前后 5 个源文件 SHA-256 完全一致"],
                ["格式输出", "HTML、Excel、Word、PDF、JSON、Markdown、CSV 均生成有效文件"],
            ],
            [43 * mm, 127 * mm],
        ),
        p("已核验的输出特征", H2),
        p(
            "PDF 嵌入 Noto Sans SC 字体，样例为 2 页，包含总体汇总、实体对照、差异明细、证据索引和来源清单；"
            "文本可提取，字符位于页面边界内。Excel 包含五个工作表，Word 可直接编辑，JSON 保留完整结构化证据。"
        ),
        p("八、使用限制", H1),
        p(
            "核心版本支持可解析的表格和文本型文件；扫描版 PDF 和图片需要 OCR，当前不会在文本为空时猜测内容。"
            "文档字段采用“字段名：值”形式，复杂版式可能需要补充别名或预处理配置。"
            "规则必须由用户或责任岗位定义，工具不判断材料在法律、财务或业务上的真实性。"
        ),
        p("九、可复现运行", H1),
        p(
            "安装依赖后运行示例生成脚本，再执行核验命令并选择输出格式。README 提供 Windows 和 Linux/macOS 的安装、"
            "运行、配置与测试步骤；项目使用 MIT 许可证，PDF 字体使用 SIL Open Font License 1.1。"
        ),
        p(
            "本研究作品的全部测试结果来自当前源码和虚构样例，不包含真实个人数据、企业内部资料或未经授权的上传链接。",
            SMALL,
        ),
    ]
    document.build(story, onFirstPage=cover, onLaterPages=page_decoration)
    print(OUTPUT)


if __name__ == "__main__":
    build()
