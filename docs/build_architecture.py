"""Build the architecture diagram used by the project introduction."""

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

ROOT = Path(__file__).resolve().parents[1]
FONT = (
    ROOT
    / "multi-source-consistency-auditor"
    / "scripts"
    / "multi_source_consistency_auditor"
    / "assets"
    / "NotoSansSC-VF.ttf"
)
OUTPUT = ROOT / "docs" / "assets" / "architecture.png"


def box(ax, x, y, w, h, title, detail, color):
    patch = FancyBboxPatch(
        (x, y),
        w,
        h,
        boxstyle="round,pad=0.018,rounding_size=0.025",
        facecolor=color,
        edgecolor="#C8D5DD",
        linewidth=1.2,
    )
    ax.add_patch(patch)
    ax.text(
        x + w / 2,
        y + h * 0.64,
        title,
        ha="center",
        va="center",
        fontsize=15,
        color="#12384B",
        fontweight="bold",
        fontproperties=FONT_PROP,
    )
    ax.text(
        x + w / 2,
        y + h * 0.30,
        detail,
        ha="center",
        va="center",
        fontsize=9.5,
        color="#526575",
        linespacing=1.4,
        fontproperties=FONT_PROP,
    )


def arrow(ax, start, end):
    ax.add_patch(
        FancyArrowPatch(
            start,
            end,
            arrowstyle="-|>",
            mutation_scale=15,
            color="#156F73",
            linewidth=2,
            shrinkA=4,
            shrinkB=4,
        )
    )


font_manager.fontManager.addfont(str(FONT))
FONT_PROP = font_manager.FontProperties(fname=str(FONT))
plt.rcParams["axes.unicode_minus"] = False
fig, ax = plt.subplots(figsize=(13.5, 5.8), dpi=160)
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis("off")
fig.patch.set_facecolor("#F4F7F8")

ax.text(
    0.05,
    0.90,
    "多源资料一致性核验引擎",
    fontsize=23,
    color="#12384B",
    fontweight="bold",
    fontproperties=FONT_PROP,
)
ax.text(
    0.05,
    0.84,
    "确定性解析与规则计算，报告可追溯，原始资料不写入",
    fontsize=11,
    color="#607080",
    fontproperties=FONT_PROP,
)

box(ax, 0.04, 0.48, 0.17, 0.23, "多源输入", "Excel / CSV / JSON\nWord / PDF / Markdown", "#E4F1F1")
box(ax, 0.25, 0.48, 0.17, 0.23, "只读快照", "SHA-256 核验\n临时字节副本", "#EAF0F3")
box(ax, 0.46, 0.48, 0.17, 0.23, "统一事实", "编号 / 名称 / 金额\n日期与来源证据", "#EAF0F3")
box(ax, 0.67, 0.48, 0.17, 0.23, "实体关联", "按业务编号分组\n支持一实体多附件", "#EAF0F3")
box(ax, 0.80, 0.12, 0.16, 0.23, "可配置报告", "HTML / Excel / Word\nPDF / JSON / Markdown / CSV", "#E7F3EC")

arrow(ax, (0.21, 0.595), (0.25, 0.595))
arrow(ax, (0.42, 0.595), (0.46, 0.595))
arrow(ax, (0.63, 0.595), (0.67, 0.595))
arrow(ax, (0.76, 0.48), (0.82, 0.35))

box(ax, 0.46, 0.16, 0.24, 0.20, "规则核验层", "来源覆盖 · 字段冲突\n重复编号 · 必填 · 日期顺序", "#FFF3DD")
arrow(ax, (0.755, 0.48), (0.65, 0.36))
arrow(ax, (0.70, 0.26), (0.80, 0.26))

ax.text(
    0.05,
    0.20,
    "数据边界",
    fontsize=13,
    color="#12384B",
    fontweight="bold",
    fontproperties=FONT_PROP,
)
ax.text(
    0.05,
    0.14,
    "源文件只读  |  不转换格式  |  不改写内容  |  不联网上传\n"
    "输出目录与输入目录隔离  |  差异报告保留原文证据",
    fontsize=10,
    color="#526575",
    linespacing=1.7,
    fontproperties=FONT_PROP,
)

fig.tight_layout()
OUTPUT.parent.mkdir(parents=True, exist_ok=True)
fig.savefig(OUTPUT, bbox_inches="tight", facecolor=fig.get_facecolor())
plt.close(fig)
print(OUTPUT)
