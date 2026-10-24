# 多源资料一致性核验助手

一个可本地运行、只读处理原始资料的多来源一致性核验工具。它将表格和文档转换成带原文证据的事实记录，通过同一实体编号关联，自动检查缺失来源、重复记录、字段冲突和日期顺序。

项目以 Agent Skill 为入口，以 Python 确定性核验引擎为核心。报告生成不依赖大模型或在线 API，因此相同输入和规则可以得到可复现结果。

## 核心功能

- 同时读取任意数量的 CSV、TSV、Excel、JSON 表格来源。
- 读取 Word、文本型 PDF、TXT、Markdown 和 HTML 文档来源。
- 支持一个 Excel 工作簿的多个 Sheet，以及一个实体对应多份附件。
- 按单一编号或配置中的业务编号关联多来源记录。
- 检查来源覆盖、重复编号、字段缺失、字段冲突、金额容差和日期顺序。
- 输出 HTML、Excel、Markdown、JSON 和 CSV 报告。
- 每条问题保留来源文件、Sheet、行号、页码或段落及原文摘录。
- 源文件只读；解析器仅处理自动清理的临时字节副本。
- 运行前后执行 SHA-256 完整性核验，检测外部改动。
- 全程离线，不修改、转换、上传或发布原始资料。

## 技术栈

- Python 3.10+
- openpyxl：Excel 读取与报告生成
- python-docx：Word 文档读取
- pypdf：文本型 PDF 读取
- ReportLab + Noto Sans SC：中文 PDF 报告生成与字体嵌入
- python-dateutil：日期标准化
- 标准库：CSV、JSON、HTML、规则和报告逻辑

## 安装

```bash
python -m venv .venv
```

Windows：

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Linux/macOS：

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
```

## 快速运行

生成模拟采购样例：

```bash
python examples/build_samples.py
```

执行核验：

```bash
python multi-source-consistency-auditor/scripts/audit.py \
  --profile examples/procurement/profile.json \
  --output examples/procurement/output \
  --format xlsx
```

打开 `examples/procurement/output/differences-summary.xlsx` 查看差异汇总。输出格式由用户选择，可选 `html`、`xlsx`、`docx`、`pdf`、`markdown`、`json`、`csv`；重复指定 `--format` 可一次生成多个格式，指定 `--format all` 可生成全部格式。命令行发现 error 级问题时可增加 `--fail-on-error` 返回非零退出码。生成 Word 或 PDF 示例：

```bash
python multi-source-consistency-auditor/scripts/audit.py \
  --profile examples/procurement/profile.json \
  --output examples/procurement/output \
  --format docx
```

## 选择自己的文件

双击 `tools/file_picker.cmd`，浏览器会打开本地选择页面。一次选择多份表格和文档，填写共同编号字段并选择报告格式即可核验。

也可以运行：

```bash
python tools/file_picker.py
```

选择页面只监听本机 `127.0.0.1`。所选文件会复制到已忽略提交的 `user_data/file-picker-uploads/`，报告写入 `user_data/file-picker-runs/`，原始文件不会被修改。详见 `tools/README.md`。

```bash
python multi-source-consistency-auditor/scripts/audit.py \
  --profile examples/procurement/profile.json \
  --output examples/procurement/output \
  --format pdf
```

## 示例输入

示例包含三类来源：

- `采购台账.xlsx`：三个采购项目；
- `采购申请.csv`：三个申请项目；
- `合同文件/`：两份 Word 合同和一份 PDF 合同。

数据是自动生成的虚构样例。项目编号 `P-2026-002` 的金额在采购申请和合同之间不同，其他项目包含来源缺失，用于验证报告是否能定位问题。

## 示例输出

```text
examples/procurement/output/
└── differences-summary.xlsx
```

差异汇总包含总体统计、实体对照、问题明细、来源清单和可追溯证据。JSON 结果适合与其他工具集成，CSV 可用于二次筛选。无论选择哪种格式，程序只生成新的 `differences-summary` 文件，不转换或覆盖源资料。

用户只需一种结果时，可以使用以下任一命令：

```bash
--format html
--format xlsx
--format docx
--format pdf
--format markdown
--format json
--format csv
```

也可重复传入 `--format`，或使用 `--format all`。

## 配置自己的核验

复制 `examples/procurement/profile.json`，修改：

1. `identifier`：实体编号字段和文档编号正则；
2. `sources`：每个来源的名称、类型和路径；
3. `fields`：业务字段、别名和类型；
4. `checks`：需要执行的字段一致性、必填和日期顺序规则。

详细格式见 [multi-source-consistency-auditor/references/profile-format.md](multi-source-consistency-auditor/references/profile-format.md)。

## 测试

```bash
python -m unittest discover -s tests -v
```

测试覆盖数值和日期标准化、多格式解析、跨来源金额冲突、来源缺失、重复编号、输出生成、非法配置和文件只读完整性。

## 安全与数据边界

- 工具不会修改、重命名、移动、转换格式或回写源文件。
- 第三方解析器不会直接打开原文件，只解析独立临时副本。
- 报告只能写入独立的输出目录；工具拒绝把输出写入被扫描的来源目录。
- 工具不访问网络，不会上传或公开任何资料。
- 报告中包含字段原文摘录，应按资料敏感等级控制查看和传递范围。
- 扫描版 PDF、图片 OCR 和任意自然语言语义推理不在核心范围内。
- 核验结果是规则计算结果，不替代业务、财务或法律复核。

## 项目结构

```text
multi-source-consistency-auditor/
├── multi-source-consistency-auditor/  # Agent Skill
├── examples/procurement/              # 虚构采购样例与配置
├── tools/                             # 本地多文件选择页面
├── tests/                             # 自动化测试
├── docs/                              # 测试记录与参赛说明
├── requirements.txt
├── pyproject.toml
├── LICENSE
└── README.md
```

## 许可证

本项目使用 MIT License。

PDF 报告优先使用可选的 Noto Sans SC 字体文件；公开源码仓库不包含该大体积字体，缺失时自动回退到系统 CJK 字体。字体使用 SIL Open Font License 1.1，许可文本位于同一 `assets` 目录。
