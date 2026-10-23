# 核验配置格式

配置使用 UTF-8 JSON。所有相对路径以配置文件所在目录为基准。

## 最小结构

```json
{
  "name": "资料一致性核验",
  "entity": "项目",
  "identifier": {
    "field": "project_id",
    "aliases": ["项目编号", "项目ID"],
    "pattern": "(?:P|PRJ)-\\d{4}-\\d{3}"
  },
  "sources": [
    {
      "name": "源表",
      "type": "table",
      "paths": ["input/清单.xlsx"],
      "expected": true
    },
    {
      "name": "附件",
      "type": "document",
      "paths": ["input/附件"],
      "include": ["*.docx", "*.pdf"],
      "expected": true
    }
  ],
  "fields": {
    "title": {
      "label": "名称",
      "aliases": ["项目名称", "名称"],
      "type": "text"
    },
    "amount": {
      "label": "金额",
      "aliases": ["合同金额", "金额"],
      "type": "amount"
    }
  },
  "tolerances": {"amount": "0.01"},
  "auto_consistent_fields": true,
  "coverage_severity": "warning",
  "checks": [
    {
      "type": "consistent",
      "field": "amount",
      "sources": ["源表", "附件"],
      "severity": "error",
      "tolerance": "0.01"
    },
    {
      "type": "required",
      "field": "title",
      "sources": ["源表"],
      "severity": "warning"
    },
    {
      "type": "date_order",
      "start": "application_date",
      "end": "signed_date",
      "severity": "warning"
    }
  ]
}
```

## 字段说明

- `identifier.field`：实体编号字段名，仅用于内部标识。
- `identifier.aliases`：表格列名或文档字段别名。表格优先完全匹配，再按包含关系匹配。
- `identifier.pattern`：可选正则。先从文件名查找，再从正文查找；适合“编号_合同.docx”。
- `sources[].name`：唯一来源名称，规则通过该名称引用。
- `sources[].type`：`auto`、`table` 或 `document`。`auto` 根据扩展名判断。
- `sources[].paths`：一个或多个文件、目录或 glob 表达式。
- `sources[].include`：目录来源的可选文件通配符。
- `sources[].expected`：为 `true` 时，每个实体缺少该来源会产生来源覆盖问题。
- `fields.*.type`：`text`、`amount`、`number`、`date`。
- `tolerances`：按字段名称设置默认数值容差。
- `auto_consistent_fields`：为未显式核验的字段自动执行跨来源一致性检查，默认 `true`。
- `coverage_severity`：来源缺失的严重程度，可选 `error`、`warning`、`info`。

## 检查类型

### consistent

比较同一实体在指定来源中的标准化字段值。数值检查支持容差。未指定 `sources` 时比较所有来源。

### required

若指定来源中存在该实体记录，但缺少字段，则产生问题。未指定 `sources` 时检查所有来源。

### date_order

比较同一实体所有可用 `start` 与 `end` 日期。只要最晚开始日期晚于最早结束日期，就产生问题。该规则适合发现申请晚于签署、开始晚于结束等时序矛盾。

## 文档字段格式

文档解析器识别以下形式：

```text
项目编号：P-2026-001
项目名称：办公电脑采购
供应商：示例供应商
合同金额：人民币 12,000.00 元
签订日期：2026年9月8日
```

同一字段别名在一份文档中优先采用首次出现的值。一个文档默认对应一个实体；一个实体可对应多份文档。

## 安全约束

- 不支持在源目录内生成报告。
- 不转换或回写源文件。
- 不访问网络。
- 报告含原文摘录，应按来源资料的敏感等级管理。
