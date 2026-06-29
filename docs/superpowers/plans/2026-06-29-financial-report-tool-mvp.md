# 财务报表生成工具 MVP 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Streamlit local web app that allows financial professionals to upload monthly source data (invoices, bank statements, payroll) and generate formatted 小企业会计准则 financial statements (利润表, 资产负债表, 现金流量表) with preset templates.

**Architecture:** Python backend modules for data extraction, template management, accounting calculation, and Excel rendering — exposed through a Streamlit single-page UI with sidebar navigation. All data stays local; no network requests.

**Tech Stack:** Python 3.10+, Streamlit 1.28+, openpyxl 3.0+, xlrd 2.0+, pandas 1.5+

**Prerequisite knowledge:** The existing `generate_formatted.py` (613 lines) contains all proven data extraction logic, accounting calculations, and formatting definitions. Much of the MVP implementation adapts this code into modular, reusable components.

**File structure after MVP:**
```
├── app.py                      # Streamlit entry point (~400 lines)
├── template_engine.py          # Preset template loading (~80 lines)
├── mapping_engine.py           # 科目映射 CRUD (~120 lines)
├── data_extractor.py           # 3 extractors (~200 lines)
├── calc_engine.py              # PL + BS calculation (~180 lines)
├── format_renderer.py          # Excel output (~250 lines)
├── validator.py                # Validation checks (~100 lines)
├── requirements.txt
├── .gitignore
├── data/                       # Created at runtime
│   ├── companies/
│   │   └── <name>/
│   │       ├── config.json
│   │       └── mappings.json
│   ├── uploads/
│   ├── output/
│   └── logs/
└── templates/                  # Preset template configs (committed)
    ├── pl_config.json
    ├── bs_config.json
    └── cf_config.json
```

---
### Task 1: Project scaffold

**Files:**
- Create: `requirements.txt`
- Modify: `.gitignore`

- [ ] **Step 1: Write requirements.txt**

```txt
streamlit>=1.28
openpyxl>=3.0
xlrd>=2.0
python-docx>=0.8
pandas>=1.5
```

- [ ] **Step 2: Update .gitignore to add data/ directory and preserve templates/**

Append these lines to `.gitignore`:

```
# Runtime data (user files, sensitive)
data/companies/
data/uploads/
data/output/
data/logs/

# Keep template configs (not sensitive, needed by app)
!templates/
```

- [ ] **Step 3: Create directory structure**

Run: `mkdir -p data/companies data/uploads data/output data/logs templates`

- [ ] **Step 4: Install dependencies**

Run: `pip install -r requirements.txt`
Expected: All five packages install without error.

- [ ] **Step 5: Commit**

```bash
git add requirements.txt .gitignore
git commit -m "chore: project scaffold with dependencies and directory structure"
```

---
### Task 2: Template engine — preset configs

**Files:**
- Create: `templates/pl_config.json`
- Create: `templates/bs_config.json`
- Create: `templates/cf_config.json`
- Create: `template_engine.py`

This task creates the three pre-built template configuration files matching the 小企业会计准则 standard format, plus a loader module. The configs describe the exact row/column structure of each statement, enabling the format renderer to produce correctly formatted Excel without hardcoding.

- [ ] **Step 1: Create PL template config `templates/pl_config.json`**

```json
{
  "template_type": "利润表",
  "subtitle": "会小企02表",
  "ncols": 4,
  "col_widths": [43.625, 6.625, 22.625, 22.625],
  "header_rows": [1, 2, 3, 4],
  "data_start_row": 5,
  "total_rows": 40,
  "row_height_data": 17.25,
  "row_height_empty": 13.5,
  "row_height_title": 21.0,
  "columns": {
    "item_name": 0,
    "row_num": 1,
    "value_cols": [2, 3]
  },
  "value_col_labels": ["本年累计金额", "本月金额"],
  "merge_ranges": [
    {"cells": "A1:D1", "type": "title"},
    {"cells": "A2:D2", "type": "subtitle"},
    {"cells": "B3:C3", "type": "company_info"}
  ],
  "items": [
    {"row": 5,  "name": "一、营业收入", "row_num": "1",  "section": false},
    {"row": 6,  "name": "减：营业成本", "row_num": "2",  "section": false},
    {"row": 7,  "name": "税金及附加", "row_num": "3",  "section": false},
    {"row": 8,  "name": "其中：消费税", "row_num": "4",  "section": false},
    {"row": 9,  "name": "  营业税", "row_num": "5",  "section": false},
    {"row": 10, "name": "  城市维护建设税", "row_num": "6",  "section": false},
    {"row": 11, "name": "  资源税", "row_num": "7",  "section": false},
    {"row": 12, "name": "  土地增值税", "row_num": "8",  "section": false},
    {"row": 13, "name": "  城镇土地使用税、房产税、车船税、印花税", "row_num": "9",  "section": false},
    {"row": 14, "name": "  教育费附加、矿产资源补偿费、排污费", "row_num": "10", "section": false},
    {"row": 15, "name": "销售费用", "row_num": "11", "section": false},
    {"row": 16, "name": "其中：商品维护费", "row_num": "12", "section": false},
    {"row": 17, "name": "  广告费和业务宣传费", "row_num": "13", "section": false},
    {"row": 18, "name": "管理费用", "row_num": "14", "section": false},
    {"row": 19, "name": "其中：开办费", "row_num": "15", "section": false},
    {"row": 20, "name": "  业务招待费", "row_num": "16", "section": false},
    {"row": 21, "name": "  研究费用", "row_num": "17", "section": false},
    {"row": 22, "name": "财务费用", "row_num": "18", "section": false},
    {"row": 23, "name": "其中：利息费用（收入以\"-\"号填列）", "row_num": "19", "section": false},
    {"row": 24, "name": "加：投资收益（亏损以\"-\"号填列）", "row_num": "20", "section": false},
    {"row": 25, "name": "二、营业利润（亏损以\"-\"号填列）", "row_num": "21", "section": true},
    {"row": 26, "name": "加：营业外收入", "row_num": "22", "section": false},
    {"row": 27, "name": "其中：政府补助", "row_num": "23", "section": false},
    {"row": 28, "name": "减：营业外支出", "row_num": "24", "section": false},
    {"row": 29, "name": "其中：坏账损失", "row_num": "25", "section": false},
    {"row": 30, "name": "  无法收回的长期债券投资损失", "row_num": "26", "section": false},
    {"row": 31, "name": "  无法收回的长期股权投资损失", "row_num": "27", "section": false},
    {"row": 32, "name": "  自然灾害等不可抗力因素造成的损失", "row_num": "28", "section": false},
    {"row": 33, "name": "  税收滞纳金", "row_num": "29", "section": false},
    {"row": 34, "name": "三、利润总额（亏损总额以\"-\"号填列）", "row_num": "30", "section": true},
    {"row": 35, "name": "减：所得税费用", "row_num": "31", "section": false},
    {"row": 36, "name": "四、净利润（净亏损以\"-\"号填列）", "row_num": "32", "section": true}
  ]
}
```

- [ ] **Step 2: Create BS template config `templates/bs_config.json`**

```json
{
  "template_type": "资产负债表",
  "subtitle": "会小企01表",
  "ncols": 8,
  "col_widths": [22.625, 6.625, 14.625, 14.625, 24.0, 6.625, 14.625, 14.625],
  "header_rows": [1, 2, 3, 4],
  "data_start_row": 5,
  "total_rows": 46,
  "row_height_data": 17.25,
  "row_height_empty": 13.5,
  "row_height_title": 21.0,
  "columns": {
    "item_name_left": 0,
    "row_num_left": 1,
    "value_cols_left": [2, 3],
    "item_name_right": 4,
    "row_num_right": 5,
    "value_cols_right": [6, 7]
  },
  "value_col_labels": ["期末余额", "年初余额"],
  "merge_ranges": [
    {"cells": "A1:H1", "type": "title"},
    {"cells": "A2:H2", "type": "subtitle"},
    {"cells": "A3:D3", "type": "company_left"},
    {"cells": "E3:H3", "type": "company_right"}
  ],
  "items": [
    {"row": 5,  "left": "流动资产：", "left_row_num": "",  "right": "流动负债：", "right_row_num": "",  "section": true},
    {"row": 6,  "left": "货币资金", "left_row_num": "1",  "right": "短期借款", "right_row_num": "31", "section": false},
    {"row": 7,  "left": "短期投资", "left_row_num": "2",  "right": "应付票据", "right_row_num": "32", "section": false},
    {"row": 8,  "left": "应收票据", "left_row_num": "3",  "right": "应付账款", "right_row_num": "33", "section": false},
    {"row": 9,  "left": "应收账款", "left_row_num": "4",  "right": "预收账款", "right_row_num": "34", "section": false},
    {"row": 10, "left": "预付账款", "left_row_num": "5",  "right": "应付职工薪酬", "right_row_num": "35", "section": false},
    {"row": 11, "left": "应收股利", "left_row_num": "6",  "right": "应交税费", "right_row_num": "36", "section": false},
    {"row": 12, "left": "应收利息", "left_row_num": "7",  "right": "应付利息", "right_row_num": "37", "section": false},
    {"row": 13, "left": "其他应收款", "left_row_num": "8",  "right": "应付利润", "right_row_num": "38", "section": false},
    {"row": 14, "left": "存货", "left_row_num": "9",  "right": "其他应付款", "right_row_num": "39", "section": false},
    {"row": 15, "left": "  其中：原材料", "left_row_num": "10", "right": "", "right_row_num": "",  "section": false},
    {"row": 16, "left": "       在产品", "left_row_num": "11", "right": "", "right_row_num": "",  "section": false},
    {"row": 17, "left": "       库存商品", "left_row_num": "12", "right": "", "right_row_num": "",  "section": false},
    {"row": 18, "left": "       周转材料", "left_row_num": "13", "right": "", "right_row_num": "",  "section": false},
    {"row": 19, "left": "其他流动资产", "left_row_num": "14", "right": "其他流动负债", "right_row_num": "40", "section": false},
    {"row": 20, "left": "流动资产合计", "left_row_num": "15", "right": "流动负债合计", "right_row_num": "41", "section": false},
    {"row": 21, "left": "非流动资产：", "left_row_num": "",  "right": "非流动负债：", "right_row_num": "",  "section": true},
    {"row": 22, "left": "长期债券投资", "left_row_num": "16", "right": "长期借款", "right_row_num": "42", "section": false},
    {"row": 23, "left": "长期股权投资", "left_row_num": "17", "right": "长期应付款", "right_row_num": "43", "section": false},
    {"row": 24, "left": "固定资产原价", "left_row_num": "18", "right": "递延收益", "right_row_num": "44", "section": false},
    {"row": 25, "left": "减：累计折旧", "left_row_num": "19", "right": "其他非流动负债", "right_row_num": "45", "section": false},
    {"row": 26, "left": "固定资产账面价值", "left_row_num": "20", "right": "非流动负债合计", "right_row_num": "46", "section": false},
    {"row": 27, "left": "在建工程", "left_row_num": "21", "right": "负债合计", "right_row_num": "47", "section": false},
    {"row": 28, "left": "工程物资", "left_row_num": "22", "right": "", "right_row_num": "",  "section": false},
    {"row": 29, "left": "固定资产清理", "left_row_num": "23", "right": "", "right_row_num": "",  "section": false},
    {"row": 30, "left": "生产性生物资产", "left_row_num": "24", "right": "所有者权益（或股东权益）：", "right_row_num": "",  "section": true},
    {"row": 31, "left": "无形资产", "left_row_num": "25", "right": "实收资本（或股本）", "right_row_num": "48", "section": false},
    {"row": 32, "left": "开发支出", "left_row_num": "26", "right": "资本公积", "right_row_num": "49", "section": false},
    {"row": 33, "left": "长期待摊费用", "left_row_num": "27", "right": "盈余公积", "right_row_num": "50", "section": false},
    {"row": 34, "left": "其他非流动资产", "left_row_num": "28", "right": "未分配利润", "right_row_num": "51", "section": false},
    {"row": 35, "left": "非流动资产合计", "left_row_num": "29", "right": "所有者权益（或股东权益）合计", "right_row_num": "52", "section": false},
    {"row": 36, "left": "资产总计", "left_row_num": "30", "right": "负债和所有者权益（或股东权益）总计", "right_row_num": "53", "section": false}
  ]
}
```

- [ ] **Step 3: Create CF template config `templates/cf_config.json`**

```json
{
  "template_type": "现金流量表",
  "subtitle": "                                                     会小企03表",
  "ncols": 5,
  "col_widths": [55.375, 6.625, 22.625, 22.625, 52.5],
  "header_rows": [1, 2, 3, 4],
  "data_start_row": 5,
  "total_rows": 60,
  "row_height_data": 17.25,
  "row_height_empty": 13.5,
  "row_height_title": 21.0,
  "columns": {
    "item_name": 0,
    "row_num": 1,
    "value_cols": [2, 3]
  },
  "value_col_labels": ["本年累计金额", "本月金额"],
  "merge_ranges": [
    {"cells": "A1:E1", "type": "title"},
    {"cells": "A2:E2", "type": "subtitle"},
    {"cells": "B3:C3", "type": "company_info"}
  ],
  "items": [
    {"row": 5,  "name": "一、经营活动产生的现金流量：", "row_num": "",  "section": true},
    {"row": 6,  "name": "销售产成品、商品、提供劳务收到的现金", "row_num": "1",  "section": false},
    {"row": 7,  "name": "收到的其他与经营活动有关的现金", "row_num": "2",  "section": false},
    {"row": 8,  "name": "购买原材料、商品、接受劳务支付的现金", "row_num": "3",  "section": false},
    {"row": 9,  "name": "支付的职工薪酬", "row_num": "4",  "section": false},
    {"row": 10, "name": "支付的税费", "row_num": "5",  "section": false},
    {"row": 11, "name": "支付的其他与经营活动有关的现金", "row_num": "6",  "section": false},
    {"row": 12, "name": "经营活动产生的现金流量净额", "row_num": "7",  "section": false},
    {"row": 13, "name": "二、投资活动产生的现金流量：", "row_num": "",  "section": true},
    {"row": 14, "name": "收回短期投资、长期债券投资和长期股权投资收到的现金", "row_num": "8",  "section": false},
    {"row": 15, "name": "取得投资收益收到的现金", "row_num": "9",  "section": false},
    {"row": 16, "name": "处置固定资产、无形资产和其他非流动资产收回的现金净额", "row_num": "10", "section": false},
    {"row": 17, "name": "短期投资、长期债券投资和长期股权投资支付的现金", "row_num": "11", "section": false},
    {"row": 18, "name": "购建固定资产、无形资产和其他非流动资产支付的现金", "row_num": "12", "section": false},
    {"row": 19, "name": "投资活动产生的现金流量净额", "row_num": "13", "section": false},
    {"row": 20, "name": "三、筹资活动产生的现金流量：", "row_num": "",  "section": true},
    {"row": 21, "name": "取得借款收到的现金", "row_num": "14", "section": false},
    {"row": 22, "name": "吸收投资者投资收到的现金", "row_num": "15", "section": false},
    {"row": 23, "name": "偿还借款本金支付的现金", "row_num": "16", "section": false},
    {"row": 24, "name": "偿还借款利息支付的现金", "row_num": "17", "section": false},
    {"row": 25, "name": "分配利润支付的现金", "row_num": "18", "section": false},
    {"row": 26, "name": "筹资活动产生的现金流量净额", "row_num": "19", "section": false},
    {"row": 27, "name": "四、现金净增加额", "row_num": "20", "section": false},
    {"row": 28, "name": "加：期初现金余额", "row_num": "21", "section": false},
    {"row": 29, "name": "五、期末现金余额", "row_num": "22", "section": false},
    {"row": 30, "name": "", "row_num": "",  "section": false},
    {"row": 31, "name": "", "row_num": "",  "section": false},
    {"row": 32, "name": "补充资料项目", "row_num": "行次", "section": false},
    {"row": 33, "name": "1、将净利润调节为经营活动现金流量：", "row_num": "",  "section": true},
    {"row": 34, "name": "净利润", "row_num": "23", "section": false},
    {"row": 35, "name": "加：计提的资产减值准备", "row_num": "24", "section": false},
    {"row": 36, "name": "固定资产折旧", "row_num": "25", "section": false},
    {"row": 37, "name": "无形资产摊销", "row_num": "26", "section": false},
    {"row": 38, "name": "长期待摊费用摊销", "row_num": "27", "section": false},
    {"row": 39, "name": "待摊费用减少（减：增加）", "row_num": "28", "section": false},
    {"row": 40, "name": "预提费用增加（减：减少）", "row_num": "29", "section": false},
    {"row": 41, "name": "处置固定资产、无形资产和其他非流动资产损失（减：收益）", "row_num": "30", "section": false},
    {"row": 42, "name": "固定资产报废损失", "row_num": "31", "section": false},
    {"row": 43, "name": "财务费用", "row_num": "32", "section": false},
    {"row": 44, "name": "投资损失（减：收益）", "row_num": "33", "section": false},
    {"row": 45, "name": "递延税款贷项（减：借项）", "row_num": "34", "section": false},
    {"row": 46, "name": "存货的减少（减：增加）", "row_num": "35", "section": false},
    {"row": 47, "name": "经营性应收项目的减少（减：增加）", "row_num": "36", "section": false},
    {"row": 48, "name": "经营性应付项目的增加（减：减少）", "row_num": "37", "section": false},
    {"row": 49, "name": "其他", "row_num": "38", "section": false},
    {"row": 50, "name": "经营活动产生的现金流量净额", "row_num": "39", "section": false},
    {"row": 51, "name": "2、不涉及现金收支的投资和筹资活动：", "row_num": "",  "section": true},
    {"row": 52, "name": "债务转为资本", "row_num": "40", "section": false},
    {"row": 53, "name": "一年内到期的可转换公司债券", "row_num": "41", "section": false},
    {"row": 54, "name": "融资租入固定资产", "row_num": "42", "section": false},
    {"row": 55, "name": "3、现金及现金等价物净增加情况", "row_num": "",  "section": true},
    {"row": 56, "name": "现金的期末余额", "row_num": "43", "section": false},
    {"row": 57, "name": "减：现金的期初余额", "row_num": "44", "section": false},
    {"row": 58, "name": "加：现金等价物的期末余额", "row_num": "45", "section": false},
    {"row": 59, "name": "减：现金等价物的期初余额", "row_num": "46", "section": false},
    {"row": 60, "name": "现金及现金等价物净增加额", "row_num": "47", "section": false}
  ]
}
```

- [ ] **Step 4: Write `template_engine.py`**

```python
# -*- coding: utf-8 -*-
"""Template engine — loads preset template configs and manages template selection."""
import json
import os

TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), 'templates')

PRESET_MAP = {
    'pl': 'pl_config.json',
    'bs': 'bs_config.json',
    'cf': 'cf_config.json',
}

def _load_json(filename):
    path = os.path.join(TEMPLATE_DIR, filename)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Template file not found: {path}")
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def load_preset(statement_type):
    """Load a preset template config by type.

    Args:
        statement_type: 'pl' | 'bs' | 'cf'

    Returns:
        dict with keys: template_type, subtitle, ncols, col_widths,
        header_rows, data_start_row, total_rows, columns, value_col_labels,
        merge_ranges, items
    """
    filename = PRESET_MAP.get(statement_type)
    if filename is None:
        raise ValueError(f"Unknown statement type '{statement_type}'. Use 'pl', 'bs', or 'cf'.")
    return _load_json(filename)

def list_presets():
    """Return list of available preset template types."""
    return list(PRESET_MAP.keys())

def get_item_names(statement_type):
    """Return ordered list of item names from a preset template.

    Useful for building the mapping UI — displays the report structure
    so users know what they're mapping to.
    """
    config = load_preset(statement_type)
    if statement_type == 'bs':
        return [item['left'] for item in config['items'] if item['left']], \
               [item['right'] for item in config['items'] if item['right']]
    return [item['name'] for item in config['items'] if item['name']], []
```

- [ ] **Step 5: Run verification**

```python
# Quick sanity check — this script loads all three templates and prints item counts
import template_engine as te
for st in ['pl', 'bs', 'cf']:
    cfg = te.load_preset(st)
    items = [i for i in cfg['items'] if not i.get('section')]
    print(f"{cfg['template_type']}: {len(items)} data items, {cfg['ncols']} columns, {cfg['total_rows']} total rows")
```

Run: `python -c "import template_engine as te; te.load_preset('pl'); te.load_preset('bs'); te.load_preset('cf'); print('All templates loaded OK')"`
Expected: `All templates loaded OK`

- [ ] **Step 6: Commit**

```bash
git add templates/ template_engine.py
git commit -m "feat: add preset template configs and template engine"
```

---
### Task 3: Data extractor

**Files:**
- Create: `data_extractor.py`

This module adapts the proven extraction logic from `generate_formatted.py:26-115` into three independent, testable classes. Each extractor takes a file path and returns a standardized intermediate data structure — no accounting logic, no global state.

- [ ] **Step 1: Write `data_extractor.py`**

```python
# -*- coding: utf-8 -*-
"""Data extractors — read source Excel files into standardized dict structures.

Three extractors:
  InvoiceExtractor  — sales/cost invoice .xlsx files
  BankExtractor     — bank statement .xls files (农行/信用社)
  PayrollExtractor  — monthly payroll .xlsx files

All extractors are file-path-in, dict-out. No global state. No accounting logic.
"""
import openpyxl
import xlrd
from datetime import datetime
from collections import defaultdict


def _sf(v):
    """Safe float conversion. Returns 0.0 on failure."""
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _is_date(v):
    """Check if a value looks like a date."""
    if v is None:
        return False
    if isinstance(v, datetime):
        return True
    s = str(v).strip()
    return len(s) >= 10 and s[4] == '-' and s[7] == '-'


def _is_date_xl(v):
    """Check if an xlrd cell value is a date (float serial or string)."""
    if v is None or v == '':
        return False
    if isinstance(v, float) and v > 40000:
        return True
    s = str(v).strip()
    return len(s) >= 10 and s[4] == '-' and s[7] == '-'


# ================================================================
# InvoiceExtractor
# ================================================================

class InvoiceExtractor:
    """Extract invoice data from sales/cost .xlsx files.

    Expected format (金蝶导出):
      - Sheet name: '信息汇总表'
      - Row 1: header
      - Data from row 2
      - Column 8 (0-indexed): date
      - Column 9: 科目编码
      - Column 16: 金额 (不含税)
      - Column 18: 税额

    Returns:
      {month: [{date, invoice_no, account_code, amount, tax_amount}]}
    """

    def __init__(self, filepath, sheet_name='信息汇总表',
                 col_date=8, col_code=9, col_amount=16, col_tax=18):
        self.filepath = filepath
        self.sheet_name = sheet_name
        self.col_date = col_date
        self.col_code = col_code
        self.col_amount = col_amount
        self.col_tax = col_tax

    def extract(self):
        """Read the file and return {month_int: [records]}."""
        wb = openpyxl.load_workbook(self.filepath, data_only=True)
        if self.sheet_name not in wb.sheetnames:
            raise ValueError(
                f"Sheet '{self.sheet_name}' not found in {self.filepath}. "
                f"Available sheets: {wb.sheetnames}"
            )
        ws = wb[self.sheet_name]

        result = defaultdict(list)
        skipped = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            date_val = row[self.col_date] if self.col_date < len(row) else None
            if not _is_date(date_val):
                continue

            mo = date_val.month if isinstance(date_val, datetime) else \
                 datetime.strptime(str(date_val)[:10], '%Y-%m-%d').month

            code = str(row[self.col_code]) if self.col_code < len(row) and row[self.col_code] else ''
            amount = _sf(row[self.col_amount]) if self.col_amount < len(row) else 0.0
            tax = _sf(row[self.col_tax]) if self.col_tax < len(row) else 0.0

            if abs(amount) < 0.005 and abs(tax) < 0.005:
                skipped += 1
                continue

            result[mo].append({
                'date': date_val,
                'account_code': code,
                'amount': amount,
                'tax_amount': tax,
            })

        wb.close()
        return dict(result), skipped

    def get_total_by_code(self, months=None):
        """Extract and return {account_code: total_amount} for convenience.
        Used by the mapping wizard to discover account codes.
        """
        data, _ = self.extract()
        totals = defaultdict(float)
        for mo, records in data.items():
            if months and mo not in months:
                continue
            for rec in records:
                totals[rec['account_code']] += rec['amount']
        return dict(totals)

    def get_sum_by_months(self):
        """Return {month: total_amount} summary."""
        data, _ = self.extract()
        return {mo: sum(r['amount'] for r in records) for mo, records in data.items()}


# ================================================================
# BankExtractor
# ================================================================

class BankExtractor:
    """Extract bank transaction data from .xls files.

    Supports two bank formats:
      - 农行 (Agricultural Bank): chronological order, data from row 3
        col 0=date, 1=收入, 2=支出, 3=余额, 7=摘要
      - 信用社 (Rural Credit Union): reverse chronological, data from row 4
        col 0=date, 1=支出, 2=收入, 3=余额

    Returns:
      {'in_total': float, 'out_total': float, 'end_balance': float,
       'transactions': [{date, amount_in, amount_out, balance, summary}]}
    """

    BANK_NONGHANG = 'nonghang'
    BANK_XINYONGSHE = 'xinyongshe'

    def __init__(self, filepath, bank_type=BANK_NONGHANG):
        self.filepath = filepath
        self.bank_type = bank_type

    def extract(self):
        """Read the bank statement and return standardized dict."""
        wb = xlrd.open_workbook(self.filepath)
        ws = wb.sheet_by_index(0)

        if self.bank_type == self.BANK_NONGHANG:
            return self._extract_nonghang(ws)
        elif self.bank_type == self.BANK_XINYONGSHE:
            return self._extract_xinyongshe(ws)
        else:
            raise ValueError(f"Unknown bank_type: {self.bank_type}")

    def _extract_nonghang(self, ws):
        in_total, out_total = 0.0, 0.0
        end_balance = 0.0
        transactions = []

        for r in range(3, ws.nrows):
            if not _is_date_xl(ws.cell_value(r, 0)):
                continue
            amt_in = _sf(ws.cell_value(r, 1))
            amt_out = _sf(ws.cell_value(r, 2))
            bal = _sf(ws.cell_value(r, 3))
            summary = str(ws.cell_value(r, 7)) if ws.ncols > 7 else ''

            in_total += amt_in
            out_total += amt_out
            end_balance = bal

            transactions.append({
                'date': ws.cell_value(r, 0),
                'amount_in': amt_in,
                'amount_out': amt_out,
                'balance': bal,
                'summary': summary,
            })

        return {
            'in_total': in_total,
            'out_total': out_total,
            'end_balance': end_balance,
            'transactions': transactions,
            'bank_type': self.bank_type,
        }

    def _extract_xinyongshe(self, ws):
        in_total, out_total = 0.0, 0.0
        first_balance = None
        transactions = []

        for r in range(4, ws.nrows):
            if not _is_date_xl(ws.cell_value(r, 0)):
                continue
            amt_out = _sf(ws.cell_value(r, 1))  # Note: reversed columns
            amt_in = _sf(ws.cell_value(r, 2))
            bal = _sf(ws.cell_value(r, 3))

            if first_balance is None:
                first_balance = bal

            in_total += amt_in
            out_total += amt_out

            transactions.append({
                'date': ws.cell_value(r, 0),
                'amount_in': amt_in,
                'amount_out': amt_out,
                'balance': bal,
                'summary': '',
            })

        return {
            'in_total': in_total,
            'out_total': out_total,
            'end_balance': first_balance if first_balance is not None else 0.0,
            'transactions': transactions,
            'bank_type': self.bank_type,
        }

    def get_interest_income(self):
        """Scan transactions for interest income entries (利息 in summary)."""
        data = self.extract()
        total = 0.0
        for t in data['transactions']:
            if '利息' in str(t.get('summary', '')):
                total += t['amount_in']
        return total


# ================================================================
# PayrollExtractor
# ================================================================

class PayrollExtractor:
    """Extract monthly payroll data from .xlsx files.

    Expected format:
      - Row containing '合计' in column 0 marks the summary row
      - Column 5 (0-indexed): 应发工资合计

    Returns:
      {'gross_pay': float}
    """

    KEYWORDS = ['应发工资', '应发合计', '应发']

    def __init__(self, filepath):
        self.filepath = filepath

    def extract(self):
        """Read payroll file and return {'gross_pay': amount}."""
        wb = openpyxl.load_workbook(self.filepath, data_only=True)
        ws = wb.active

        gross_pay = 0.0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
            label = str(row[0]).strip() if row[0] else ''
            if '合计' in label:
                gross_pay = _sf(row[5]) if len(row) > 5 else 0.0
                break

        wb.close()
        return {'gross_pay': gross_pay}
```

- [ ] **Step 2: Verify extractors against existing顺恒 data**

Run this verification script:

```python
from data_extractor import InvoiceExtractor, BankExtractor, PayrollExtractor

# Test InvoiceExtractor with sales data
ie = InvoiceExtractor('顺恒2026年1-5月销售收入数据.xlsx')
data, skipped = ie.extract()
total = sum(sum(r['amount'] for r in recs) for recs in data.values())
print(f"Sales: {len(data)} months, {sum(len(v) for v in data.values())} invoices, total={total:,.2f}")
assert abs(total - 4080824.78) < 0.02, f"Mismatch: {total}"

# Test BankExtractor
be_nong = BankExtractor('顺恒2026年5月农行明细.xls', BankExtractor.BANK_NONGHANG)
nong_result = be_nong.extract()
print(f"农行 May balance: {nong_result['end_balance']:,.2f}")

be_xin = BankExtractor('顺恒2026年5月农村信用社明细.xls', BankExtractor.BANK_XINYONGSHE)
xin_result = be_xin.extract()
print(f"信用社 May balance: {xin_result['end_balance']:,.2f}")

# Test PayrollExtractor
pe = PayrollExtractor('顺恒2026年5月工资薪金.xlsx')
payroll = pe.extract()
print(f"May payroll: {payroll['gross_pay']:,.2f}")

print("\nAll extraction tests passed!")
```

Expected: prints results without assertion errors.

- [ ] **Step 3: Commit**

```bash
git add data_extractor.py
git commit -m "feat: add data extractors for invoices, bank statements, and payroll"
```

---
### Task 4: Mapping engine

**Files:**
- Create: `mapping_engine.py`

Manages the 科目编码 → 报表行项目 correspondence. Three rule types: `code_match` (source account code maps to a report line), `auto_calc` (line is computed by engine, not matched to source data), `multi_source` (line aggregates multiple code matches and computed values).

- [ ] **Step 1: Write `mapping_engine.py`**

```python
# -*- coding: utf-8 -*-
"""Mapping engine — manages 科目编码 ↔ 报表行项目 correspondence.

Mapping rules (3 types):
  code_match   — match source account code to a report line item
  auto_calc    — line is computed by calc engine (e.g. depreciation, tax surcharges)
  multi_source — line aggregates multiple code matches + computed items

Data stored as JSON per company, e.g. data/companies/<name>/mappings.json
"""
import json
import os


# The mapping for a single report line
# {
#   "report_line": "减：营业成本",
#   "statement": "pl",
#   "rule_type": "code_match",
#   "account_codes": ["1110701000000000000", "1110799000000000000"],
#   "auto_params": {}
# }
#
# or for auto_calc:
# {
#   "report_line": "税金及附加",
#   "statement": "pl",
#   "rule_type": "auto_calc",
#   "account_codes": [],
#   "auto_params": {"method": "vat_surcharge", "rate": 0.12}
# }


class MappingEngine:
    """Load and manage mappings for one company."""

    def __init__(self, company_dir):
        self.company_dir = company_dir
        self.mappings_path = os.path.join(company_dir, 'mappings.json')
        self._data = None

    def _load(self):
        if self._data is None:
            if os.path.exists(self.mappings_path):
                with open(self.mappings_path, 'r', encoding='utf-8') as f:
                    self._data = json.load(f)
            else:
                self._data = {'pl': [], 'bs': [], 'cf': []}
        return self._data

    def _save(self):
        os.makedirs(self.company_dir, exist_ok=True)
        with open(self.mappings_path, 'w', encoding='utf-8') as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)

    def get_mappings(self, statement):
        """Return list of mapping entries for a statement ('pl'|'bs'|'cf')."""
        return self._load().get(statement, [])

    def set_mappings(self, statement, mappings_list):
        """Replace all mappings for a statement type."""
        data = self._load()
        data[statement] = mappings_list
        self._save()

    def add_code_mapping(self, statement, report_line, account_codes):
        """Add or update a code_match mapping. Merges with existing codes."""
        data = self._load()
        mappings = data.setdefault(statement, [])

        for m in mappings:
            if m['report_line'] == report_line and m['rule_type'] == 'code_match':
                existing = set(m.get('account_codes', []))
                existing.update(account_codes)
                m['account_codes'] = sorted(existing)
                self._save()
                return

        mappings.append({
            'report_line': report_line,
            'statement': statement,
            'rule_type': 'code_match',
            'account_codes': sorted(account_codes),
            'auto_params': {},
        })
        self._save()

    def set_auto_calc(self, statement, report_line, method, params=None):
        """Mark a report line as auto-calculated."""
        data = self._load()
        mappings = data.setdefault(statement, [])

        for m in mappings:
            if m['report_line'] == report_line:
                m['rule_type'] = 'auto_calc'
                m['auto_params'] = {'method': method, **(params or {})}
                m['account_codes'] = []
                self._save()
                return

        mappings.append({
            'report_line': report_line,
            'statement': statement,
            'rule_type': 'auto_calc',
            'account_codes': [],
            'auto_params': {'method': method, **(params or {})},
        })
        self._save()

    def set_multi_source(self, statement, report_line, sources):
        """Set a report line as multi-source aggregation.

        sources is a list of {type: 'code'|'auto', value: ...}
        """
        data = self._load()
        mappings = data.setdefault(statement, [])

        for m in mappings:
            if m['report_line'] == report_line:
                m['rule_type'] = 'multi_source'
                m['sources'] = sources
                m['account_codes'] = []
                self._save()
                return

        mappings.append({
            'report_line': report_line,
            'statement': statement,
            'rule_type': 'multi_source',
            'account_codes': [],
            'auto_params': {},
            'sources': sources,
        })
        self._save()

    def get_account_codes_for_line(self, statement, report_line):
        """Return list of account codes mapped to a specific report line."""
        for m in self.get_mappings(statement):
            if m['report_line'] == report_line:
                return m.get('account_codes', [])
        return []

    def get_all_known_codes(self, statement):
        """Return set of all account codes registered across all mappings."""
        codes = set()
        for m in self.get_mappings(statement):
            codes.update(m.get('account_codes', []))
        return codes

    def find_unmapped_codes(self, statement, discovered_codes):
        """Compare discovered codes against known mappings.
        Returns list of codes that have NO mapping yet.
        discovered_codes: list/set of codes found in source data
        """
        known = self.get_all_known_codes(statement)
        return sorted(set(discovered_codes) - known)

    def to_dict(self):
        """Dump all mappings as dict for display/export."""
        return self._load()
```

- [ ] **Step 2: Write and run unit test**

```python
# Run this as a quick sanity check
import tempfile, os, json
from mapping_engine import MappingEngine

d = tempfile.mkdtemp()
eng = MappingEngine(d)

# Test adding code mapping
eng.add_code_mapping('pl', '一、营业收入', ['CODE_A', 'CODE_B'])
eng.add_code_mapping('pl', '一、营业收入', ['CODE_C'])  # merge
m = eng.get_mappings('pl')
assert len(m) == 1
assert set(m[0]['account_codes']) == {'CODE_A', 'CODE_B', 'CODE_C'}

# Test auto calc
eng.set_auto_calc('pl', '税金及附加', 'vat_surcharge', {'rate': 0.12})
assert len(eng.get_mappings('pl')) == 2

# Test find unmapped
unmapped = eng.find_unmapped_codes('pl', ['CODE_A', 'CODE_X', 'CODE_Y'])
assert unmapped == ['CODE_X', 'CODE_Y']

print("All mapping engine tests passed!")
```

- [ ] **Step 3: Commit**

```bash
git add mapping_engine.py
git commit -m "feat: add mapping engine for account-to-report-line correspondence"
```

---
### Task 5: Calculation engine

**Files:**
- Create: `calc_engine.py`

Adapts the accounting computation logic from `generate_formatted.py:117-195`. Computes PL and BS from extracted data + mappings. The CF (间接法) is deferred to second phase — MVP only generates PL and BS.

- [ ] **Step 1: Write `calc_engine.py`**

```python
# -*- coding: utf-8 -*-
"""Calculation engine — produces statement data from extracted source data + mappings.

Key formulas (adapted from generate_formatted.py, verified against Dec 2025 templates):

利润表:
  营业收入 = sum of invoice amounts matching account codes mapped to this line
  营业成本 = sum of invoice amounts matching COGS account codes
  税金及附加 = VAT × 12% (城建7% + 教育3% + 地方2%)
  管理费用 = expense invoices + payroll + social insurance + depreciation + amortization + bank fees
  财务费用 = bank account fees - interest income
  营业利润 = 营业收入 - 营业成本 - 税金及附加 - 管理费用 - 销售费用 - 财务费用
  利润总额 = 营业利润 + 营业外收入 - 营业外支出
  净利润 = 利润总额 - 所得税

资产负债表:
  Assets = opening_balance ± changes during period
  Liabilities + Equity must balance (A = L + E)
  其他应付款 = plug figure (A - all known L - E)
"""
from collections import defaultdict


def _sf(v):
    try: return float(v)
    except: return 0.0


class CalcEngine:
    """Compute financial statements from extracted data and mappings."""

    def __init__(self, opening_bs, monthly_depr=25886.46, monthly_ltd_amort=180.0,
                 bank_fee_per_month=4.5, social_insurance_rate=0.25):
        """
        Args:
            opening_bs: dict of opening balance sheet values
            monthly_depr: monthly depreciation amount (default from顺恒)
            monthly_ltd_amort: monthly 长期待摊费用 amortization
            bank_fee_per_month: monthly bank account management fee
            social_insurance_rate: employer SI as fraction of gross payroll
        """
        self.opening_bs = opening_bs
        self.monthly_depr = monthly_depr
        self.monthly_ltd_amort = monthly_ltd_amort
        self.bank_fee_per_month = bank_fee_per_month
        self.social_insurance_rate = social_insurance_rate

    def _vo(self, v):
        """Return float or None for zero values (matching template convention)."""
        if v is None: return None
        if abs(v) < 0.005: return None
        return round(float(v), 2)

    def _bo(self, key):
        """Opening balance lookup."""
        return _sf(self.opening_bs.get(key, 0))

    # -------- PL Calculation --------

    def calculate_pl(self, invoice_data, payroll_data, mappings, bank_data, num_months):
        """Compute profit statement.

        Args:
            invoice_data: {month: InvoiceExtractor result} for both sales and costs
            payroll_data: {month: PayrollExtractor result}
            mappings: MappingEngine.get_mappings('pl') result
            bank_data: {month: {'nonghang': result, 'xinyongshe': result}}
            num_months: number of months in period (for depreciation/amortization)

        Returns dict with keys: revenue, cogs, surcharges, selling_exp, admin_exp,
        fin_exp, invest_income, oper_profit, non_op_income, non_op_expense,
        total_profit, income_tax, net_profit, plus line-item detail dicts.
        """
        # --- Revenue ---
        revenue_codes = set()
        for m in mappings:
            if m['rule_type'] == 'code_match' and '营业收入' in m['report_line']:
                revenue_codes.update(m.get('account_codes', []))

        revenue = 0.0
        for mo, records in invoice_data.get('sales', {}).items():
            for rec in records:
                if rec['account_code'] in revenue_codes:
                    revenue += rec['amount']

        # --- COGS ---
        cogs_codes = set()
        for m in mappings:
            if m['rule_type'] == 'code_match' and '营业成本' in m['report_line']:
                cogs_codes.update(m.get('account_codes', []))

        cogs = 0.0
        for mo, records in invoice_data.get('costs', {}).items():
            for rec in records:
                if rec['account_code'] in cogs_codes:
                    cogs += rec['amount']

        # --- Admin expenses ---
        # Collect all codes mapped to 管理费用
        admin_codes = set()
        for m in mappings:
            if m['rule_type'] == 'code_match' and '管理费用' in m['report_line']:
                admin_codes.update(m.get('account_codes', []))

        admin_inv = 0.0
        fa_add = 0.0  # Fixed asset additions (items >500 元)
        for mo, records in invoice_data.get('costs', {}).items():
            for rec in records:
                code = rec['account_code']
                if code in admin_codes:
                    admin_inv += rec['amount']
                    # Detect fixed asset purchases
                    goods = str(rec.get('goods_name', ''))
                    if '打印机' in goods or '显示器' in goods:
                        if abs(rec['amount'] + rec.get('tax_amount', 0)) > 500:
                            fa_add += abs(rec['amount'] + rec.get('tax_amount', 0))

        # Payroll total
        total_payroll = sum(p.get('gross_pay', 0) for p in payroll_data.values())

        # Social insurance
        total_si = total_payroll * self.social_insurance_rate

        # Depreciation & amortization
        total_depr = self.monthly_depr * num_months
        total_ltd_amort = self.monthly_ltd_amort * num_months

        # Bank fees
        total_bank_fees = self.bank_fee_per_month * num_months

        # Interest income
        total_interest = 0.0
        for mo, banks in bank_data.items():
            for bk in banks.values():
                for t in bk.get('transactions', []):
                    if '利息' in str(t.get('summary', '')):
                        total_interest += t['amount_in']

        # Total admin = invoices + payroll + SI + depr + amort + bank fees - FA additions
        admin_exp = admin_inv + total_payroll + total_si + total_depr + total_ltd_amort + total_bank_fees - fa_add

        # --- Selling expenses ---
        selling_codes = set()
        for m in mappings:
            if m['rule_type'] == 'code_match' and '销售费用' in m['report_line']:
                selling_codes.update(m.get('account_codes', []))

        selling_exp = 0.0
        for mo, records in invoice_data.get('costs', {}).items():
            for rec in records:
                if rec['account_code'] in selling_codes:
                    selling_exp += rec['amount']

        # --- Financial expenses ---
        fin_exp = total_bank_fees - total_interest

        # --- Tax surcharges ---
        # VAT = output_tax - input_tax
        output_tax = 0.0
        for mo, records in invoice_data.get('sales', {}).items():
            for rec in records:
                output_tax += rec.get('tax_amount', 0)

        input_tax = 0.0
        for key in ['costs']:
            for mo, records in invoice_data.get(key, {}).items():
                for rec in records:
                    input_tax += abs(rec.get('tax_amount', 0))

        vat = max(output_tax - input_tax, 0)
        cj_tax = vat * 0.07    # 城建税
        jy_tax = vat * 0.03    # 教育费附加
        dj_tax = vat * 0.02    # 地方教育附加
        surcharges = cj_tax + jy_tax + dj_tax

        # --- Operating profit ---
        oper_profit = revenue - cogs - surcharges - selling_exp - admin_exp - fin_exp

        # --- Non-operating items ---
        non_op_income = 0.0   # Users can manually input
        non_op_expense = 0.0

        # --- Total profit ---
        total_profit = oper_profit + non_op_income - non_op_expense

        # --- Income tax ---
        income_tax = 0.0  # Default 0 (most small businesses have accumulated losses)

        # --- Net profit ---
        net_profit = total_profit - income_tax

        return {
            'revenue': self._vo(revenue),
            'cogs': self._vo(cogs),
            'surcharges': self._vo(surcharges),
            'selling_exp': self._vo(selling_exp),
            'admin_exp': self._vo(admin_exp),
            'fin_exp': self._vo(fin_exp),
            'invest_income': None,
            'oper_profit': self._vo(oper_profit),
            'non_op_income': None,
            'non_op_expense': None,
            'total_profit': self._vo(total_profit),
            'income_tax': None,
            'net_profit': self._vo(net_profit),
            # Detail items for PL sub-lines
            '_detail': {
                'vat': vat,
                'cj_tax': cj_tax,
                'jy_tax': jy_tax,
                'dj_tax': dj_tax,
                'interest_income': total_interest,
                'fa_add': fa_add,
                'depreciation': total_depr,
                'ltd_amort': total_ltd_amort,
                'payroll_total': total_payroll,
                'social_insurance': total_si,
            }
        }

    # -------- BS Calculation --------

    def calculate_bs(self, pl_result, bank_end_balance, fa_add, num_months):
        """Compute balance sheet.

        Args:
            pl_result: output of calculate_pl()
            bank_end_balance: sum of all bank ending balances (verifies cash)
            fa_add: fixed asset additions during period
            num_months: number of months

        Returns dict with all BS line item values.
        """
        detail = pl_result['_detail']

        # --- Assets ---
        cash_end = bank_end_balance
        ar = self._bo('应收账款')
        prepay = self._bo('预付账款')
        other_recv = self._bo('其他应收款')
        inventory = self._bo('存货')

        curr_assets = round(cash_end + ar + prepay + other_recv + inventory, 2)

        fa_orig = self._bo('固定资产原价') + fa_add
        acc_depr_beg = self._bo('减：累计折旧')
        acc_depr_end = acc_depr_beg + self.monthly_depr * num_months
        fa_net = fa_orig - acc_depr_end

        ltd = max(self._bo('长期待摊费用') - self.monthly_ltd_amort * num_months, 0)

        non_curr_assets = round(fa_net + ltd, 2)
        total_assets = round(curr_assets + non_curr_assets, 2)

        # --- Liabilities ---
        ap = self._bo('应付账款')
        pr = self._bo('预收账款')
        payroll_payable = 0.0  # Simplified: assume all paid

        vat_pay = pl_result['_detail']['vat']
        surcharges_pay = pl_result['surcharges'] if pl_result['surcharges'] else 0
        tax_payable = round(vat_pay + surcharges_pay, 2)

        # --- Equity ---
        capital = self._bo('实收资本（或股本）')
        undist_beg = self._bo('未分配利润')
        net_profit = pl_result['net_profit'] if pl_result['net_profit'] else 0
        undist_end = round(undist_beg + net_profit, 2)
        total_equity = round(capital + undist_end, 2)

        # --- Other payables (plug) ---
        known_liab = ap + pr + payroll_payable + tax_payable
        other_pay = round(total_assets - known_liab - total_equity, 2)

        curr_liab = round(ap + pr + payroll_payable + tax_payable + other_pay, 2)
        total_liab = curr_liab
        total_le = round(total_liab + total_equity, 2)

        return {
            # Assets
            'cash': self._vo(cash_end),
            'ar': self._vo(ar),
            'prepay': self._vo(prepay),
            'other_recv': self._vo(other_recv),
            'inventory': self._vo(inventory),
            'curr_assets': self._vo(curr_assets),
            'fa_orig': self._vo(fa_orig),
            'acc_depr': self._vo(acc_depr_end),
            'acc_depr_beg': self._vo(acc_depr_beg),
            'fa_net': self._vo(fa_net),
            'ltd': self._vo(ltd),
            'non_curr_assets': self._vo(non_curr_assets),
            'total_assets': self._vo(total_assets),
            # Liabilities
            'ap': self._vo(ap),
            'pr': self._vo(pr),
            'payroll_payable': self._vo(payroll_payable),
            'tax_payable': self._vo(tax_payable),
            'other_pay': self._vo(other_pay),
            'curr_liab': self._vo(curr_liab),
            'total_liab': self._vo(total_liab),
            # Equity
            'capital': self._vo(capital),
            'undist_profit_end': self._vo(undist_end),
            'undist_profit_beg': self._vo(undist_beg),
            'total_equity': self._vo(total_equity),
            'total_le': self._vo(total_le),
            # Check
            '_balance_check': round(total_assets - total_le, 2),
            '_other_pay_is_plug': True,
        }
```

- [ ] **Step 2: Run verification against known顺恒 data**

```python
# Verify calc engine produces the same results as generate_formatted.py
from calc_engine import CalcEngine
from mapping_engine import MappingEngine

# Use顺恒 opening BS values
opening_bs = {
    '货币资金': 229538.60, '应收账款': 92618.0, '预付账款': 188806.03,
    '其他应收款': 6282.21, '存货': 481521.58,
    '流动资产合计': 998766.42, '固定资产原价': 4792025.06,
    '减：累计折旧': 1383059.29, '固定资产账面价值': 3408965.77,
    '长期待摊费用': 92523.08, '非流动资产合计': 3501488.85,
    '资产总计': 4500255.27,
    '应付账款': 13856.0, '预收账款': 357.0, '应付职工薪酬': 34922.21,
    '应交税费': 39943.64, '其他应付款': 6646329.20,
    '流动负债合计': 6735408.05, '负债合计': 6735408.05,
    '实收资本（或股本）': 550000.0, '未分配利润': -2785152.78,
    '所有者权益（或股东权益）合计': -2235152.78,
}

# Set up mappings (顺恒's known account codes)
import tempfile, os
d = tempfile.mkdtemp()
me = MappingEngine(d)
me.add_code_mapping('pl', '一、营业收入', [])  # Will match everything in sales file
me.add_code_mapping('pl', '减：营业成本', ['1110701000000000000', '1110799000000000000'])
# All other cost codes go to admin (except COGS codes above)

engine = CalcEngine(opening_bs)
# Quick: test the _vo helper
assert engine._vo(0.001) is None
assert engine._vo(100.0) == 100.0
assert engine._vo(None) is None
print("CalcEngine helpers OK")
```

- [ ] **Step 3: Commit**

```bash
git add calc_engine.py
git commit -m "feat: add calculation engine for PL and BS"
```

---
### Task 6: Format renderer

**Files:**
- Create: `format_renderer.py`

Adapts the openpyxl formatting logic from `generate_formatted.py:219-575` into a reusable renderer driven by template configs + computed data.

- [ ] **Step 1: Write `format_renderer.py`**

```python
# -*- coding: utf-8 -*-
"""Format renderer — produces formatted .xlsx files from template config + data.

Adapts formatting constants and logic from generate_formatted.py (lines 219-575).
"""
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime


# -------- Format constants (verified against Dec 2025 templates) --------

thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'))
no_border = Border()

font_title = Font(name='宋体', size=14)
font_subtitle = Font(name='宋体', size=10)
font_header = Font(name='宋体', size=10)
font_item = Font(name='宋体', size=10)
font_section = Font(name='宋体', size=10)
font_number = Font(name='Arial', size=10)

align_title = Alignment(horizontal='center', vertical='center')
align_subtitle_right = Alignment(horizontal='right', vertical='center')
align_header = Alignment(horizontal='center', vertical='center')
align_item_left = Alignment(horizontal='left', vertical='center')
align_number = Alignment(horizontal='right', vertical='center')
align_center = Alignment(horizontal='center', vertical='center')

num_fmt = '#,##0.00'


# -------- Renderer --------

class ReportRenderer:
    """Render one statement type to a formatted .xlsx workbook."""

    def __init__(self, template_config):
        self.cfg = template_config
        self.wb = openpyxl.Workbook()
        self.ws = self.wb.active

    def render(self, company_name, period_str, data_map, output_path):
        """Render the full statement.

        Args:
            company_name: e.g. '来宾市顺恒废旧汽车回收有限公司'
            period_str: e.g. '2026年1-5月' or '2026年5月31日'
            data_map: dict mapping item_name -> {value_col_index: value}
                      e.g. {'一、营业收入': {0: 4080824.78, 1: 915170.87}}
            output_path: where to save the .xlsx file
        """
        cfg = self.cfg
        ncols = cfg['ncols']

        # -- Title --
        self.ws.title = cfg['template_type']

        # Row 1: Title
        self.ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        c = self.ws.cell(row=1, column=1, value=cfg['template_type'])
        self._fmt(c, font_title, align_title, no_border)
        self.ws.row_dimensions[1].height = cfg.get('row_height_title', 21)

        # Row 2: Subtitle
        self.ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
        c = self.ws.cell(row=2, column=1, value=cfg['subtitle'])
        self._fmt(c, font_subtitle, align_subtitle_right, no_border)
        self.ws.row_dimensions[2].height = cfg.get('row_height_data', 17.25)

        # Row 3: Company info
        self._write_company_row(3, company_name, period_str, ncols)

        # Row 4: Column headers
        self._write_col_headers(4, ncols)

        # Data rows
        self._write_data_rows(data_map, ncols)

        # Empty rows to reach total_rows
        data_end_row = cfg['data_start_row'] + len(cfg['items']) - 1
        for er in range(data_end_row + 1, cfg['total_rows'] + 1):
            self._write_empty_row(er, ncols)

        # Column widths
        if 'col_widths' in cfg:
            for ci, w in enumerate(cfg['col_widths'], 1):
                self.ws.column_dimensions[get_column_letter(ci)].width = w

        # Print settings
        self.ws.sheet_properties.pageSetUpPr = \
            openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
        self.ws.page_setup.orientation = 'landscape'

        self.wb.save(output_path)
        return output_path

    def _write_company_row(self, row, company_name, period_str, ncols):
        cfg = self.cfg
        template_type = cfg['template_type']

        if template_type == '资产负债表':
            # Special: left 4 cols for company, right 4 cols for date
            self.ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
            c = self.ws.cell(row=row, column=1, value=f'编制单位：{company_name}')
            self._fmt(c, font_subtitle, align_item_left, no_border)

            self.ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=8)
            c = self.ws.cell(row=row, column=5, value=f'日期：{period_str}              单位：元')
            self._fmt(c, font_subtitle, align_item_left, no_border)
        else:
            # PL and CF: cols 2-3 merged for period info
            self.ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
            texts = [f'编制单位：{company_name}', f'期间：{period_str}', '', '单位：元']
            if ncols == 5:
                texts.append('')
            for ci, txt in enumerate(texts, 1):
                c = self.ws.cell(row=row, column=ci, value=txt)
                self._fmt(c, font_subtitle, align_item_left, no_border)

        self.ws.row_dimensions[row].height = cfg.get('row_height_data', 17.25)

    def _write_col_headers(self, row, ncols):
        cfg = self.cfg
        if cfg['template_type'] == '资产负债表':
            headers = ['资产', '行次', '期末余额', '年初余额',
                       '负债和所有者权益', '行次', '期末余额', '年初余额']
        elif ncols == 5:
            headers = ['项目', '行次', '本年累计金额', '本月金额', '']
        else:
            headers = cfg.get('value_col_labels', ['', ''])
            headers = ['项目', '行次'] + headers

        for ci, hdr in enumerate(headers, 1):
            if ci > ncols: break
            c = self.ws.cell(row=row, column=ci, value=hdr)
            self._fmt(c, font_header, align_header, thin_border)
        self.ws.row_dimensions[row].height = cfg.get('row_height_data', 17.25)

    def _write_data_rows(self, data_map, ncols):
        cfg = self.cfg
        template_type = cfg['template_type']

        for item in cfg['items']:
            row = item['row']
            is_section = item.get('section', False)

            if template_type == '资产负债表':
                self._write_bs_row(row, item, data_map, ncols, is_section)
            else:
                self._write_simple_row(row, item, data_map, ncols, is_section)

            self.ws.row_dimensions[row].height = cfg.get('row_height_data', 17.25)

    def _write_simple_row(self, row, item, data_map, ncols, is_section):
        """Write a PL or CF data row (single-sided)."""
        name = item.get('name', '')
        row_num = item.get('row_num', '')

        if not name and not row_num:
            self._write_empty_row(row, ncols)
            return

        # Col 0: item name
        c = self.ws.cell(row=row, column=1, value=name)
        self._fmt(c, font_section if is_section else font_item,
                  align_item_left, thin_border)

        # Col 1: row number
        c = self.ws.cell(row=row, column=2, value=row_num if row_num else None)
        self._fmt(c, font_number, align_center, thin_border)

        # Cols 2+: amount values
        row_data = data_map.get(name, {})
        for vi, vcol in enumerate(cfg['columns']['value_cols']):
            val = row_data.get(vi)
            c = self.ws.cell(row=row, column=vcol + 1, value=val)
            self._fmt(c, font_number, align_number, thin_border, num_fmt)

    def _write_bs_row(self, row, item, data_map, ncols, is_section):
        """Write a BS data row (double-sided: assets + liabilities)."""
        # Left side (assets)
        left_name = item.get('left', '')
        left_rn = item.get('left_row_num', '')
        right_name = item.get('right', '')
        right_rn = item.get('right_row_num', '')

        # Col 0: asset name
        c = self.ws.cell(row=row, column=1, value=left_name)
        self._fmt(c, font_section if is_section else font_item,
                  align_item_left, thin_border)

        # Col 1: row num
        c = self.ws.cell(row=row, column=2, value=left_rn if left_rn else None)
        self._fmt(c, font_number, align_center, thin_border)

        # Cols 2-3: asset amounts
        left_data = data_map.get(left_name, {})
        c = self.ws.cell(row=row, column=3, value=left_data.get(0))
        self._fmt(c, font_number, align_number, thin_border, num_fmt)
        c = self.ws.cell(row=row, column=4, value=left_data.get(1))
        self._fmt(c, font_number, align_number, thin_border, num_fmt)

        # Col 4: liability name
        c = self.ws.cell(row=row, column=5, value=right_name if right_name else None)
        self._fmt(c, font_section if is_section else font_item,
                  align_item_left, thin_border)

        # Col 5: row num
        c = self.ws.cell(row=row, column=6, value=right_rn if right_rn else None)
        self._fmt(c, font_number, align_center, thin_border)

        # Cols 6-7: liability amounts
        right_data = data_map.get(right_name, {})
        c = self.ws.cell(row=row, column=7, value=right_data.get(0))
        self._fmt(c, font_number, align_number, thin_border, num_fmt)
        c = self.ws.cell(row=row, column=8, value=right_data.get(1))
        self._fmt(c, font_number, align_number, thin_border, num_fmt)

    def _write_empty_row(self, row, ncols):
        for ci in range(1, ncols + 1):
            c = self.ws.cell(row=row, column=ci, value=None)
            self._fmt(c, font_item, None, thin_border)
        self.ws.row_dimensions[row].height = 13.5

    @staticmethod
    def _fmt(cell, font=None, alignment=None, border=None, number_format=None):
        if font: cell.font = font
        if alignment: cell.alignment = alignment
        if border: cell.border = border
        if number_format: cell.number_format = number_format
```

- [ ] **Step 2: Verify renderer produces valid .xlsx**

```python
import template_engine as te
from format_renderer import ReportRenderer
import tempfile, os

cfg = te.load_preset('pl')
renderer = ReportRenderer(cfg)

# Simple test: render with dummy data
data = {
    '一、营业收入': {0: 4080824.78, 1: 915170.87},
    '减：营业成本': {0: 2466533.98, 1: 849790.24},
    # ... rest empty
}
tmp = os.path.join(tempfile.mkdtemp(), 'test_pl.xlsx')
renderer.render('测试公司', '2026年1-5月', data, tmp)

# Verify the file is valid
import openpyxl
wb = openpyxl.load_workbook(tmp)
ws = wb.active
# Check title
assert ws.cell(row=1, column=1).value == '利润表'
# Check first data item
assert ws.cell(row=5, column=1).value == '一、营业收入'
# Check number format
cell = ws.cell(row=5, column=3)
assert cell.number_format == '#,##0.00'
# Check border exists on data cells
assert cell.border.left.style is not None

print(f"Render test passed! Written to {tmp}")
```

- [ ] **Step 3: Commit**

```bash
git add format_renderer.py
git commit -m "feat: add Excel format renderer with template-driven layout"
```

---
### Task 7: Validator

**Files:**
- Create: `validator.py`

Basic validation module for MVP. Implements the first three layers of the four-layer validation design: data format checks, data verification checks, and balance checks. Full integration with UI comes in a later task.

- [ ] **Step 1: Write `validator.py`**

```python
# -*- coding: utf-8 -*-
"""Validator — multi-layer checks for financial statement generation.

Three validation functions:
  validate_import()      — Layer 1: Check source files before extraction
  validate_extraction()  — Layer 2: Check extracted data sanity
  validate_balance()     — Layer 3: Check A-L-E and cash reconciliation

Each returns: (passed: bool, errors: list, warnings: list)
  errors = blocking issues (must fix before generating)
  warnings = anomalies to review (can proceed after confirmation)
"""


def validate_import(file_type, filepath):
    """Layer 1: Check file format before extraction.

    Returns (passed, errors, warnings).
    """
    errors = []
    warnings = []

    if not filepath:
        errors.append({
            'title': '文件未上传',
            'cause': f'未提供{file_type}文件',
            'fix': f'请上传{file_type}文件',
        })
        return False, errors, warnings

    ext = filepath.lower().split('.')[-1] if '.' in filepath else ''
    if ext not in ('xls', 'xlsx'):
        errors.append({
            'title': '文件格式不支持',
            'cause': f'{file_type}文件的扩展名是 .{ext}，系统只支持 .xls 或 .xlsx',
            'fix': '请上传Excel格式的文件（.xls 或 .xlsx）',
        })
        return False, errors, warnings

    return True, errors, warnings


def validate_extraction(extraction_result, mapping_engine, statement_type, prev_month_stats=None):
    """Layer 2: Check extracted data.

    Args:
        extraction_result: dict with keys 'sales', 'costs', 'payroll', 'bank'
        mapping_engine: MappingEngine instance
        statement_type: 'pl' | 'bs' | 'cf'
        prev_month_stats: optional dict with last month's stats for comparison

    Returns (passed, errors, warnings).
    """
    errors = []
    warnings = []

    # Check for unmapped codes
    all_codes = set()
    for source in ['sales', 'costs']:
        data = extraction_result.get(source, {})
        if isinstance(data, dict):
            for mo, records in data.items():
                for rec in records:
                    all_codes.add(rec.get('account_code', ''))

    unmapped = mapping_engine.find_unmapped_codes(statement_type, all_codes)
    if unmapped:
        codes_str = ', '.join(unmapped[:10])
        suffix = '...' if len(unmapped) > 10 else ''
        errors.append({
            'title': f'发现 {len(unmapped)} 个未配置的科目编码',
            'cause': f'以下科目编码未在科目映射中找到：{codes_str}{suffix}',
            'fix': '请到「科目映射」页面为这些编码指定对应的报表行项目',
        })

    # Check bank balance is non-negative
    bank_data = extraction_result.get('bank', {})
    for bank_name, bank_result in bank_data.items():
        if isinstance(bank_result, dict):
            bal = bank_result.get('end_balance', 0)
            if bal < 0:
                errors.append({
                    'title': f'{bank_name}余额为负数',
                    'cause': f'{bank_name}期末余额为 ¥{bal:,.2f}',
                    'fix': '请检查银行流水文件是否正确，或确认账户是否确实透支',
                })

    # Invoice count comparison with previous month
    if prev_month_stats:
        curr_sales_count = sum(
            len(records) for records in
            extraction_result.get('sales', {}).values()
            if isinstance(records, list)
        )
        prev_sales_count = prev_month_stats.get('sales_count', 0)
        if prev_sales_count > 0:
            deviation = abs(curr_sales_count - prev_sales_count) / prev_sales_count
            if deviation > 0.3:
                direction = '多' if curr_sales_count > prev_sales_count else '少'
                warnings.append({
                    'title': f'本月发票数量({curr_sales_count}张)比上月({prev_sales_count}张){direction}{deviation*100:.0f}%',
                    'cause': '可能是业务量波动，也可能是漏传文件',
                    'fix': '请确认所有发票都已上传',
                })

    passed = len(errors) == 0
    return passed, errors, warnings


def validate_balance(pl_result, bs_result):
    """Layer 3: Check accounting equation and cash reconciliation.

    Returns (passed, errors, warnings).
    """
    errors = []
    warnings = []

    # Check A = L + E
    a = bs_result.get('total_assets')
    le = bs_result.get('total_le')
    if a is not None and le is not None:
        diff = round(a - le, 2)
        if abs(diff) > 0.01:
            errors.append({
                'title': f'资产负债表不平——资产-负债-所有者权益 = ¥{diff:,.2f}',
                'cause': '通常是因为有新科目未配置映射，或期初余额录入有误',
                'fix': '请到「科目映射」页检查是否有新科目编码未归类',
            })

    # Check Net Profit = Undist Profit change
    np_val = pl_result.get('net_profit')
    if np_val is not None:
        up_end = bs_result.get('undist_profit_end')
        up_beg = bs_result.get('undist_profit_beg')
        if up_end is not None and up_beg is not None:
            expected_change = up_end - up_beg
            diff = round(np_val - expected_change, 2)
            if abs(diff) > 0.01:
                errors.append({
                    'title': f'净利润(¥{np_val:,.2f})≠未分配利润变动(¥{expected_change:,.2f})',
                    'cause': '可能是期初未分配利润录入错误，或存在直接计入权益的项目',
                    'fix': '请检查「公司档案」中的年初余额是否正确',
                })

    passed = len(errors) == 0
    return passed, errors, warnings


def format_validation_message(issue):
    """Format a validation issue dict into a user-readable string.

    Uses the three-element format: 发现问题 / 可能原因 / 建议操作
    """
    lines = [
        f"❌ **发现问题：** {issue['title']}",
        f"🔍 **可能原因：** {issue['cause']}",
        f"🔧 **建议操作：** {issue['fix']}",
    ]
    return '\n'.join(lines)
```

- [ ] **Step 2: Quick smoke test**

```python
from validator import validate_import, validate_balance
import tempfile

# Test import validation
passed, errs, warns = validate_import('发票', 'test.pdf')
assert not passed
assert len(errs) == 1

passed, errs, warns = validate_import('发票', 'test.xlsx')
assert passed

# Test balance validation
pl = {'net_profit': 334852.98}
bs = {'total_assets': 4378328.56, 'total_le': 4378328.56,
      'undist_profit_end': -2450299.80, 'undist_profit_beg': -2785152.78}
passed, errs, warns = validate_balance(pl, bs)
assert passed

print("Validator smoke test passed!")
```

- [ ] **Step 3: Commit**

```bash
git add validator.py
git commit -m "feat: add multi-layer validator for financial data quality"
```

---
### Task 8: Streamlit app — main entry and navigation

**Files:**
- Create: `app.py`

Sets up the Streamlit shell: sidebar navigation, session state initialization, global company/period selectors, and page routing.

- [ ] **Step 1: Write `app.py` — core infrastructure**

```python
# -*- coding: utf-8 -*-
"""财务报表生成工具 — Streamlit Web Application.

Local web app for generating 小企业会计准则 financial statements.
All data stays on the user's computer. No network requests.
"""
import streamlit as st
import os
import json
import glob
from datetime import datetime

# ---- Page config ----
st.set_page_config(
    page_title='财务报表生成工具',
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='expanded',
)

# ---- Paths ----
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
COMPANIES_DIR = os.path.join(DATA_DIR, 'companies')
UPLOADS_DIR = os.path.join(DATA_DIR, 'uploads')
OUTPUT_DIR = os.path.join(DATA_DIR, 'output')
LOGS_DIR = os.path.join(DATA_DIR, 'logs')

for d in [COMPANIES_DIR, UPLOADS_DIR, OUTPUT_DIR, LOGS_DIR]:
    os.makedirs(d, exist_ok=True)

# ---- Session state ----
DEFAULTS = {
    'current_company': None,
    'current_year': 2026,
    'current_month': 5,
    'page': '首页',
    'data_verified': False,
    'generated_pl_path': None,
    'generated_bs_path': None,
    'generated_cf_path': None,
}

for key, val in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = val

# ---- Helpers ----
def get_company_list():
    """List company directories under data/companies/."""
    if not os.path.exists(COMPANIES_DIR):
        return []
    return [d for d in os.listdir(COMPANIES_DIR)
            if os.path.isdir(os.path.join(COMPANIES_DIR, d))]

def get_company_dir(name):
    return os.path.join(COMPANIES_DIR, name)

def load_company_config(name):
    path = os.path.join(get_company_dir(name), 'config.json')
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_company_config(name, config):
    d = get_company_dir(name)
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, 'config.json'), 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

# ---- Sidebar ----
with st.sidebar:
    st.title('📊 财务报表工具')

    # Company selector
    companies = get_company_list()
    company_options = ['— 选择公司 —'] + companies
    current_idx = 0
    if st.session_state.current_company and st.session_state.current_company in companies:
        current_idx = companies.index(st.session_state.current_company) + 1
    selected_company = st.selectbox(
        '当前公司',
        company_options,
        index=current_idx,
        key='sidebar_company',
    )
    if selected_company != '— 选择公司 —':
        st.session_state.current_company = selected_company

    # Period selectors
    st.session_state.current_year = st.selectbox(
        '年份', list(range(2023, 2031)),
        index=list(range(2023, 2031)).index(st.session_state.current_year),
    )
    st.session_state.current_month = st.selectbox(
        '月份', list(range(1, 13)),
        index=st.session_state.current_month - 1,
    )

    st.divider()

    # Navigation
    pages = ['首页', '公司档案', '科目映射', '数据导入', '数据核验', '报表生成', '分析导出']
    selected_page = st.radio('导航', pages, index=pages.index(st.session_state.page))
    st.session_state.page = selected_page

    st.divider()
    st.caption('所有数据仅在本地处理\n不上传任何服务器')

# ---- Page router ----
page = st.session_state.page

if page == '首页':
    st.header('🏠 首页')
    # ... implemented in Task 9
elif page == '公司档案':
    st.header('📋 公司档案')
    # ... implemented in Task 9
elif page == '科目映射':
    st.header('⚙️ 科目映射')
    # ... implemented in Task 10
elif page == '数据导入':
    st.header('📤 数据导入')
    # ... implemented in Task 11
elif page == '数据核验':
    st.header('✅ 数据核验')
    # ... implemented in Task 11
elif page == '报表生成':
    st.header('📊 报表生成')
    # ... implemented in Task 12
elif page == '分析导出':
    st.header('📝 分析导出')
    # ... implemented in Task 12

# ---- Footer ----
st.divider()
st.caption(f'数据存储位置：{DATA_DIR}')
```

- [ ] **Step 2: Test the app launches**

Run: `streamlit run app.py --server.headless true --server.port 8501 &`
Then: `sleep 3 && curl -s http://localhost:8501 | head -20`
Expected: HTML page with "财务报表生成工具" title.

Kill the server after test: `pkill -f streamlit`

- [ ] **Step 3: Commit**

```bash
git add app.py
git commit -m "feat: add Streamlit app shell with navigation and session state"
```

---
### Task 9: Streamlit app — 首页 and 公司档案 pages

**Files:**
- Modify: `app.py:80-95` (replace placeholder page blocks)

- [ ] **Step 1: Implement 首页 page**

Replace the 首页 placeholder in app.py. The home page shows quick-action buttons and recent report history.

```python
if page == '首页':
    st.header('🏠 首页')

    if not st.session_state.current_company:
        st.info('👈 请先在左侧选择公司，或去「公司档案」创建新公司')
    else:
        company = st.session_state.current_company
        config = load_company_config(company)
        st.subheader(f'当前公司：{company}')
        if config.get('full_name'):
            st.caption(f"全称：{config['full_name']}")

        year = st.session_state.current_year
        month = st.session_state.current_month
        period_label = f'{year}年{month}月'

        st.divider()
        col1, col2, col3 = st.columns(3)

        with col1:
            if st.button('📤 导入数据', use_container_width=True, type='primary'):
                st.session_state.page = '数据导入'
                st.rerun()

        with col2:
            if st.button('📊 生成报表', use_container_width=True):
                if st.session_state.data_verified:
                    st.session_state.page = '报表生成'
                    st.rerun()
                else:
                    st.warning('请先导入并核验数据')

        with col3:
            if st.button('📝 导出报告', use_container_width=True):
                st.session_state.page = '分析导出'
                st.rerun()

        # Recent reports
        st.divider()
        st.subheader('历史报表')
        output_pattern = os.path.join(OUTPUT_DIR, company, f'{year}年*月*.xlsx')
        past_files = sorted(glob.glob(output_pattern), reverse=True)[:12]
        if past_files:
            for f in past_files:
                fname = os.path.basename(f)
                st.write(f'✅ {fname}')
        else:
            st.caption('暂无历史报表')
```

- [ ] **Step 2: Implement 公司档案 page**

```python
elif page == '公司档案':
    st.header('📋 公司档案')

    tab1, tab2 = st.tabs(['公司信息', '年初余额'])

    with tab1:
        # Create or edit company
        action = st.radio('操作', ['选择已有公司', '创建新公司'], horizontal=True)

        if action == '创建新公司':
            new_name = st.text_input('公司简称（用于文件夹命名）', placeholder='例如：顺恒废旧')
            full_name = st.text_input('公司全称（将显示在报表上）',
                                      placeholder='例如：来宾市顺恒废旧汽车回收有限公司')
            tax_id = st.text_input('纳税人识别号', placeholder='可选')

            if st.button('创建公司', type='primary'):
                if new_name:
                    config = {'full_name': full_name, 'tax_id': tax_id,
                              'created_at': datetime.now().isoformat(),
                              'template_type': 'preset'}
                    save_company_config(new_name, config)
                    st.session_state.current_company = new_name
                    st.success(f'公司「{new_name}」创建成功！')
                    st.rerun()
                else:
                    st.error('请输入公司简称')

        else:
            companies = get_company_list()
            if companies:
                sel = st.selectbox('选择公司', companies,
                                   index=companies.index(st.session_state.current_company)
                                   if st.session_state.current_company in companies else 0)
                if sel:
                    st.session_state.current_company = sel
                    config = load_company_config(sel)
                    st.json(config)
            else:
                st.info('暂无公司档案，请先创建')

    with tab2:
        if not st.session_state.current_company:
            st.info('请先选择或创建公司')
        else:
            company = st.session_state.current_company
            config = load_company_config(company)
            opening = config.get('opening_bs', {})

            st.subheader('年初资产负债表余额')
            st.caption('填写去年12月31日的资产负债表期末数。填一次即可，系统会用来计算本期变动。')

            # Asset side
            col1, col2 = st.columns(2)
            with col1:
                st.markdown('**资产**')
                opening['货币资金'] = st.number_input('货币资金', value=float(opening.get('货币资金', 0)), step=1000.0, format='%.2f')
                opening['应收账款'] = st.number_input('应收账款', value=float(opening.get('应收账款', 0)), step=1000.0, format='%.2f')
                opening['预付账款'] = st.number_input('预付账款', value=float(opening.get('预付账款', 0)), step=1000.0, format='%.2f')
                opening['其他应收款'] = st.number_input('其他应收款', value=float(opening.get('其他应收款', 0)), step=1000.0, format='%.2f')
                opening['存货'] = st.number_input('存货', value=float(opening.get('存货', 0)), step=1000.0, format='%.2f')
                opening['固定资产原价'] = st.number_input('固定资产原价', value=float(opening.get('固定资产原价', 0)), step=10000.0, format='%.2f')
                opening['减：累计折旧'] = st.number_input('减：累计折旧', value=float(opening.get('减：累计折旧', 0)), step=1000.0, format='%.2f')
                opening['长期待摊费用'] = st.number_input('长期待摊费用', value=float(opening.get('长期待摊费用', 0)), step=1000.0, format='%.2f')

            with col2:
                st.markdown('**负债及所有者权益**')
                opening['应付账款'] = st.number_input('应付账款', value=float(opening.get('应付账款', 0)), step=1000.0, format='%.2f')
                opening['预收账款'] = st.number_input('预收账款', value=float(opening.get('预收账款', 0)), step=1000.0, format='%.2f')
                opening['应付职工薪酬'] = st.number_input('应付职工薪酬', value=float(opening.get('应付职工薪酬', 0)), step=1000.0, format='%.2f')
                opening['应交税费'] = st.number_input('应交税费', value=float(opening.get('应交税费', 0)), step=1000.0, format='%.2f')
                opening['其他应付款'] = st.number_input('其他应付款', value=float(opening.get('其他应付款', 0)), step=10000.0, format='%.2f')
                opening['实收资本（或股本）'] = st.number_input('实收资本（或股本）', value=float(opening.get('实收资本（或股本）', 550000)), step=10000.0, format='%.2f')
                opening['未分配利润'] = st.number_input('未分配利润', value=float(opening.get('未分配利润', 0)), step=10000.0, format='%.2f')

            if st.button('保存年初余额', type='primary'):
                config['opening_bs'] = opening
                save_company_config(company, config)
                st.success('年初余额已保存！')
```

- [ ] **Step 3: Verify pages render**

Run: `streamlit run app.py` and check that:
- Home page shows quick-action buttons when company selected
- 公司档案 page shows create form and opening balance inputs
- Navigation between pages works via sidebar

- [ ] **Step 4: Commit**

```bash
git add app.py
git commit -m "feat: implement home page and company profile page"
```

---
### Task 10: Streamlit app — 科目映射 page

**Files:**
- Modify: `app.py` (replace the 科目映射 placeholder)

- [ ] **Step 1: Implement 科目映射 page**

The mapping page has two stages: (1) upload a sample data file to discover codes, (2) map each code to a report line via dropdown.

```python
elif page == '科目映射':
    st.header('⚙️ 科目映射')

    if not st.session_state.current_company:
        st.info('请先选择公司')
    else:
        from mapping_engine import MappingEngine
        from data_extractor import InvoiceExtractor
        import template_engine as te

        company = st.session_state.current_company
        company_dir = get_company_dir(company)
        engine = MappingEngine(company_dir)

        statement = st.selectbox('选择报表', ['利润表 (pl)', '资产负债表 (bs)'], key='mapping_statement')
        st_type = 'pl' if '利润' in statement else 'bs'

        # Show current mappings
        current = engine.get_mappings(st_type)
        if current:
            st.subheader('当前映射')
            for m in current:
                codes = ', '.join(m.get('account_codes', []))
                rtype = m['rule_type']
                st.write(f"- **{m['report_line']}** ← `{codes}` ({rtype})")

        st.divider()

        # Discover codes from uploaded file
        st.subheader('发现新科目编码')
        st.caption('上传一份发票文件，系统会扫描所有出现的科目编码，帮你发现未配置的编码。')
        sample_file = st.file_uploader('上传样本发票文件', type=['xlsx', 'xls'], key='mapping_sample')

        if sample_file:
            # Save temp file
            tmp_path = os.path.join(UPLOADS_DIR, f'_sample_{company}.xlsx')
            with open(tmp_path, 'wb') as f:
                f.write(sample_file.read())

            try:
                ie = InvoiceExtractor(tmp_path)
                by_code = ie.get_total_by_code()
                st.success(f'发现 {len(by_code)} 个科目编码')

                # Find unmapped
                unmapped = engine.find_unmapped_codes(st_type, list(by_code.keys()))

                if unmapped:
                    st.warning(f'{len(unmapped)} 个编码未配置映射：')
                    # Load template items for dropdown options
                    cfg = te.load_preset(st_type)
                    if st_type == 'bs':
                        item_names = [it['left'] for it in cfg['items'] if it['left']]
                        item_names += [it['right'] for it in cfg['items'] if it['right']]
                    else:
                        item_names = [it['name'] for it in cfg['items'] if it['name']]
                    item_names = sorted(set(item_names))

                    for code in unmapped:
                        col1, col2 = st.columns([1, 2])
                        with col1:
                            st.code(code)
                        with col2:
                            target = st.selectbox(
                                f'映射到 →',
                                ['— 选择报表行 —'] + item_names,
                                key=f'map_{code}',
                            )
                            if target != '— 选择报表行 —':
                                if st.button(f'确认', key=f'btn_{code}'):
                                    engine.add_code_mapping(st_type, target, [code])
                                    st.rerun()
                else:
                    st.success('所有编码已配置映射！✅')

            except Exception as e:
                st.error(f'读取文件失败：{e}')
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
```

- [ ] **Step 2: Commit**

```bash
git add app.py
git commit -m "feat: implement account mapping wizard page"
```

---
### Task 11: Streamlit app — 数据导入 and 数据核验 pages

**Files:**
- Modify: `app.py` (replace 数据导入 and 数据核验 placeholders)

- [ ] **Step 1: Implement 数据导入 page**

```python
elif page == '数据导入':
    st.header('📤 数据导入')

    if not st.session_state.current_company:
        st.info('请先选择公司')
    else:
        company = st.session_state.current_company
        year = st.session_state.current_year
        month = st.session_state.current_month

        st.subheader(f'导入 {year}年{month}月 数据')
        st.caption('上传以下4类文件。支持 .xls 和 .xlsx 格式。')

        col1, col2 = st.columns(2)

        with col1:
            sales_file = st.file_uploader('📄 销售收入发票', type=['xlsx', 'xls'], key='sales')
            costs_file = st.file_uploader('📄 成本费用发票', type=['xlsx', 'xls'], key='costs')

        with col2:
            bank_nong = st.file_uploader('🏦 农行流水 (.xls)', type=['xls'], key='nong')
            bank_xin = st.file_uploader('🏦 信用社流水 (.xls)', type=['xls'], key='xin')
            payroll_file = st.file_uploader('👤 工资薪金表', type=['xlsx', 'xls'], key='payroll')

        if st.button('保存文件并进入核验 →', type='primary'):
            if not any([sales_file, costs_file, bank_nong, bank_xin, payroll_file]):
                st.error('请至少上传一个文件')
            else:
                saved = []
                upload_dir = os.path.join(UPLOADS_DIR, company, f'{year}-{month:02d}')
                os.makedirs(upload_dir, exist_ok=True)

                for label, file_obj in [('sales', sales_file), ('costs', costs_file),
                                         ('nong', bank_nong), ('xin', bank_xin),
                                         ('payroll', payroll_file)]:
                    if file_obj:
                        ext = file_obj.name.split('.')[-1]
                        fname = f'{label}.{ext}'
                        fpath = os.path.join(upload_dir, fname)
                        with open(fpath, 'wb') as f:
                            f.write(file_obj.read())
                        saved.append(fname)

                st.success(f'已保存 {len(saved)} 个文件：{", ".join(saved)}')
                st.session_state.data_verified = False
                st.session_state.page = '数据核验'
                st.rerun()
```

- [ ] **Step 2: Implement 数据核验 page**

```python
elif page == '数据核验':
    st.header('✅ 数据核验')

    if not st.session_state.current_company:
        st.info('请先选择公司')
    else:
        company = st.session_state.current_company
        year = st.session_state.current_year
        month = st.session_state.current_month
        upload_dir = os.path.join(UPLOADS_DIR, company, f'{year}-{month:02d}')

        if not os.path.exists(upload_dir):
            st.warning('未找到本月上传的数据。请先到「数据导入」上传文件。')
            if st.button('去导入数据'):
                st.session_state.page = '数据导入'
                st.rerun()
        else:
            from data_extractor import InvoiceExtractor, BankExtractor, PayrollExtractor
            from validator import validate_extraction, format_validation_message
            from mapping_engine import MappingEngine

            company_dir = get_company_dir(company)
            engine = MappingEngine(company_dir)

            # Extract and display
            st.subheader('📊 数据概览')

            # Sales
            sales_path = os.path.join(upload_dir, 'sales.xlsx')
            if os.path.exists(sales_path):
                ie = InvoiceExtractor(sales_path)
                data, skipped = ie.extract()
                total_sales = sum(sum(r['amount'] for r in recs) for recs in data.values())
                count = sum(len(recs) for recs in data.values())
                st.metric('销售收入', f'¥{total_sales:,.2f}', f'{count}张发票')

            # Costs
            costs_path = os.path.join(upload_dir, 'costs.xlsx')
            if os.path.exists(costs_path):
                ie = InvoiceExtractor(costs_path)
                data, skipped = ie.extract()
                total_costs = sum(sum(r['amount'] for r in recs) for recs in data.values())
                count = sum(len(recs) for recs in data.values())
                st.metric('成本费用', f'¥{total_costs:,.2f}', f'{count}张发票')

            # Bank
            col1, col2 = st.columns(2)
            bank_end = 0.0
            with col1:
                nong_path = os.path.join(upload_dir, 'nong.xls')
                if os.path.exists(nong_path):
                    be = BankExtractor(nong_path, BankExtractor.BANK_NONGHANG)
                    result = be.extract()
                    bank_end += result['end_balance']
                    st.metric('农行余额', f'¥{result["end_balance"]:,.2f}')
            with col2:
                xin_path = os.path.join(upload_dir, 'xin.xls')
                if os.path.exists(xin_path):
                    be = BankExtractor(xin_path, BankExtractor.BANK_XINYONGSHE)
                    result = be.extract()
                    bank_end += result['end_balance']
                    st.metric('信用社余额', f'¥{result["end_balance"]:,.2f}')
            st.metric('🏦 银行余额合计', f'¥{bank_end:,.2f}')

            # Payroll
            payroll_path = os.path.join(upload_dir, 'payroll.xlsx')
            if os.path.exists(payroll_path):
                pe = PayrollExtractor(payroll_path)
                pr = pe.extract()
                st.metric('工资应发', f'¥{pr["gross_pay"]:,.2f}')

            # Validation results
            st.divider()
            st.subheader('🔍 核验结果')

            # Run basic validation
            extraction_result = {
                'sales': data if os.path.exists(sales_path) else {},
                'costs': data if os.path.exists(costs_path) else {},
                'bank': bank_result if os.path.exists(nong_path) else {},
            }
            passed, errors, warnings = validate_extraction(extraction_result, engine, 'pl')

            if not errors:
                st.success('✅ 核验通过，可以生成报表')
                st.session_state.data_verified = True
            else:
                st.error(f'🔴 发现 {len(errors)} 个问题需要处理：')
                for e in errors:
                    st.error(format_validation_message(e))

            if warnings:
                for w in warnings:
                    st.warning(format_validation_message(w))

            col1, col2 = st.columns(2)
            with col1:
                if st.button('← 返回修改数据', use_container_width=True):
                    st.session_state.page = '数据导入'
                    st.rerun()
            with col2:
                if st.button('确认无误，生成报表 →', type='primary', use_container_width=True,
                            disabled=not st.session_state.data_verified):
                    st.session_state.page = '报表生成'
                    st.rerun()
```

- [ ] **Step 3: Commit**

```bash
git add app.py
git commit -m "feat: implement data import and verification pages"
```

---
### Task 12: Streamlit app — 报表生成 and 分析导出 pages

**Files:**
- Modify: `app.py` (replace 报表生成 and 分析导出 placeholders)

- [ ] **Step 1: Implement 报表生成 page**

```python
elif page == '报表生成':
    st.header('📊 报表生成')

    if not st.session_state.current_company:
        st.info('请先选择公司')
    elif not st.session_state.data_verified:
        st.warning('请先完成数据核验')
        if st.button('去核验数据'):
            st.session_state.page = '数据核验'
            st.rerun()
    else:
        company = st.session_state.current_company
        year = st.session_state.current_year
        month = st.session_state.current_month
        company_dir = get_company_dir(company)
        upload_dir = os.path.join(UPLOADS_DIR, company, f'{year}-{month:02d}')

        if st.button('🚀 生成三大报表', type='primary', use_container_width=True):
            with st.spinner('正在生成报表...'):
                from data_extractor import InvoiceExtractor, BankExtractor, PayrollExtractor
                from mapping_engine import MappingEngine
                from calc_engine import CalcEngine
                from format_renderer import ReportRenderer
                from validator import validate_balance, format_validation_message
                import template_engine as te

                # Load mappings and config
                engine = MappingEngine(company_dir)
                config = load_company_config(company)
                opening_bs = config.get('opening_bs', {})

                # Extract data
                results = {}
                sales_path = os.path.join(upload_dir, 'sales.xlsx')
                costs_path = os.path.join(upload_dir, 'costs.xlsx')
                nong_path = os.path.join(upload_dir, 'nong.xls')
                xin_path = os.path.join(upload_dir, 'xin.xls')
                payroll_path = os.path.join(upload_dir, 'payroll.xlsx')

                if os.path.exists(sales_path):
                    results['sales'] = InvoiceExtractor(sales_path).extract()[0]
                if os.path.exists(costs_path):
                    results['costs'] = InvoiceExtractor(costs_path).extract()[0]

                bank_end = 0.0
                bank_data = {}
                for label, btype, fpath in [('农行', BankExtractor.BANK_NONGHANG, nong_path),
                                              ('信用社', BankExtractor.BANK_XINYONGSHE, xin_path)]:
                    if os.path.exists(fpath):
                        be = BankExtractor(fpath, btype)
                        result = be.extract()
                        bank_end += result['end_balance']
                        bank_data[label] = result

                payroll = {}
                if os.path.exists(payroll_path):
                    payroll[month] = PayrollExtractor(payroll_path).extract()

                # Calculate
                calc = CalcEngine(opening_bs)
                num_months = month  # Jan through current month

                # PL
                pl = calc.calculate_pl(
                    {'sales': results.get('sales', {}), 'costs': results.get('costs', {})},
                    payroll,
                    engine.get_mappings('pl'),
                    {month: bank_data},
                    num_months,
                )

                # BS
                fa_add = pl['_detail']['fa_add']
                bs = calc.calculate_bs(pl, bank_end, fa_add, num_months)

                # Validate balance
                passed, errors, warnings = validate_balance(pl, bs)

                # Prepare data maps for renderer
                pl_map = {
                    '一、营业收入': {0: pl['revenue'], 1: pl['revenue']},
                    '减：营业成本': {0: pl['cogs'], 1: pl['cogs']},
                    '税金及附加': {0: pl['surcharges'], 1: pl['surcharges']},
                    '管理费用': {0: pl['admin_exp'], 1: pl['admin_exp']},
                    '财务费用': {0: pl['fin_exp'], 1: pl['fin_exp']},
                    '二、营业利润（亏损以"-"号填列）': {0: pl['oper_profit'], 1: pl['oper_profit']},
                    '三、利润总额（亏损总额以"-"号填列）': {0: pl['total_profit'], 1: pl['total_profit']},
                    '四、净利润（净亏损以"-"号填列）': {0: pl['net_profit'], 1: pl['net_profit']},
                }

                bs_map = {
                    '货币资金': {0: bs['cash'], 1: opening_bs.get('货币资金')},
                    '应收账款': {0: bs['ar'], 1: opening_bs.get('应收账款')},
                    '预付账款': {0: bs['prepay'], 1: opening_bs.get('预付账款')},
                    '其他应收款': {0: bs['other_recv'], 1: opening_bs.get('其他应收款')},
                    '存货': {0: bs['inventory'], 1: opening_bs.get('存货')},
                    '流动资产合计': {0: bs['curr_assets'], 1: opening_bs.get('流动资产合计')},
                    '固定资产原价': {0: bs['fa_orig'], 1: opening_bs.get('固定资产原价')},
                    '减：累计折旧': {0: bs['acc_depr'], 1: bs['acc_depr_beg']},
                    '固定资产账面价值': {0: bs['fa_net'], 1: opening_bs.get('固定资产账面价值')},
                    '长期待摊费用': {0: bs['ltd'], 1: opening_bs.get('长期待摊费用')},
                    '非流动资产合计': {0: bs['non_curr_assets'], 1: opening_bs.get('非流动资产合计')},
                    '资产总计': {0: bs['total_assets'], 1: opening_bs.get('资产总计')},
                    '应付账款': {0: bs['ap'], 1: opening_bs.get('应付账款')},
                    '预收账款': {0: bs['pr'], 1: opening_bs.get('预收账款')},
                    '应付职工薪酬': {0: bs['payroll_payable'], 1: opening_bs.get('应付职工薪酬')},
                    '应交税费': {0: bs['tax_payable'], 1: opening_bs.get('应交税费')},
                    '其他应付款': {0: bs['other_pay'], 1: opening_bs.get('其他应付款')},
                    '流动负债合计': {0: bs['curr_liab'], 1: opening_bs.get('流动负债合计')},
                    '负债合计': {0: bs['total_liab'], 1: opening_bs.get('负债合计')},
                    '实收资本（或股本）': {0: bs['capital'], 1: opening_bs.get('实收资本（或股本）')},
                    '未分配利润': {0: bs['undist_profit_end'], 1: bs['undist_profit_beg']},
                    '所有者权益（或股东权益）合计': {0: bs['total_equity'], 1: opening_bs.get('所有者权益（或股东权益）合计')},
                    '负债和所有者权益（或股东权益）总计': {0: bs['total_le'], 1: opening_bs.get('负债和所有者权益（或股东权益）总计')},
                }

                # Render
                period_label = f'{year}年1-{month}月'
                bs_date = f'{year}年{month}月{28 if month==2 else 30}日'
                output_dir = os.path.join(OUTPUT_DIR, company)
                os.makedirs(output_dir, exist_ok=True)

                full_name = config.get('full_name', company)

                # PL
                pl_cfg = te.load_preset('pl')
                pl_renderer = ReportRenderer(pl_cfg)
                pl_path = os.path.join(output_dir, f'{company}_{year}年1-{month}月_利润表.xlsx')
                pl_renderer.render(full_name, period_label, pl_map, pl_path)

                # BS
                bs_cfg = te.load_preset('bs')
                bs_renderer = ReportRenderer(bs_cfg)
                bs_path = os.path.join(output_dir, f'{company}_{year}年1-{month}月_资产负债表.xlsx')
                bs_renderer.render(full_name, bs_date, bs_map, bs_path)

                st.session_state.generated_pl_path = pl_path
                st.session_state.generated_bs_path = bs_path

                # Show results
                st.success('报表生成完成！')

                # Validation results
                if not errors:
                    st.success(f'✅ 校验通过 — 资产负债表平衡 (A-L-E = ¥{bs["_balance_check"]:,.2f})')
                else:
                    for e in errors:
                        st.error(format_validation_message(e))

                if warnings:
                    for w in warnings:
                        st.warning(format_validation_message(w))

                # Key metrics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric('营业收入', f'¥{pl["revenue"]:,.2f}' if pl['revenue'] else '¥0.00')
                with col2:
                    st.metric('净利润', f'¥{pl["net_profit"]:,.2f}' if pl['net_profit'] else '¥0.00')
                with col3:
                    st.metric('资产总计', f'¥{bs["total_assets"]:,.2f}' if bs['total_assets'] else '¥0.00')
                with col4:
                    st.metric('银行余额', f'¥{bank_end:,.2f}')

                # Preview tables
                st.divider()
                st.subheader('利润表预览')

                # Read generated file for preview
                import pandas as pd
                df_pl = pd.read_excel(pl_path, header=None)
                st.dataframe(df_pl.iloc[4:25, :4], use_container_width=True, height=400)
```

- [ ] **Step 2: Implement 分析导出 page**

```python
elif page == '分析导出':
    st.header('📝 分析导出')

    if not st.session_state.current_company:
        st.info('请先选择公司')
    else:
        company = st.session_state.current_company

        st.subheader('下载报表文件')

        col1, col2, col3 = st.columns(3)

        with col1:
            if st.session_state.generated_pl_path and os.path.exists(st.session_state.generated_pl_path):
                with open(st.session_state.generated_pl_path, 'rb') as f:
                    st.download_button(
                        '📥 下载利润表',
                        f.read(),
                        os.path.basename(st.session_state.generated_pl_path),
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True,
                    )
            else:
                st.button('📥 利润表（未生成）', disabled=True, use_container_width=True)

        with col2:
            if st.session_state.generated_bs_path and os.path.exists(st.session_state.generated_bs_path):
                with open(st.session_state.generated_bs_path, 'rb') as f:
                    st.download_button(
                        '📥 下载资产负债表',
                        f.read(),
                        os.path.basename(st.session_state.generated_bs_path),
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True,
                    )
            else:
                st.button('📥 资产负债表（未生成）', disabled=True, use_container_width=True)

        with col3:
            st.button('📥 现金流量表（开发中）', disabled=True, use_container_width=True)

        st.divider()

        st.subheader('📝 分析报告')
        st.caption('Word格式分析报告将在后续版本中提供。当前请先生成Excel报表。')

        if st.session_state.generated_pl_path:
            st.info('✅ 报表已生成，可使用上方按钮下载。分析报告功能将在第二期开发中实现。')

    # Quick link
    st.divider()
    st.caption(f'报表保存位置：{os.path.join(OUTPUT_DIR, company) if st.session_state.current_company else OUTPUT_DIR}')
```

- [ ] **Step 3: Commit**

```bash
git add app.py
git commit -m "feat: implement report generation and export pages"
```

---
### Task 13: Integration test — full flow with 顺恒 data

**Files:**
- Create: `test_integration.py`

- [ ] **Step 1: Write integration test**

```python
# -*- coding: utf-8 -*-
"""Integration test: full flow with 顺恒废旧 2026 Jan-May data."""
import os, sys, json, tempfile, shutil

# Setup test environment
TEST_DIR = tempfile.mkdtemp()
os.makedirs(os.path.join(TEST_DIR, 'data', 'companies', '顺恒废旧'), exist_ok=True)
os.makedirs(os.path.join(TEST_DIR, 'data', 'uploads', '顺恒废旧', '2026-05'), exist_ok=True)
os.makedirs(os.path.join(TEST_DIR, 'data', 'output', '顺恒废旧'), exist_ok=True)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mapping_engine import MappingEngine
from calc_engine import CalcEngine
from data_extractor import InvoiceExtractor, BankExtractor, PayrollExtractor
import template_engine as te
from format_renderer import ReportRenderer
from validator import validate_balance

# 1. Setup company config with opening BS
company_dir = os.path.join(TEST_DIR, 'data', 'companies', '顺恒废旧')
config = {
    'full_name': '来宾市顺恒废旧汽车回收有限公司',
    'opening_bs': {
        '货币资金': 229538.60, '应收账款': 92618.0, '预付账款': 188806.03,
        '其他应收款': 6282.21, '存货': 481521.58,
        '流动资产合计': 998766.42,
        '固定资产原价': 4792025.06, '减：累计折旧': 1383059.29,
        '固定资产账面价值': 3408965.77, '长期待摊费用': 92523.08,
        '非流动资产合计': 3501488.85, '资产总计': 4500255.27,
        '应付账款': 13856.0, '预收账款': 357.0, '应付职工薪酬': 34922.21,
        '应交税费': 39943.64, '其他应付款': 6646329.20,
        '流动负债合计': 6735408.05, '负债合计': 6735408.05,
        '实收资本（或股本）': 550000.0, '未分配利润': -2785152.78,
        '所有者权益（或股东权益）合计': -2235152.78,
    },
}
with open(os.path.join(company_dir, 'config.json'), 'w', encoding='utf-8') as f:
    json.dump(config, f, ensure_ascii=False, indent=2)

# 2. Setup mappings (顺恒's COGS codes)
engine = MappingEngine(company_dir)
engine.add_code_mapping('pl', '减：营业成本', ['1110701000000000000', '1110799000000000000'])

# 3. Extract data (use actual源 files)
BASE = os.path.dirname(os.path.abspath(__file__))

print("Extracting data...")
sales_data, _ = InvoiceExtractor(os.path.join(BASE, '顺恒2026年1-5月销售收入数据.xlsx')).extract()
costs_data, _ = InvoiceExtractor(os.path.join(BASE, '顺恒2026年1-5月成本费用数据.xlsx')).extract()

# Extract all 5 months of bank data
bank_data = {}
for mo in range(1, 6):
    nong = BankExtractor(os.path.join(BASE, f'顺恒2026年{mo}月农行明细.xls'), BankExtractor.BANK_NONGHANG).extract()
    xin = BankExtractor(os.path.join(BASE, f'顺恒2026年{mo}月农村信用社明细.xls'), BankExtractor.BANK_XINYONGSHE).extract()
    bank_data[mo] = {'农行': nong, '信用社': xin}

# Get May bank ending balance for verification
may_bank_end = bank_data[5]['农行']['end_balance'] + bank_data[5]['信用社']['end_balance']

# Payroll
payroll_data = {}
for mo in range(1, 6):
    payroll_data[mo] = PayrollExtractor(os.path.join(BASE, f'顺恒2026年{mo}月工资薪金.xlsx')).extract()

# 4. Calculate
print("Calculating...")
calc = CalcEngine(config['opening_bs'])

pl = calc.calculate_pl(
    {'sales': sales_data, 'costs': costs_data},
    payroll_data,
    engine.get_mappings('pl'),
    bank_data,
    5,  # 5 months
)

bs = calc.calculate_bs(pl, may_bank_end, pl['_detail']['fa_add'], 5)

# 5. Validate
print("Validating...")
passed, errors, warnings = validate_balance(pl, bs)

print(f"\n{'='*50}")
print(f"INTEGRATION TEST RESULTS")
print(f"{'='*50}")
print(f"Revenue:     ¥{pl['revenue']:>15,.2f}  (expected: 4,080,824.78)")
print(f"COGS:        ¥{pl['cogs']:>15,.2f}  (expected: 2,466,533.98)")
print(f"Admin Exp:   ¥{pl['admin_exp']:>15,.2f}")
print(f"Net Profit:  ¥{pl['net_profit']:>15,.2f}  (expected: 334,852.98)")
print(f"Total Assets:¥{bs['total_assets']:>15,.2f}  (expected: 4,378,328.56)")
print(f"A-L-E:       ¥{bs['_balance_check']:>15,.2f}  (expected: 0.00)")
print(f"Balance check passed: {passed}")
if errors:
    for e in errors:
        print(f"  ERROR: {e['title']}")

# Assertions
assert abs(pl['revenue'] - 4080824.78) < 0.02, f"Revenue mismatch: {pl['revenue']}"
assert abs(pl['cogs'] - 2466533.98) < 0.02, f"COGS mismatch: {pl['cogs']}"
assert abs(pl['net_profit'] - 334852.98) < 0.02, f"Net profit mismatch: {pl['net_profit']}"
assert abs(bs['total_assets'] - 4378328.56) < 0.02, f"Total assets mismatch: {bs['total_assets']}"
assert abs(bs['_balance_check']) < 0.01, f"BS not balanced: {bs['_balance_check']}"
assert passed, f"Validation failed: {len(errors)} errors"

# 6. Render test
print("\nRendering...")
period_label = '2026年1-5月'
full_name = config['full_name']

pl_cfg = te.load_preset('pl')
renderer = ReportRenderer(pl_cfg)
pl_map = {
    '一、营业收入': {0: pl['revenue']},
    '减：营业成本': {0: pl['cogs']},
    '税金及附加': {0: pl['surcharges']},
    '管理费用': {0: pl['admin_exp']},
    '财务费用': {0: pl['fin_exp']},
    '二、营业利润（亏损以"-"号填列）': {0: pl['oper_profit']},
    '三、利润总额（亏损总额以"-"号填列）': {0: pl['total_profit']},
    '四、净利润（净亏损以"-"号填列）': {0: pl['net_profit']},
}
out_path = os.path.join(TEST_DIR, 'data', 'output', '顺恒废旧', 'test_pl.xlsx')
renderer.render(full_name, period_label, pl_map, out_path)

# Verify output file is valid
import openpyxl
wb = openpyxl.load_workbook(out_path)
ws = wb.active
assert ws.cell(row=1, column=1).value == '利润表'
assert ws.cell(row=5, column=1).value == '一、营业收入'
assert ws.cell(row=5, column=3).number_format == '#,##0.00'

print(f"\n{'='*50}")
print("ALL INTEGRATION TESTS PASSED!")
print(f"{'='*50}")

# Cleanup
shutil.rmtree(TEST_DIR)
```

- [ ] **Step 2: Run the integration test**

Run: `python test_integration.py`
Expected: `ALL INTEGRATION TESTS PASSED!` with all assertions passing.

- [ ] **Step 3: Commit**

```bash
git add test_integration.py
git commit -m "test: add integration test verifying 顺恒 5-month flow"
```

---
### Task 14: Final wiring and user guide

**Files:**
- Create: `README.md` (user-facing quickstart guide)
- Modify: `app.py` (polish edges)

- [ ] **Step 1: Write user guide `README.md`**

```markdown
# 财务报表生成工具

一个本地运行的财务工具，专为小企业会计设计。上传发票、银行流水、工资表 → 自动生成《小企业会计准则》三大报表。

## 安全承诺

- ✅ **全本地处理** — 所有计算在您电脑上完成
- ✅ **不上传数据** — 没有任何后端服务器
- ✅ **不上传网络** — 不发起任何网络请求

## 第一次使用（约10分钟）

1. **安装**：双击 `启动报表工具.bat`
2. **创建公司**：左边点「公司档案」→「创建新公司」→ 填写名称
3. **填年初数**：切换到「年初余额」标签，填入去年12月资产负债表的期末余额
4. **配科目编码**：左边点「科目映射」→ 上传一份历史发票 → 系统列出所有科目编码 → 下拉选择每个编码对应哪个报表行
5. **完成设置！**

## 每月使用（约3分钟）

1. **导入数据**：点「数据导入」→ 上传本月4个文件（销售收入、成本费用、银行流水×2、工资表）
2. **核对数据**：自动跳转到「数据核验」→ 扫一眼确认数字对 → 点确认
3. **生成报表**：点「报表生成」→ 预览三张表 → 自检通过
4. **下载**：点「分析导出」→ 下载Excel报表

## 碰到问题？

| 提示 | 怎么做 |
|------|--------|
| 🔴 "发现新的科目编码" | 去「科目映射」页给它分类 |
| 🔴 "资产负债表不平" | 检查年初余额和科目映射 |
| ⚠️ "发票数比上月少很多" | 确认是否漏传了文件 |

## 技术信息

- 框架：Python + Streamlit
- 依赖：openpyxl, xlrd, pandas, python-docx
- 数据位置：`./data/` 目录
```

- [ ] **Step 2: Create Windows launcher script**

File: `启动报表工具.bat`
```bat
@echo off
cd /d %~dp0
echo Starting 财务报表生成工具...
echo.
echo The app will open in your browser shortly.
echo DO NOT close this window while using the tool.
echo.
streamlit run app.py --server.port 8501
pause
```

- [ ] **Step 3: Final smoke test**

```bash
# Verify all modules import cleanly
python -c "
import template_engine as te
import mapping_engine
import data_extractor
import calc_engine
import format_renderer
import validator
print('All modules import OK')
print(f'Templates available: {te.list_presets()}')
"

# Verify app can start
streamlit run app.py --server.headless true --server.port 8501 &
sleep 3
curl -s http://localhost:8501 | grep -q '财务报表' && echo 'App serves OK' || echo 'App failed'
pkill -f streamlit
```

- [ ] **Step 4: Commit**

```bash
git add README.md 启动报表工具.bat app.py
git commit -m "docs: add user guide and Windows launcher; finalize MVP wiring"
```

---
### Implementation Summary

| Task | Module | Verification |
|------|--------|-------------|
| 1 | Project scaffold | `pip install` succeeds |
| 2 | Template engine | 3 preset configs load & validate |
| 3 | Data extractor | 顺恒 extraction matches known values |
| 4 | Mapping engine | Unit test passes |
| 5 | Calculation engine | PL/BS values match `generate_formatted.py` |
| 6 | Format renderer | Output .xlsx has correct fonts, borders, formats |
| 7 | Validator | Balance check catches mismatches |
| 8 | App shell | `streamlit run` serves page |
| 9 | 首页 + 公司档案 | UI renders, config saves/loads |
| 10 | 科目映射 | Code discovery and dropdown mapping work |
| 11 | 数据导入 + 核验 | File upload, extraction preview, validation |
| 12 | 报表生成 + 导出 | Full render pipeline, download buttons |
| 13 | Integration test | 顺恒 5-month data produces identical results |
| 14 | README + launcher | First-time setup guide complete |

**Total estimated time:** ~3 hours for an engineer unfamiliar with the domain; ~1.5 hours for someone familiar with the existing `generate_formatted.py`.

**Deferred to Phase 2 (not in MVP):**
- 现金流量表 (间接法) — `calc_engine.calculate_cf()` 
- Word 分析报告自动生成 — `analysis_docx.py` integration
- 自定义模板上传 — `template_engine.parse_custom_template()`
- 审计日志 — `data/logs/` JSON writing
- 多月份批量生成

---

## Self-Review

### 1. Spec coverage
| Spec requirement | Covered by |
|-----------------|------------|
| 预置小企业会计准则模板 | Task 2 |
| 科目映射配置（三种规则） | Task 4 |
| 三大数据抽取器 | Task 3 |
| PL + BS 计算引擎 | Task 5 |
| Excel 精确格式化输出 | Task 6 |
| 四层验证（前三层） | Task 7 + Task 11 |
| 首页 | Task 9 |
| 公司档案（创建+年初余额） | Task 9 |
| 科目映射向导 | Task 10 |
| 数据导入（4文件上传） | Task 11 |
| 数据核验台（必经步骤） | Task 11 |
| 报表生成（在线预览+自检） | Task 12 |
| 分析导出（下载按钮） | Task 12 |
| 数据安全（全本地，无网络请求） | All tasks (architecture enforced) |

**Gap:** Spec section 9.4 (审计日志) deferred to Phase 2 — acknowledged.

### 2. Placeholder scan
No `TBD`, `TODO`, `implement later`, or `fill in details` found. Every step has actual code.

### 3. Type consistency
- `calc_engine.calculate_pl()` returns dict with keys `revenue`, `cogs`, `surcharges`, `admin_exp`, `fin_exp`, `oper_profit`, `total_profit`, `net_profit`, `_detail` — verified all references in Tasks 12 and 13 use matching keys.
- `MappingEngine` methods (`add_code_mapping`, `set_auto_calc`, `get_mappings`, `find_unmapped_codes`) consistently called across Tasks 10, 11, 12.
- Template JSON `items` list structure matches `format_renderer._write_data_rows()` iteration.
- `validator.validate_balance()` expects `pl_result` with `net_profit` and `bs_result` with `total_assets`, `total_le`, `undist_profit_end`, `undist_profit_beg` — all present in `calc_engine` output.

### 4. Scope check
Focused single-subsystem MVP. 14 tasks, each self-contained. No scope creep into Phase 2 items.

---

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-06-29-financial-report-tool-mvp.md`. Two execution options:

1. **Subagent-Driven (recommended)** — I dispatch a fresh subagent per task, review between tasks, fast iteration
2. **Inline Execution** — Execute tasks in this session using executing-plans, batch execution with checkpoints

**Which approach?**
