# 财务报表生成工具 V3 简化设计

## 背景

用户是非技术背景的"中间人"——帮做财务的朋友开发了一个 Streamlit 工具，但工具太复杂（9页、~5000行），自己和朋友都不会用、不会改。需要大幅简化。

目标用户：懂财务、不太懂电脑的财务人员。每月操作一次。

## 设计原则

- **3 步完成**：上传 → 核对 → 下载。不需要教程。
- **单一公司**：公司信息写死在配置里，砍掉多公司管理。
- **自动化**：文件类型自动识别、科目映射自动加载、模板自动选择。
- **代码精简**：5000行 → ~1500行，砍掉不需要的功能模块。

## 页面架构

3 页导航 + 简单侧边栏：

```
┌──────────────────────┐
│  📊 财务报表生成工具   │
│                      │
│  2026年 ▼  5月 ▼     │
│                      │
│  ○ ① 上传文件         │
│  ○ ② 核对数据         │
│  ○ ③ 生成报表         │
│                      │
│  公司: 顺恒废旧公司    │
└──────────────────────┘
```

### 页面1: 上传文件

- 批量拖拽上传（Ctrl多选），自动调用 smart_import.classify_file() 识别类型
- 显示识别结果：✅销售发票 ✅成本发票 ✅农行 ✅信用社 ✅工资 ⚠️未识别
- 上传完成自动保存到 `data/uploads/<公司>/<月份>/`
- 未识别文件标红提示，可手动指定类型

### 页面2: 核对数据

- 自动提取数据并计算，显示关键数字卡片：
  - 销售收入（本月/累计）
  - 成本费用（本月/累计）
  - 工资总额
  - 银行期末余额
- 发现的问题高亮显示（如：销售发票未识别到任何记录）
- 两个按钮：「确认，生成报表」「不对，重新上传」

### 页面3: 生成报表

- 四份报表下载卡片，带文件大小和生成时间：
  - 利润表_2026年X月.xlsx
  - 资产负债表_2026年X月X日.xlsx
  - 现金流量表_2026年X月.xlsx
  - 科目余额表_2026年X月.xlsx（新增）
- 「生成下个月」按钮：月份+1，跳回上传页

## 科目余额表

新增 `balance_sheet.py`（~80行）。从发票数据提取所有科目编码，汇总借/贷方：

| 科目编码 | 本期借方 | 本期贷方 | 余额 |
|---------|---------|---------|------|

数据来源：
- 销售发票科目 → 贷方
- 成本发票科目 → 借方

## 代码改造

### 保留（5个核心模块，不改）
- `data_extractor.py` — 数据提取引擎
- `calc_engine.py` — 三大报表计算
- `template_engine.py` — 模板解析和填充
- `mapping_engine.py` — 科目映射CRUD
- `smart_import.py` — 文件类型识别

### 重写（1个文件）
- `app.py`：1405行 → ~400行，9页 → 3页

### 新增（1个文件）
- `balance_sheet.py`：~80行，科目余额表生成

### 砍掉（8个模块）
- `generate_formatted.py` — 功能已被 calc_engine + template_engine 替代
- `generate_statements.py` — 合并进 app.py
- `generate_final.py` — 合并进 app.py
- `generate_analysis_docx.py` — 砍掉 Word 报告功能
- `self_check.py` — 不需要
- `annotate_images.py` — 不需要
- `validator.py` — 核对功能在页面直接展示
- `format_renderer.py` — 用模板引擎替代

### 数据流

```
上传文件(.xlsx/.xls)
  → smart_import.classify_file() 识别类型
  → 保存到 data/uploads/<公司>/<月份>/
  → data_extractor 提取数据
  → calc_engine 计算三大报表
  → template_engine 填充模板
  → 下载 .xlsx 报表
```

### 配置

`data/companies/<公司>/config.json`：
```json
{
  "full_name": "来宾市顺恒废旧汽车有限公司",
  "opening_bs": { ... },
  "templates": {
    "pl": "data/templates/<公司>/pl_custom.xlsx",
    "bs": "data/templates/<公司>/bs_custom.xlsx",
    "cf": "data/templates/<公司>/cf_custom.xlsx"
  }
}
```

## 已完成的修复（V2 → V3 过渡期间）

- calc_engine.py: 添加COGS无映射时的回退逻辑
- template_engine.py: 填充前先清空所有映射单元格，防止旧数据残留
- mapping_engine.py: 配置了顺恒废旧公司的利润表科目映射（9个营业收入编码 + 2个营业成本编码）

## 不做的

- 多公司切换管理
- Word 分析报告导出
- 历史多期对比
- 内嵌图文教程
- 自定义模板上传（改用预设模板，换公司时手动配）
- 手动科目映射配置界面（自动从 mappings.json 加载）
