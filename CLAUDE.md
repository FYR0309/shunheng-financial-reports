# 财务报表生成工具 — 项目上下文

## 项目定位

顺恒废旧公司专用财务报表生成工具（Streamlit本地Web应用）。三步完成：上传文件 → 核对数据 → 生成4份报表。

## 核心约束

- **全本地处理** — 数据不上传任何服务器，不发网络请求
- **Streamlit 框架** — 3页极简界面（上传/核对/生成）
- **小企业会计准则** — 自定义模板驱动，保留原始格式
- **单一公司** — 顺恒废旧公司，不做多公司管理

## 当前进度

**V3简化版已完成** — 2026-06-30

| 模块 | 文件 | 状态 |
|------|------|------|
| Web界面 | `app.py` (~586行) | ✅ 3页：上传→核对→生成 |
| 计算引擎 | `calc_engine.py` | ✅ PL + BS + CF |
| 科目余额表 | `trial_balance.py` | ✅ 发票科目汇总 |
| 数据抽取器 | `data_extractor.py` | ✅ 发票/银行/工资 |
| 模板引擎 | `template_engine.py` | ✅ 自定义模板解析+填充 |
| 智能导入 | `smart_import.py` | ✅ 内容+文件名兜底识别 |
| 科目映射器 | `mapping_engine.py` | ✅ code_match规则 |

## 输出报表

1. 利润表 (利润表_2026年X月.xlsx)
2. 资产负债表 (资产负债表_2026年X月X日.xlsx)
3. 现金流量表 (现金流量表_2026年X月.xlsx)
4. 科目余额表 (科目余额表_2026年X月.xlsx)

## 关键数字（顺恒已验证）

- Revenue 4,080,824.78 / COGS 2,466,533.98 / Net Profit 513,057.33
- Total Assets 4,391,280.16
- Sales month 1: ¥570,877.22 / Costs month 1: ¥444,927.61

## 关键文件

| 文件 | 用途 |
|------|------|
| `app.py` | Streamlit 入口，启动方式：`streamlit run app.py` |
| `trial_balance.py` | 科目余额表生成（原名balance_sheet.py，因pyc缓存问题重命名） |
| `calc_engine.py` | 三大报表计算引擎 |
| `template_engine.py` | 模板解析（parse_custom_template）+ 填充（fill_custom_template） |
| `data_extractor.py` | 发票/银行/工资数据抽取 |
| `smart_import.py` | 文件类型智能识别（内容检测+文件名兜底） |
| `mapping_engine.py` | 科目映射管理 |
| `data/companies/顺恒废旧公司/config.json` | 公司配置 |
| `data/companies/顺恒废旧公司/mappings.json` | 科目映射（9个营业收入编码+2个营业成本编码） |
| `data/templates/顺恒废旧公司/` | 自定义模板（pl/bs/cf_custom.xlsx） |
| `docs/superpowers/specs/2026-06-30-financial-report-tool-v3-simplification-design.md` | V3设计文档 |

## 技术细节

### 数据流
```
源文件(.xlsx/.xls) → smart_import.classify_file() 识别类型
                   → data_extractor 提取数据
                   → calc_engine 计算三大报表
                   → template_engine 填充模板
                   → trial_balance 生成科目余额表
                   → 下载 .xlsx 报表
```

### 文件名兜底识别
当 classify_file() 返回 unknown 时（Streamlit 进程中偶发），用文件名关键词检测：
- 销售/sales/收入 → sales_invoice
- 成本/cost/费用/采购 → cost_invoice
- 农行/nong → bank_nong
- 信用/信用社/xin → bank_xin
- 工资/payroll/薪金 → payroll

### COGS映射
2个税收分类编码：`1110701000000000000`, `1110799000000000000`

### 已知问题
- Streamlit模块重载可能导致.pyc缓存不一致，重命名模块可彻底解决
- classify_file() 在Streamlit中可能返回 unknown，已通过文件名兜底修复

## 启动命令

```bash
pip install -r requirements.txt
streamlit run app.py
```
