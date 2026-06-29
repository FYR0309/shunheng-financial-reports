# 财务报表生成工具 — 项目上下文

## 项目定位

把顺恒废旧公司的报表生成脚本改造为**通用本地Web应用**。财务人员通过浏览器上传发票/银行流水/工资表 → 自动生成《小企业会计准则》三大报表。

## 核心约束

- **全本地处理** — 数据不上传任何服务器，不发网络请求
- **Streamlit 框架** — 单页应用，侧边栏导航
- **小企业会计准则** — 预置标准模板，用户可配科目映射

## 当前进度

**MVP（第一期）已完成** — 2026-06-29

| 模块 | 文件 | 状态 |
|------|------|------|
| 模板引擎 | `template_engine.py` | ✅ 预置三表JSON配置 |
| 数据抽取器 | `data_extractor.py` | ✅ 发票/银行/工资 |
| 科目映射器 | `mapping_engine.py` | ✅ 三种规则类型 |
| 计算引擎 | `calc_engine.py` | ✅ PL + BS 计算 |
| 格式渲染器 | `format_renderer.py` | ✅ 宋体/Arial，完整边框 |
| 验证器 | `validator.py` | ✅ 三层检查 |
| Web界面 | `app.py` | ✅ 7个页面全部实现 |

## 待开发（第二期）

- [ ] 现金流量表（间接法）完整生成 `calc_engine.calculate_cf()`
- [ ] Word 分析报告 `analysis_docx.py` 集成到 app.py
- [ ] 自定义模板上传解析 `template_engine.parse_custom_template()`
- [ ] 审计日志 `data/logs/`
- [ ] 多月份批量生成

## 关键数字（顺恒已验证）

- Revenue 4,080,824.78 / COGS 2,466,533.98 / Net Profit 334,852.98
- Total Assets 4,378,328.56 / A-L-E = 0.00
- Monthly depreciation 25,886.46 / Monthly LTD amort 180.00
- Bank fee 4.50/month / Social insurance 25% of gross payroll

## 关键文件

| 文件 | 用途 |
|------|------|
| `app.py` | Streamlit 入口，启动方式：`streamlit run app.py` 或双击 `启动报表工具.bat` |
| `generate_formatted.py` | 原始顺恒专用脚本（保留参考，不再开发） |
| `templates/*.json` | 预置报表模板配置（PL/BS/CF） |
| `data/` | 运行时目录（公司档案、上传、输出、日志） |
| `docs/superpowers/specs/2026-06-29-financial-report-tool-design.md` | 设计文档 |
| `docs/superpowers/plans/2026-06-29-financial-report-tool-mvp.md` | 实施计划 |
| `.gitignore` | 已排除源数据文件（*.xlsx/*.xls）和data/companies/ |
| `requirements.txt` | Python依赖 |

## 技术细节

### 数据流
```
源文件(.xlsx/.xls) → data_extractor → {标准化dict}
                                    → calc_engine → {计算结果}
                                    → format_renderer → .xlsx输出
```

### 科目映射三种规则
1. `code_match` — 源数据科目编码 → 报表行
2. `auto_calc` — 计算引擎产出（折旧、税金附加）
3. `multi_source` — 多来源汇总（管理费用 = 发票 + 工资 + 社保 + 折旧）

### 资产负债表其他应付款
倒挤（plug figure）: 其他应付款 = 资产总计 - 所有已知负债 - 所有者权益
标注为"待确认"，用户可手动覆盖。

### 银行流水格式
- 农行：正序，col 0=日期 1=收入 2=支出 3=余额 7=摘要，数据从第3行
- 信用社：倒序，col 0=日期 1=支出 2=收入 3=余额，数据从第4行

### 发票格式（金蝶导出）
- Sheet名: "信息汇总表"
- col 8=日期, 9=科目编码, 11=货品名称, 16=金额, 18=税额

## 启动命令

```bash
# 安装依赖
pip install -r requirements.txt

# 启动应用
streamlit run app.py

# 或双击 Windows 用户
启动报表工具.bat
```

## 测试验证

```bash
# 运行生成好的顺恒数据验证
python self_check.py

# 生成原始报表
python generate_formatted.py
```
