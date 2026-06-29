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


def validate_extraction(extraction_result, mapping_engine, statement_type,
                        prev_month_stats=None):
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
