# -*- coding: utf-8 -*-
"""Smart file classifier — auto-detects file type by analyzing content structure."""
import os, re
from difflib import SequenceMatcher

import openpyxl

def _fuzzy(a, b):
    return SequenceMatcher(None, a.strip(), b.strip()).ratio()

def classify_file(filepath):
    """Auto-detect what kind of financial file this is.

    Returns one of: 'sales_invoice', 'cost_invoice', 'bank_nong', 'bank_xin',
                    'payroll', 'balance_sheet', 'unknown'
    """
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in ('.xlsx', '.xls'):
        return 'unknown'

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception:
        return 'unknown'

    sheets = [s.lower() for s in wb.sheetnames]

    # Check for 金蝶 invoice format
    if any('汇总' in s for s in sheets) or any('信息' in s for s in sheets):
        ws = wb.active
        # Read first row headers
        headers = [str(c.value).strip() if c.value else '' for c in next(ws.iter_rows(min_row=1, max_row=1))]
        header_str = ' '.join(headers)

        # Check for invoice-specific columns
        has_date = any(h for h in headers if '日期' in h)
        has_code = any(h for h in headers if '科目编码' in h or '科目' in h)
        has_amount = any(h for h in headers if '金额' in h)
        has_tax = any(h for h in headers if '税额' in h)
        has_goods = any(h for h in headers if '货品' in h or '货物' in h)

        if has_date and has_code and has_amount:
            # Try to distinguish sales vs costs by checking data
            # Sales invoices tend to have fewer records, larger amounts
            # Or look for keywords in goods names
            sales_kw = 0
            cost_kw = 0
            for row in ws.iter_rows(min_row=2, max_row=min(10, ws.max_row), values_only=True):
                goods = str(row[11]) if len(row) > 11 and row[11] else ''
                if any(kw in goods for kw in ['回收', '销售', '废铁', '废铜', '废铝', '废钢']):
                    sales_kw += 1
                if any(kw in goods for kw in ['采购', '配件', '运费', '维修', '办公', '耗材', '修理', '油']):
                    cost_kw += 1

            if sales_kw > cost_kw:
                return 'sales_invoice'
            elif cost_kw > sales_kw:
                return 'cost_invoice'
            # Fallback: check amounts — costs often have more records
            row_count = sum(1 for _ in ws.iter_rows(min_row=2, values_only=True) if _[0] is not None)
            if row_count > 10:
                return 'cost_invoice'
            return 'sales_invoice'

    # Check for bank statement format
    headers = [str(c.value).strip() if c.value else '' for c in next(wb.active.iter_rows(min_row=1, max_row=1))]
    header_str = ' '.join(headers)

    if '日期' in header_str and '收入' in header_str and '支出' in header_str:
        if '摘要' in header_str:
            return 'bank_nong'
        # Check if col0 is 日期, col1 is 支出 (信用社 reverse)
        if len(headers) >= 4 and '支出' in headers[1] and '收入' in headers[2]:
            return 'bank_xin'
        return 'bank_nong'

    # Check for payroll format
    if any('工资' in h or '薪酬' in h or '应发' in h for h in headers):
        # Check for 合计 row
        for row in wb.active.iter_rows(min_row=1, max_row=min(50, wb.active.max_row), values_only=True):
            first_cell = str(row[0]).strip() if row[0] else ''
            if '合计' in first_cell or '合計' in first_cell:
                return 'payroll'

    # Check for balance sheet format
    asset_kw = sum(1 for h in headers if any(kw in h for kw in ['资产', '负债', '所有者权益', '期末余额', '年初余额']))
    if asset_kw >= 2:
        return 'balance_sheet'

    wb.close()
    return 'unknown'


def parse_opening_balance(filepath):
    """Parse a balance sheet xlsx and extract line items with ending balances.

    Handles the standard 小企业会计准则 balance sheet format:
    - Left side: assets (cols A-D)
    - Right side: liabilities + equity (cols E-H)
    - Returns: {item_name: ending_balance, ...}
    """
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception:
        return {}, ['无法打开文件']

    ws = wb.active
    result = {}
    errors = []

    # Read all rows to find the data section
    all_rows = []
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 100), values_only=True):
        all_rows.append([v for v in row])

    # Detect column structure
    # Standard format: col A=资产, col B=行次, col C=期末余额, col D=年初余额
    #                  col E=负债和所有者权益, col F=行次, col G=期末余额, col H=年初余额
    # But we need to be flexible

    # Find the data start row (skip headers)
    data_start = None
    for i, row in enumerate(all_rows):
        row_str = ' '.join([str(v) for v in row if v])
        if any(kw in row_str for kw in ['货币资金', '应收账款', '短期借款', '实收资本']):
            data_start = i
            break

    if data_start is None:
        # Try to find by scanning for common BS items
        for i, row in enumerate(all_rows):
            first = str(row[0]).strip() if row[0] else ''
            if first and len(first) > 1 and not any(kw in first for kw in ['资产', '负债', '编制', '项目', '单位']):
                data_start = i
                break

    if data_start is None:
        data_start = 3  # Default: data starts from row 4

    # Scan for asset items (left side, cols 0-3)
    # and liability items (right side, cols 4-7 if available)
    has_right_side = len(all_rows[0]) >= 7 if all_rows else False

    for row in all_rows[data_start:]:
        # Left side: asset items
        left_name = str(row[0]).strip() if len(row) > 0 and row[0] else ''
        left_end = None
        # Find the ending balance column (usually col 2 or 3)
        for ci in [2, 3]:
            if len(row) > ci and row[ci] is not None:
                try:
                    left_end = float(row[ci])
                    break
                except (TypeError, ValueError):
                    continue

        if left_name and left_end is not None and left_name not in ['', 'None']:
            # Skip section headers
            if not any(kw in left_name for kw in ['资产', '负债', '所有者权益', '流动', '非流动', '合计', '：']):
                result[left_name] = left_end
            elif '合计' in left_name:
                result[left_name] = left_end

        # Right side: liability/equity items
        if has_right_side and len(row) > 4:
            right_name = str(row[4]).strip() if row[4] else ''
            right_end = None
            for ci in [6, 7]:
                if len(row) > ci and row[ci] is not None:
                    try:
                        right_end = float(row[ci])
                        break
                    except (TypeError, ValueError):
                        continue

            if right_name and right_end is not None and right_name not in ['', 'None']:
                if not any(kw in right_name for kw in ['资产', '负债', '所有者权益', '流动', '非流动', '合计', '：']):
                    result[right_name] = right_end
                elif '合计' in right_name:
                    result[right_name] = right_end

    wb.close()
    return result, errors


def match_bs_to_config(bs_items, config_fields):
    """Fuzzy-match parsed balance sheet items to config field names.

    Args:
        bs_items: {item_name: value} from parse_opening_balance()
        config_fields: list of expected field names

    Returns:
        {config_field: value, ...} for matched items
        list of unmatched bs_items
    """
    matched = {}
    unmatched = []

    for bs_name, bs_value in bs_items.items():
        best_score = 0
        best_field = None
        for field in config_fields:
            score = _fuzzy(bs_name, field)
            # Exact substring match gets a boost
            if bs_name in field or field in bs_name:
                score += 0.3
            if score > best_score:
                best_score = score
                best_field = field

        if best_score >= 0.55 and best_field:
            matched[best_field] = bs_value
        else:
            unmatched.append((bs_name, bs_value))

    return matched, unmatched
