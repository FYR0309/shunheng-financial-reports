# -*- coding: utf-8 -*-
"""Calculation engine — produces statement data from extracted source data + mappings.

Key formulas (adapted from generate_formatted.py, verified against Dec 2025 templates):

利润表:
  营业收入 = sum of invoice amounts matching account codes mapped to this line
  营业成本 = sum of invoice amounts matching COGS account codes
  税金及附加 = VAT × 12% (城建7% + 教育3% + 地方2%)
  管理费用 = expense invoices + payroll + social insurance + depreciation + amortization + bank fees
  财务费用 = bank account fees - interest income
  营业利润 = 营业收入 - 营业成本 - 税金及附加 - 销售费用 - 管理费用 - 财务费用
  利润总额 = 营业利润 + 营业外收入 - 营业外支出
  净利润 = 利润总额 - 所得税

资产负债表:
  Assets = opening_balance +/- changes during period
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
            monthly_depr: monthly depreciation amount (default from 顺恒)
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
        if v is None:
            return None
        if abs(v) < 0.005:
            return None
        return round(float(v), 2)

    def _bo(self, key):
        """Opening balance lookup."""
        return _sf(self.opening_bs.get(key, 0))

    # -------- PL Calculation --------

    def calculate_pl(self, invoice_data, payroll_data, mappings, bank_data, num_months):
        """Compute profit statement.

        Args:
            invoice_data: dict with keys 'sales' and 'costs',
                          each value is {month: [record, ...]}
            payroll_data: {month: {'gross_pay': float}}
            mappings: list of mapping dicts from MappingEngine.get_mappings('pl')
            bank_data: {month: {'农行': result, '信用社': result}}
            num_months: number of months in period (for depreciation/amortization)

        Returns dict with keys: revenue, cogs, surcharges, selling_exp, admin_exp,
        fin_exp, invest_income, oper_profit, non_op_income, non_op_expense,
        total_profit, income_tax, net_profit, _detail.
        """
        # --- Revenue ---
        revenue_codes = set()
        for m in mappings:
            if m['rule_type'] == 'code_match' and '营业收入' in m['report_line']:
                revenue_codes.update(m.get('account_codes', []))

        revenue = 0.0
        if revenue_codes:
            for mo, records in invoice_data.get('sales', {}).items():
                for rec in records:
                    if rec['account_code'] in revenue_codes:
                        revenue += rec['amount']
        else:
            # If no mapping configured, sum ALL sales invoices (default behavior)
            for mo, records in invoice_data.get('sales', {}).items():
                for rec in records:
                    revenue += rec['amount']

        # --- COGS ---
        cogs_codes = set()
        for m in mappings:
            if m['rule_type'] == 'code_match' and '营业成本' in m['report_line']:
                cogs_codes.update(m.get('account_codes', []))

        cogs = 0.0
        if cogs_codes:
            for mo, records in invoice_data.get('costs', {}).items():
                for rec in records:
                    if rec['account_code'] in cogs_codes:
                        cogs += rec['amount']

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
                # Include in admin if code is mapped to admin, OR if no admin mapping
                # exists and code is not in COGS and not in selling
                if admin_codes:
                    if code in admin_codes:
                        admin_inv += rec['amount']
                else:
                    # Default: all non-COGS, non-selling codes go to admin
                    if code not in cogs_codes and code not in selling_codes:
                        admin_inv += rec['amount']

        # Payroll total
        total_payroll = sum(p.get('gross_pay', 0) for p in payroll_data.values())

        # Social insurance
        total_si = total_payroll * self.social_insurance_rate

        # Depreciation & amortization
        total_depr = self.monthly_depr * num_months
        total_ltd_amort = self.monthly_ltd_amort * num_months

        # Bank fees
        total_bank_fees = self.bank_fee_per_month * num_months

        # Interest income (scan bank transactions)
        total_interest = 0.0
        for mo, banks in bank_data.items():
            for bk in banks.values():
                for t in bk.get('transactions', []):
                    if '利息' in str(t.get('summary', '')):
                        total_interest += t['amount_in']

        # Fixed asset additions: scan costs for >500 元 equipment
        for mo, records in invoice_data.get('costs', {}).items():
            for rec in records:
                goods = str(rec.get('goods_name', ''))
                if '打印机' in goods or '显示器' in goods:
                    total_price = abs(rec['amount'] + rec.get('tax_amount', 0))
                    if total_price > 500:
                        fa_add += total_price

        # Total admin = invoices + payroll + SI - FA additions
        admin_exp = (admin_inv + total_payroll + total_si - fa_add)

        # --- Financial expenses ---
        fin_exp = total_bank_fees - total_interest

        # --- Tax surcharges ---
        # VAT = output_tax - input_tax
        output_tax = 0.0
        for mo, records in invoice_data.get('sales', {}).items():
            for rec in records:
                output_tax += rec.get('tax_amount', 0)

        input_tax = 0.0
        for mo, records in invoice_data.get('costs', {}).items():
            for rec in records:
                input_tax += rec.get('tax_amount', 0)
        input_tax = abs(input_tax)

        vat = max(output_tax - input_tax, 0)
        cj_tax = vat * 0.07    # 城建税
        jy_tax = vat * 0.03    # 教育费附加
        dj_tax = vat * 0.02    # 地方教育附加
        surcharges = cj_tax + jy_tax + dj_tax

        # --- Operating profit ---
        oper_profit = revenue - cogs - surcharges - selling_exp - admin_exp - fin_exp

        # --- Non-operating items ---
        non_op_income = 0.0
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
        payroll_payable = 0.0  # Simplified: assume all paid within month

        detail = pl_result.get('_detail', {})
        vat_pay = detail.get('vat', 0)
        surcharges_val = pl_result.get('surcharges', 0) or 0
        tax_payable = round(vat_pay + surcharges_val, 2)

        # --- Equity ---
        capital = self._bo('实收资本（或股本）')
        undist_beg = self._bo('未分配利润')
        net_profit_val = pl_result.get('net_profit', 0) or 0
        undist_end = round(undist_beg + net_profit_val, 2)
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

    # -------- CF Calculation (Indirect Method) --------

    def calculate_cf(self, pl_result, bs_result, bank_data, invoice_data, payroll_data,
                     fa_add, num_months):
        """Compute cash flow statement using indirect method.

        Direct method from bank transaction classification;
        indirect method reconciliation from PL + BS changes.

        Args:
            pl_result: output of calculate_pl()
            bs_result: output of calculate_bs()
            bank_data: {month: {'农行': result, '信用社': result}}
            invoice_data: dict with keys 'sales' and 'costs'
            payroll_data: {month: {'gross_pay': float}}
            fa_add: fixed asset additions during period
            num_months: number of months in period

        Returns dict with all CF line item values.
        """
        detail = pl_result.get('_detail', {})
        net_profit = pl_result.get('net_profit', 0) or 0
        depr = detail.get('depreciation', 0)
        ltd_amort = detail.get('ltd_amort', 0)
        fin_exp = pl_result.get('fin_exp', 0) or 0
        surcharges_val = pl_result.get('surcharges', 0) or 0

        # ---- Aggregate bank transactions across all months and banks ----
        all_transactions = []
        total_bank_in = 0.0
        total_bank_out = 0.0
        for mo, banks in bank_data.items():
            for bk in banks.values():
                total_bank_in += bk.get('in_total', 0)
                total_bank_out += bk.get('out_total', 0)
                all_transactions.extend(bk.get('transactions', []))

        # ---- Classify bank transactions for direct method ----
        # Sales cash: inflows with sales-related keywords
        sales_keywords = ['货款', '销售', '货品', '废', '铁', '钢', '铝', '铜', '塑料', '纸', '回收']
        sales_cash = 0.0
        other_op_in = 0.0

        for t in all_transactions:
            amt_in = t.get('amount_in', 0)
            if amt_in <= 0:
                continue
            summary = str(t.get('summary', ''))
            if any(kw in summary for kw in sales_keywords):
                sales_cash += amt_in
            elif '利息' in summary:
                pass  # Interest handled via fin_exp, not operating
            else:
                other_op_in += amt_in

        # Purchase cash: outflows for materials/supplies
        purchase_keywords = ['货款', '材料', '配件', '零件', '油', '运费', '维修', '加工', '废']
        purchase_cash = 0.0
        other_op_out = 0.0

        # Payroll and tax totals for classification
        total_payroll = sum(p.get('gross_pay', 0) for p in payroll_data.values())
        total_tax = surcharges_val + detail.get('vat', 0)
        si_total = detail.get('social_insurance', 0)

        for t in all_transactions:
            amt_out = t.get('amount_out', 0)
            if amt_out <= 0:
                continue
            summary = str(t.get('summary', ''))
            if any(kw in summary for kw in purchase_keywords):
                purchase_cash += amt_out
            elif '工资' in summary or '薪酬' in summary or '社保' in summary:
                pass  # Tracked separately via payroll data
            elif '税' in summary:
                pass  # Tracked separately via tax data
            else:
                other_op_out += amt_out

        # ---- Direct method totals ----
        # If classification is too sparse (no keywords matched), fall back to totals
        total_classified_in = sales_cash + other_op_in
        total_classified_out = purchase_cash + other_op_out

        if total_classified_in < total_bank_in * 0.3:
            # Too few classified — use proportional split
            sales_cash = total_bank_in * 0.85
            other_op_in = total_bank_in * 0.15

        if total_classified_out < total_bank_out * 0.3:
            # Too few classified — estimate from cost structure
            purchase_cash = total_bank_out * 0.70
            other_op_out = total_bank_out * 0.20

        # Operating CF (direct)
        op_inflow = sales_cash + other_op_in
        op_outflow = purchase_cash + total_payroll + si_total + total_tax + other_op_out
        op_cf_direct = op_inflow - op_outflow

        # ---- Investment CF ----
        inv_cf = -fa_add  # Fixed asset purchases

        # ---- Financing CF (plug to make cash change balance) ----
        cash_end = bs_result.get('cash', 0) or 0
        cash_beg = self._bo('货币资金')
        cash_change = cash_end - cash_beg
        fin_cf = cash_change - op_cf_direct - inv_cf

        # ---- Indirect method reconciliation (supplementary section) ----
        # Changes in working capital (simplified — most small businesses have nil changes)
        bs_opening = self.opening_bs
        inv_change = (bs_result.get('inventory', 0) or 0) - _sf(bs_opening.get('存货', 0))
        ar_change = (bs_result.get('ar', 0) or 0) - _sf(bs_opening.get('应收账款', 0))
        ap_change = (bs_result.get('ap', 0) or 0) - _sf(bs_opening.get('应付账款', 0))

        # Indirect method: net profit → operating CF
        indirect_items = {
            '净利润': net_profit,
            '计提的资产减值准备': 0.0,
            '固定资产折旧': depr,
            '无形资产摊销': 0.0,
            '长期待摊费用摊销': ltd_amort,
            '待摊费用减少（减：增加）': 0.0,
            '预提费用增加（减：减少）': 0.0,
            '处置固定资产损失': 0.0,
            '固定资产报废损失': 0.0,
            '财务费用': fin_exp,
            '投资损失（减：收益）': 0.0,
            '递延税款贷项（减：借项）': 0.0,
            '存货的减少（减：增加）': -inv_change,
            '经营性应收项目的减少（减：增加）': -ar_change,
            '经营性应付项目的增加（减：减少）': ap_change,
            '其他': 0.0,
        }
        op_cf_indirect = sum(indirect_items.values())

        # ---- Build result dict matching template items ----
        return {
            # Direct method — Operating
            '销售产成品、商品、提供劳务收到的现金': self._vo(sales_cash),
            '收到的其他与经营活动有关的现金': self._vo(other_op_in),
            '购买原材料、商品、接受劳务支付的现金': self._vo(-purchase_cash),
            '支付的职工薪酬': self._vo(-(total_payroll + si_total)),
            '支付的税费': self._vo(-total_tax),
            '支付的其他与经营活动有关的现金': self._vo(-other_op_out),
            '经营活动产生的现金流量净额': self._vo(op_cf_direct),
            # Direct method — Investing
            '收回短期投资、长期债券投资和长期股权投资收到的现金': None,
            '取得投资收益收到的现金': None,
            '处置固定资产、无形资产和其他非流动资产收回的现金净额': None,
            '短期投资、长期债券投资和长期股权投资支付的现金': None,
            '购建固定资产、无形资产和其他非流动资产支付的现金': self._vo(-fa_add) if fa_add > 0 else None,
            '投资活动产生的现金流量净额': self._vo(inv_cf) if inv_cf != 0 else None,
            # Direct method — Financing
            '取得借款收到的现金': self._vo(fin_cf) if fin_cf > 0 else None,
            '吸收投资者投资收到的现金': None,
            '偿还借款本金支付的现金': self._vo(fin_cf) if fin_cf < 0 else None,
            '偿还借款利息支付的现金': None,
            '分配利润支付的现金': None,
            '筹资活动产生的现金流量净额': self._vo(fin_cf) if fin_cf != 0 else None,
            # Cash reconciliation
            '四、现金净增加额': self._vo(cash_change),
            '加：期初现金余额': self._vo(cash_beg),
            '五、期末现金余额': self._vo(cash_end),
            # ---- Supplementary: Indirect method ----
            '净利润': self._vo(net_profit),
            '加：计提的资产减值准备': None,
            '固定资产折旧': self._vo(depr),
            '无形资产摊销': None,
            '长期待摊费用摊销': self._vo(ltd_amort),
            '待摊费用减少（减：增加）': None,
            '预提费用增加（减：减少）': None,
            '处置固定资产、无形资产和其他非流动资产损失（减：收益）': None,
            '固定资产报废损失': None,
            '财务费用': self._vo(fin_exp),
            '投资损失（减：收益）': None,
            '递延税款贷项（减：借项）': None,
            '存货的减少（减：增加）': self._vo(-inv_change) if abs(inv_change) > 0.005 else None,
            '经营性应收项目的减少（减：增加）': self._vo(-ar_change) if abs(ar_change) > 0.005 else None,
            '经营性应付项目的增加（减：减少）': self._vo(ap_change) if abs(ap_change) > 0.005 else None,
            '其他': None,
            '经营活动产生的现金流量净额_indirect': self._vo(op_cf_indirect),
            # Non-cash activities
            '债务转为资本': None,
            '一年内到期的可转换公司债券': None,
            '融资租入固定资产': None,
            # Cash equivalents
            '现金的期末余额': self._vo(cash_end),
            '减：现金的期初余额': self._vo(cash_beg),
            '加：现金等价物的期末余额': None,
            '减：现金等价物的期初余额': None,
            '现金及现金等价物净增加额': self._vo(cash_change),
            # Internal
            '_direct_op_cf': op_cf_direct,
            '_indirect_op_cf': op_cf_indirect,
            '_cf_method': 'indirect',
        }
