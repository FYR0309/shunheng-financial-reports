# -*- coding: utf-8 -*-
"""财务报表生成工具 — 三步生成三大报表 + 科目余额表.

① 上传文件 → ② 核对数据 → ③ 生成报表
"""
import streamlit as st
import os, json, shutil
from datetime import datetime
from collections import defaultdict

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
TEMPLATES_DIR = os.path.join(DATA_DIR, 'templates')

for d in [COMPANIES_DIR, UPLOADS_DIR, OUTPUT_DIR]:
    os.makedirs(d, exist_ok=True)

# ---- Default company ----
DEFAULT_COMPANY = '顺恒废旧公司'

# ---- Session state ----
DEFAULTS = {
    'current_company': DEFAULT_COMPANY,
    'current_year': 2026,
    'current_month': 5,
    '_last_year': None,         # track changes to invalidate cache
    '_last_month': None,
    'page': 'upload',
    'uploaded_files': {},       # {file_type: (filepath, filename)}
    'classified_files': [],     # [(filename, file_type), ...]
    'unknown_files': [],        # [filename, ...]
    'extracted_data': None,     # cached extraction result
    'generated_reports': {},    # {type: filepath}
}
for key, val in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = val


# ---- Helpers ----
def get_company_dir(name):
    return os.path.join(COMPANIES_DIR, name)

def load_company_config(name):
    path = os.path.join(get_company_dir(name), 'config.json')
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def get_upload_dir(company, year, month):
    return os.path.join(UPLOADS_DIR, company, f'{year}-{month:02d}')

def get_template_dir(company):
    return os.path.join(TEMPLATES_DIR, company)


def load_period_bank_payroll(company, year, months):
    """Scan multiple months' upload dirs for bank/payroll data.

    Returns:
        bank_data: {month: {'农行': data, '信用社': data}, ...}
        payroll_data: {month: {gross_pay, ...}, ...}
    """
    from data_extractor import BankExtractor, PayrollExtractor

    bank_data = {}
    payroll_data = {}

    for mo in months:
        upload_dir = get_upload_dir(company, year, mo)
        if not os.path.exists(upload_dir):
            continue

        for fname in os.listdir(upload_dir):
            fpath = os.path.join(upload_dir, fname)
            fn_lower = fname.lower()
            if '农行' in fname or 'nong' in fn_lower:
                try:
                    ext = BankExtractor(fpath, 'nonghang')
                    bank_data.setdefault(mo, {})['农行'] = ext.extract()
                except Exception:
                    pass
            elif '信用' in fname or 'xin' in fn_lower:
                try:
                    ext = BankExtractor(fpath, 'xinyongshe')
                    bank_data.setdefault(mo, {})['信用社'] = ext.extract()
                except Exception:
                    pass
            elif '工资' in fname or 'payroll' in fn_lower or '薪金' in fname:
                try:
                    ext = PayrollExtractor(fpath)
                    payroll_data[mo] = ext.extract()
                except Exception:
                    pass

    return bank_data, payroll_data


# ---- Sidebar ----
with st.sidebar:
    st.title('📊 财务报表工具')

    # Company (fixed display, no switching needed for V3)
    company = st.session_state.current_company
    config = load_company_config(company)
    st.caption(f'公司：{config.get("full_name", company)}')

    # Year / Month
    c1, c2 = st.columns(2)
    with c1:
        years = list(range(2023, 2031))
        st.session_state.current_year = st.selectbox(
            '年份', years, index=years.index(st.session_state.current_year),
            label_visibility='collapsed')
    with c2:
        st.session_state.current_month = st.selectbox(
            '月份', list(range(1, 13)), index=st.session_state.current_month - 1,
            label_visibility='collapsed')

    st.divider()

    # 3-step navigation
    pages = {
        'upload': '① 上传文件',
        'review': '② 核对数据',
        'generate': '③ 生成报表',
    }
    page_keys = list(pages.keys())
    current_idx = page_keys.index(st.session_state.page) if st.session_state.page in page_keys else 0
    selected_label = st.radio(
        '导航',
        list(pages.values()),
        index=current_idx,
    )
    # Map label back to key
    label_to_key = {v: k for k, v in pages.items()}
    st.session_state.page = label_to_key[selected_label]

    st.divider()
    st.caption(f'数据：{DATA_DIR}')

# Invalidate cache when month/year changes
if (st.session_state._last_year != st.session_state.current_year or
        st.session_state._last_month != st.session_state.current_month):
    st.session_state.extracted_data = None
    st.session_state.generated_reports = {}
    st.session_state._last_year = st.session_state.current_year
    st.session_state._last_month = st.session_state.current_month


# ---- Import backend modules (lazy) ----
from smart_import import classify_file
from mapping_engine import MappingEngine
from data_extractor import InvoiceExtractor, BankExtractor, PayrollExtractor
from calc_engine import CalcEngine
from template_engine import parse_custom_template, fill_custom_template
from trial_balance import generate_trial_balance


# ================================================================
# ① UPLOAD PAGE
# ================================================================
if st.session_state.page == 'upload':
    st.header('① 上传文件')

    year = st.session_state.current_year
    month = st.session_state.current_month
    upload_dir = get_upload_dir(company, year, month)
    period_label = f'{year}年{month}月'

    st.subheader(f'📤 {period_label} 数据文件')

    # Batch upload
    st.markdown('**把本月所有文件一起拖入（可 Ctrl+多选），系统自动识别类型。**')
    uploaded = st.file_uploader(
        '拖拽或点击选择文件',
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        key='batch_upload',
        label_visibility='collapsed',
    )

    if uploaded:
        # Save and classify each file
        os.makedirs(upload_dir, exist_ok=True)
        classified = []
        unknown = []

        for f in uploaded:
            # Save
            filepath = os.path.join(upload_dir, f.name)
            with open(filepath, 'wb') as out:
                out.write(f.getbuffer())

            # Classify
            ftype = classify_file(filepath)
            # Fallback: filename-based detection
            if ftype == 'unknown':
                fn_lower = f.name.lower()
                if '销售' in f.name or 'sales' in fn_lower or '收入' in f.name:
                    ftype = 'sales_invoice'
                elif '成本' in f.name or 'cost' in fn_lower or '费用' in f.name or '采购' in f.name:
                    ftype = 'cost_invoice'
                elif '农行' in f.name or 'nong' in fn_lower:
                    ftype = 'bank_nong'
                elif '信用' in f.name or 'xin' in fn_lower:
                    ftype = 'bank_xin'
                elif '工资' in f.name or 'payroll' in fn_lower or '薪金' in f.name:
                    ftype = 'payroll'
            classified.append((f.name, ftype))
            if ftype == 'unknown':
                unknown.append(f.name)

        st.session_state.classified_files = classified
        st.session_state.unknown_files = unknown

        # Show results
        st.divider()
        st.subheader('识别结果')

        type_labels = {
            'sales_invoice': '📈 销售发票',
            'cost_invoice': '📉 成本发票',
            'bank_nong': '🏦 农行流水',
            'bank_xin': '🏛️ 信用社流水',
            'payroll': '👥 工资表',
            'balance_sheet': '📋 报表文件（参考）',
            'unknown': '⚠️ 未识别',
        }

        for fname, ftype in classified:
            label = type_labels.get(ftype, f'❓ {ftype}')
            if ftype == 'unknown':
                st.error(f'{label} — {fname}')
            elif ftype == 'balance_sheet':
                st.info(f'{label} — {fname}')
            else:
                st.success(f'{label} — {fname}')

        if unknown:
            st.warning(f'⚠️ {len(unknown)} 个文件未能识别，请检查文件格式。')

        if not unknown:
            st.success('✅ 所有文件已识别，可以进入下一步了！')

    # Show existing files if any
    if os.path.exists(upload_dir):
        existing = os.listdir(upload_dir)
        if existing and not uploaded:
            st.info(f'📁 本月已上传 {len(existing)} 个文件')

    # Navigation
    st.divider()
    c1, c2 = st.columns([1, 4])
    with c1:
        if st.button('进入核对 →', type='primary', use_container_width=True):
            st.session_state.extracted_data = None  # force re-extraction
            st.session_state.generated_reports = {}
            st.session_state.page = 'review'
            st.rerun()


# ================================================================
# ② REVIEW PAGE
# ================================================================
if st.session_state.page == 'review':
    st.header('② 核对数据')

    year = st.session_state.current_year
    month = st.session_state.current_month
    upload_dir = get_upload_dir(company, year, month)

    # Check if files exist
    if not os.path.exists(upload_dir) or not os.listdir(upload_dir):
        st.warning('⚠️ 本月还没有上传文件，请先上传。')
        if st.button('← 返回上传'):
            st.session_state.page = 'upload'
            st.rerun()
    else:
        # Extract data (cached in session state)
        if st.session_state.extracted_data is None:
            with st.spinner('正在提取数据...'):
                # Load config
                config = load_company_config(company)
                opening_bs = config.get('opening_bs', {})

                # Find files by type
                all_files = os.listdir(upload_dir)
                sales_file = None
                costs_file = None
                nong_file = None
                xin_file = None
                payroll_file = None

                for fname in all_files:
                    fpath = os.path.join(upload_dir, fname)
                    # Content-based classification
                    try:
                        ftype = classify_file(fpath)
                    except Exception:
                        ftype = 'unknown'
                    # Fallback: filename-based detection
                    if ftype == 'unknown':
                        fn_lower = fname.lower()
                        if '销售' in fname or 'sales' in fn_lower or '收入' in fname:
                            ftype = 'sales_invoice'
                        elif '成本' in fname or 'cost' in fn_lower or '费用' in fname or '采购' in fname:
                            ftype = 'cost_invoice'
                        elif '农行' in fname or 'nong' in fn_lower:
                            ftype = 'bank_nong'
                        elif '信用' in fname or '信用社' in fname or 'xin' in fn_lower:
                            ftype = 'bank_xin'
                        elif '工资' in fname or 'payroll' in fn_lower or '薪金' in fname:
                            ftype = 'payroll'
                    if ftype == 'sales_invoice' and not sales_file:
                        sales_file = fpath
                    elif ftype == 'cost_invoice' and not costs_file:
                        costs_file = fpath
                    elif ftype == 'bank_nong' and not nong_file:
                        nong_file = fpath
                    elif ftype == 'bank_xin' and not xin_file:
                        xin_file = fpath
                    elif ftype == 'payroll' and not payroll_file:
                        payroll_file = fpath

                # Extract
                sales_data = {}
                costs_data = {}
                nong_data = None
                xin_data = None
                payroll_data = {}

                if sales_file:
                    ext = InvoiceExtractor(sales_file)
                    sales_data, _ = ext.extract()
                if costs_file:
                    ext = InvoiceExtractor(costs_file)
                    costs_data, _ = ext.extract()
                if nong_file:
                    ext = BankExtractor(nong_file, 'nonghang')
                    nong_data = ext.extract()
                if xin_file:
                    ext = BankExtractor(xin_file, 'xinyongshe')
                    xin_data = ext.extract()
                if payroll_file:
                    ext = PayrollExtractor(payroll_file)
                    payroll_data = {month: ext.extract()}

                # How many months of data?
                all_months = set()
                for mo in sales_data:
                    all_months.add(mo)
                for mo in costs_data:
                    all_months.add(mo)
                num_months = max(1, len(all_months))

                # Load mappings
                mapping_engine = MappingEngine(get_company_dir(company))
                pl_mappings = mapping_engine.get_mappings('pl')

                # Calculate
                calc = CalcEngine(opening_bs)
                invoice_data = {'sales': sales_data, 'costs': costs_data}
                bank_data = {month: {}}
                if nong_data:
                    bank_data[month]['农行'] = nong_data
                if xin_data:
                    bank_data[month]['信用社'] = xin_data

                pl = calc.calculate_pl(invoice_data, payroll_data, pl_mappings, bank_data, num_months)

                bank_end = 0.0
                if nong_data:
                    bank_end += nong_data['end_balance']
                if xin_data:
                    bank_end += xin_data['end_balance']

                bs = calc.calculate_bs(pl, bank_end, pl['_detail']['fa_add'], num_months)
                cf = calc.calculate_cf(pl, bs, bank_data, invoice_data, payroll_data,
                                       pl['_detail']['fa_add'], num_months)

                # Get sales/costs totals for current month specifically
                sales_month = sum(r['amount'] for r in sales_data.get(month, []))
                costs_month = sum(r['amount'] for r in costs_data.get(month, []))
                payroll_month = payroll_data.get(month, {}).get('gross_pay', 0)

                st.session_state.extracted_data = {
                    'pl': pl,
                    'bs': bs,
                    'cf': cf,
                    'sales_total': sales_month,
                    'costs_total': costs_month,
                    'payroll_total': payroll_month,
                    'bank_end': bank_end,
                    'num_months': num_months,
                    'sales_data': sales_data,
                    'costs_data': costs_data,
                    'payroll_data': payroll_data,
                    'nong_data': nong_data,
                    'xin_data': xin_data,
                    'bank_data': bank_data,
                    'invoice_data': invoice_data,
                    'opening_bs': opening_bs,
                    'pl_mappings': pl_mappings,
                }

        # Show summary cards
        data = st.session_state.extracted_data
        pl = data['pl']
        bs = data['bs']

        st.subheader(f'📊 {year}年{month}月 数据概览')

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            rev = data['sales_total']
            st.metric('本月销售收入', f'¥{rev:,.2f}')
        with col2:
            cos = data['costs_total']
            st.metric('本月成本费用', f'¥{cos:,.2f}')
        with col3:
            pay = data['payroll_total']
            st.metric('本月工资', f'¥{pay:,.2f}')
        with col4:
            bal = data['bank_end']
            st.metric('银行期末余额', f'¥{bal:,.2f}')

        # Profit summary
        st.divider()
        col1, col2, col3 = st.columns(3)
        with col1:
            v = pl.get('revenue')
            st.metric('累计营业收入', f'¥{v:,.2f}' if v else '—')
        with col2:
            v = pl.get('net_profit')
            st.metric('累计净利润', f'¥{v:,.2f}' if v else '—')
        with col3:
            v = bs.get('total_assets')
            st.metric('资产总计', f'¥{v:,.2f}' if v else '—')

        # Action buttons
        st.divider()
        c1, c2, c3 = st.columns([1, 1, 4])
        with c1:
            if st.button('← 重新上传', use_container_width=True):
                st.session_state.extracted_data = None
                st.session_state.page = 'upload'
                st.rerun()
        with c2:
            if st.button('✅ 确认，生成报表', type='primary', use_container_width=True):
                st.session_state.page = 'generate'
                st.rerun()


# ================================================================
# ③ GENERATE PAGE
# ================================================================
if st.session_state.page == 'generate':
    st.header('③ 生成报表')

    if st.session_state.extracted_data is None:
        st.warning('⚠️ 请先完成数据核对。')
        if st.button('← 返回核对'):
            st.session_state.page = 'review'
            st.rerun()
    else:
        data = st.session_state.extracted_data
        year = st.session_state.current_year
        month = st.session_state.current_month

        # ---- Period selector ----
        period_type = st.radio(
            '📅 报表范围',
            ['仅本月', '本年至本月', '全年'],
            horizontal=True,
            index=1,
            key='_period_type',
        )

        # Invalidate cached reports when period type changes
        last_pt = st.session_state.get('_last_period_type')
        if last_pt is not None and last_pt != period_type:
            st.session_state.generated_reports = {}
        st.session_state['_last_period_type'] = period_type

        if period_type == '仅本月':
            report_months = [month]
            period_label = f'{year}年{month}月'
        elif period_type == '本年至本月':
            report_months = list(range(1, month + 1))
            period_label = f'{year}年1-{month}月'
        else:  # 全年
            report_months = list(range(1, 13))
            period_label = f'{year}年度'

        # BS date uses last day of month (always end-of-month snapshot)
        import calendar
        last_day = calendar.monthrange(year, month)[1]
        bs_date = f'{year}年{month}月{last_day}日'

        config = load_company_config(company)
        company_full = config.get('full_name', company)

        template_dir = get_template_dir(company)
        output_dir = os.path.join(OUTPUT_DIR, company, f'{year}-{month:02d}')
        os.makedirs(output_dir, exist_ok=True)

        # Only generate if not already done
        if not st.session_state.generated_reports:
            with st.spinner('正在生成报表...'):
                # Filter invoice data to selected months
                filtered_sales = {m: recs for m, recs in data['sales_data'].items()
                                  if m in report_months}
                filtered_costs = {m: recs for m, recs in data['costs_data'].items()
                                  if m in report_months}

                # Load multi-month bank/payroll data
                period_bank, period_payroll = load_period_bank_payroll(
                    company, year, report_months)
                if not period_payroll:
                    period_payroll = data.get('payroll_data', {})
                if not period_bank:
                    period_bank = data.get('bank_data', {})

                num_months = len(report_months)
                mf = 1.0 / max(num_months, 1)

                # Recalculate with filtered data
                calc = CalcEngine(data['opening_bs'])
                invoice_data = {'sales': filtered_sales, 'costs': filtered_costs}
                pl = calc.calculate_pl(invoice_data, period_payroll,
                                       data['pl_mappings'], period_bank, num_months)

                # Bank end balance: sum last-known balance across banks
                bank_end = 0.0
                for m in report_months:
                    mb = period_bank.get(m, {})
                    for bn in ['农行', '信用社']:
                        bd = mb.get(bn)
                        if isinstance(bd, dict):
                            bank_end += bd.get('end_balance', 0)

                bs = calc.calculate_bs(pl, bank_end, pl['_detail']['fa_add'], num_months)
                cf = calc.calculate_cf(pl, bs, period_bank, invoice_data, period_payroll,
                                       pl['_detail']['fa_add'], num_months)
                detail = pl.get('_detail', {})

                # === 利润表 ===
                pl_tmpl_path = os.path.join(template_dir, 'pl_custom.xlsx')
                if os.path.exists(pl_tmpl_path):
                    pl_tmpl = parse_custom_template(pl_tmpl_path)
                    pl_data = {
                        '一、营业收入': {0: pl.get('revenue'), 1: (pl.get('revenue') or 0) * mf},
                        '减：营业成本': {0: pl.get('cogs'), 1: (pl.get('cogs') or 0) * mf},
                        '税金及附加': {0: pl.get('surcharges'), 1: (pl.get('surcharges') or 0) * mf},
                        '销售费用': {0: pl.get('selling_exp'), 1: None},
                        '管理费用': {0: pl.get('admin_exp'), 1: (pl.get('admin_exp') or 0) * mf},
                        '财务费用': {0: pl.get('fin_exp'), 1: (pl.get('fin_exp') or 0) * mf},
                        '其中：利息费用（收入以"-"号填列）': {0: -detail.get('interest_income', 0) if detail.get('interest_income') else None, 1: None},
                        '二、营业利润（亏损以"-"号填列）': {0: pl.get('oper_profit'), 1: (pl.get('oper_profit') or 0) * mf},
                        '三、利润总额（亏损总额以"-"号填列）': {0: pl.get('total_profit'), 1: (pl.get('total_profit') or 0) * mf},
                        '四、净利润（净亏损以"-"号填列）': {0: pl.get('net_profit'), 1: (pl.get('net_profit') or 0) * mf},
                    }
                    pl_out = os.path.join(output_dir, f'利润表_{period_label}.xlsx')
                    fill_custom_template(pl_tmpl_path, pl_tmpl['field_map'], pl_data, pl_out,
                                         company_name=company_full, period_label=period_label)
                    st.session_state.generated_reports['pl'] = pl_out

                # === 资产负债表 ===
                bs_tmpl_path = os.path.join(template_dir, 'bs_custom.xlsx')
                if os.path.exists(bs_tmpl_path):
                    bs_tmpl = parse_custom_template(bs_tmpl_path)
                    bs_data = {
                        '货币资金': {0: bs.get('cash'), 1: bs.get('cash')},
                        '应收账款': {0: bs.get('ar'), 1: bs.get('ar')},
                        '预付账款': {0: bs.get('prepay'), 1: bs.get('prepay')},
                        '其他应收款': {0: bs.get('other_recv'), 1: bs.get('other_recv')},
                        '存货': {0: bs.get('inventory'), 1: bs.get('inventory')},
                        '流动资产合计': {0: bs.get('curr_assets'), 1: bs.get('curr_assets')},
                        '固定资产原价': {0: bs.get('fa_orig'), 1: bs.get('fa_orig')},
                        '减：累计折旧': {0: bs.get('acc_depr'), 1: bs.get('acc_depr')},
                        '固定资产账面价值': {0: bs.get('fa_net'), 1: bs.get('fa_net')},
                        '长期待摊费用': {0: bs.get('ltd'), 1: bs.get('ltd')},
                        '非流动资产合计': {0: bs.get('non_curr_assets'), 1: bs.get('non_curr_assets')},
                        '资产总计': {0: bs.get('total_assets'), 1: bs.get('total_assets')},
                        '应付账款': {0: bs.get('ap'), 1: bs.get('ap')},
                        '预收账款': {0: bs.get('pr'), 1: bs.get('pr')},
                        '应付职工薪酬': {0: bs.get('payroll_payable'), 1: bs.get('payroll_payable')},
                        '应交税费': {0: bs.get('tax_payable'), 1: bs.get('tax_payable')},
                        '其他应付款': {0: bs.get('other_pay'), 1: bs.get('other_pay')},
                        '流动负债合计': {0: bs.get('curr_liab'), 1: bs.get('curr_liab')},
                        '负债合计': {0: bs.get('total_liab'), 1: bs.get('total_liab')},
                        '实收资本（或股本）': {0: bs.get('capital'), 1: bs.get('capital')},
                        '未分配利润': {0: bs.get('undist_profit_end'), 1: bs.get('undist_profit_end')},
                        '所有者权益（或股东权益）合计': {0: bs.get('total_equity'), 1: bs.get('total_equity')},
                        '负债和所有者权益（或股东权益）总计': {0: bs.get('total_le'), 1: bs.get('total_le')},
                    }
                    bs_out = os.path.join(output_dir, f'资产负债表_{bs_date}.xlsx')
                    fill_custom_template(bs_tmpl_path, bs_tmpl['field_map'], bs_data, bs_out,
                                         company_name=company_full, period_label=bs_date)
                    st.session_state.generated_reports['bs'] = bs_out

                # === 现金流量表 ===
                cf_tmpl_path = os.path.join(template_dir, 'cf_custom.xlsx')
                if os.path.exists(cf_tmpl_path):
                    cf_tmpl = parse_custom_template(cf_tmpl_path)
                    cf_data = {}
                    for k, v in cf.items():
                        if not k.startswith('_') and v is not None:
                            cf_data[k] = {0: v, 1: v * mf if isinstance(v, (int, float)) else v}
                    cf_out = os.path.join(output_dir, f'现金流量表_{period_label}.xlsx')
                    fill_custom_template(cf_tmpl_path, cf_tmpl['field_map'], cf_data, cf_out,
                                         company_name=company_full, period_label=period_label)
                    st.session_state.generated_reports['cf'] = cf_out

                # === 科目余额表 ===
                tb_out = os.path.join(output_dir, f'科目余额表_{period_label}.xlsx')
                generate_trial_balance(
                    filtered_sales, filtered_costs,
                    company_name=company_full, period_label=period_label,
                    output_path=tb_out,
                )
                st.session_state.generated_reports['tb'] = tb_out

        # Show download cards
        st.subheader(f'📥 {period_label} 报表')

        report_labels = {
            'pl': ('📈 利润表', f'利润表_{period_label}.xlsx'),
            'bs': ('📋 资产负债表', f'资产负债表_{bs_date}.xlsx'),
            'cf': ('💵 现金流量表', f'现金流量表_{period_label}.xlsx'),
            'tb': ('📊 科目余额表', f'科目余额表_{period_label}.xlsx'),
        }

        for key, (label, fname) in report_labels.items():
            path = st.session_state.generated_reports.get(key)
            if path and os.path.exists(path):
                col1, col2 = st.columns([3, 1])
                with col1:
                    size_kb = os.path.getsize(path) / 1024
                    st.success(f'{label} — {size_kb:.1f} KB')
                with col2:
                    with open(path, 'rb') as f:
                        st.download_button(
                            '⬇ 下载',
                            f.read(),
                            file_name=fname,
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            key=f'dl_{key}',
                            use_container_width=True,
                        )

        st.divider()
        c1, c2 = st.columns([1, 4])
        with c1:
            if st.button('→ 生成下个月', type='primary', use_container_width=True):
                next_month = st.session_state.current_month + 1
                if next_month > 12:
                    st.session_state.current_month = 1
                    st.session_state.current_year += 1
                else:
                    st.session_state.current_month = next_month
                st.session_state.extracted_data = None
                st.session_state.generated_reports = {}
                st.session_state.uploaded_files = {}
                st.session_state.classified_files = []
                st.session_state.unknown_files = []
                st.session_state.page = 'upload'
                st.rerun()
