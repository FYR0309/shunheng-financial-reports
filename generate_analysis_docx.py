# -*- coding: utf-8 -*-
"""Generate formatted Word analysis report for 2026 Jan-May financial statements."""
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import datetime

doc = Document()

# ================================================================
# PAGE SETUP
# ================================================================
for section in doc.sections:
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.18)
    section.right_margin = Cm(3.18)

style = doc.styles['Normal']
font = style.font
font.name = '宋体'
font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
pf = style.paragraph_format
pf.space_before = Pt(0)
pf.space_after = Pt(0)
pf.line_spacing = 1.5

# ================================================================
# HELPERS
# ================================================================
def add_title(text, level=0):
    """Add a formatted title/heading"""
    if level == 0:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(12)
        run = p.add_run(text)
        run.font.name = '宋体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        run.font.size = Pt(18)
        run.bold = True
    else:
        h = doc.add_heading(text, level=level)
        for run in h.runs:
            run.font.name = '宋体'
            run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    return p if level == 0 else h

def add_para(text, bold=False, indent=True, size=10.5):
    """Add a body paragraph with proper Chinese formatting"""
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    if indent:
        p.paragraph_format.first_line_indent = Pt(21)  # 2 chars
    run = p.add_run(text)
    run.font.name = '宋体'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(size)
    run.bold = bold
    return p

def add_table(headers, rows, col_widths=None):
    """Add a formatted table"""
    table = doc.add_table(rows=len(rows)+1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    for i, hdr in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(str(hdr))
        run.font.name = '宋体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        run.font.size = Pt(9)
        run.bold = True
        # Gray background
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E2F3"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    # Data rows
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.rows[ri+1].cells[ci]
            cell.text = ''
            p = cell.paragraphs[0]
            if ci == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci <= 2 else WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(str(val) if val is not None else '')
            run.font.name = '宋体'
            run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            run.font.size = Pt(9)

    # Column widths
    if col_widths:
        for ri, row in enumerate(table.rows):
            for ci, w in enumerate(col_widths):
                row.cells[ci].width = Cm(w)

    doc.add_paragraph()  # spacer
    return table

def add_warning(text):
    """Add a highlighted warning paragraph"""
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(21)
    run = p.add_run(text)
    run.font.name = '宋体'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(10.5)
    run.bold = True
    run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
    return p

# ================================================================
# CONTENT
# ================================================================

add_title('来宾市顺恒废旧汽车回收有限公司')
add_title('2026年1-5月财务报表分析说明')

# Date and company info
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(18)
run = p.add_run(f'编制日期：2026年6月29日')
run.font.name = '宋体'
run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
run.font.size = Pt(10.5)

# ---- Section 1 ----
add_title('一、报表编制基础', level=1)

add_para('本报告基于来宾市顺恒废旧汽车回收有限公司2026年1月至5月的财务数据编制，执行《小企业会计准则》。')

add_para('数据来源包括：（1）42张销售发票，含税销售收入合计约408万元；（2）617张成本费用发票，其中营业成本约247万元、管理费用（发票部分）约95万元；（3）中国农业银行及农村信用社5个月银行流水；（4）2026年1-5月各月工资薪金表。')

add_para('主要会计估计：固定资产月折旧额25,886.46元（延续2025年度折旧政策）；长期待摊费用月摊销额180.00元；固定资产确认标准为单价超过500元的打印机、显示器等电子设备。')

# ---- Section 2 ----
add_title('二、整体经营情况', level=1)

add_table(
    ['指标', '1-5月累计（元）', '5月当月（元）', '说明'],
    [
        ['营业收入', '4,080,824.78', '915,170.87', '月均约82万元'],
        ['营业成本', '2,466,533.98', '849,790.24', '成本率60.4%'],
        ['毛利', '1,614,290.80', '65,380.63', '毛利率39.6%'],
        ['税金及附加', '3,028.71', '605.74', '增值税附加税（城建7%+教育3%+地方2%）'],
        ['管理费用', '1,276,386.61', '407,970.51', '含薪酬、社保、办公费等'],
        ['财务费用', '22.50', '4.50', '农行账户管理费'],
        ['净利润', '334,852.98', '-343,200.12', '5月单月亏损34.3万元'],
    ],
    col_widths=[3.5, 3.0, 3.0, 5.5]
)

add_title('⚠ 问题一：5月单月大幅亏损', level=2)
add_warning('5月当月亏损343,200.12元，而1-5月累计盈利仅334,852.98元。这意味着前4个月盈利约678,053元，5月一个月亏损掉近一半。')
add_para('5月营业收入915,170.87元处于正常水平，但管理费用高达407,970.51元（占收入的44.6%），远超1-4月平均水平（约217,104元/月）。建议核查5月是否存在异常大额费用支出、一次性费用确认或工资集中发放等情况。同时需关注该趋势是否会延续至6月及下半年。')

# ---- Section 3 ----
add_title('三、资产负债状况', level=1)

add_table(
    ['项目', '期末余额（元）', '年初余额（元）', '变动额（元）'],
    [
        ['货币资金', '236,209.19', '229,538.60', '+6,670.59'],
        ['应收账款', '92,618.00', '92,618.00', '0.00'],
        ['预付账款', '188,806.03', '188,806.03', '0.00'],
        ['其他应收款', '6,282.21', '6,282.21', '0.00'],
        ['存货', '481,521.58', '481,521.58', '0.00'],
        ['流动资产合计', '1,005,437.01', '998,766.42', '+6,670.59'],
        ['固定资产净值', '3,281,268.47', '3,408,965.77', '-127,697.30'],
        ['长期待摊费用', '91,623.08', '92,523.08', '-900.00'],
        ['资产总计', '4,378,328.56', '4,500,255.27', '-121,926.71'],
        ['', '', '', ''],
        ['应付账款', '13,856.00', '13,856.00', '0.00'],
        ['预收账款', '357.00', '357.00', '0.00'],
        ['应付职工薪酬', '0.00', '34,922.21', '-34,922.21'],
        ['应交税费', '28,267.94', '39,943.64', '-11,675.70'],
        ['其他应付款', '6,236,147.42', '6,646,329.20', '-410,181.78'],
        ['负债合计', '6,278,628.36', '6,735,408.05', '-456,779.69'],
        ['', '', '', ''],
        ['实收资本', '550,000.00', '550,000.00', '0.00'],
        ['未分配利润', '-2,450,299.80', '-2,785,152.78', '+334,852.98'],
        ['所有者权益合计', '-1,900,299.80', '-2,235,152.78', '+334,852.98'],
    ],
    col_widths=[4.0, 3.5, 3.5, 3.5]
)

add_title('⚠ 问题二：资不抵债——持续经营重大疑虑', level=2)
add_warning('所有者权益为-1,900,299.80元，公司已处于资不抵债状态。资产负债率高达143.4%（负债÷资产=6,278,628÷4,378,329）。')
add_para('从法律角度看，这意味着即使将所有资产变卖变现，也不足以偿还全部债务。《企业破产法》规定，企业法人不能清偿到期债务且资产不足以清偿全部债务的，可以申请破产。虽然小企业通常依靠股东或关联方持续资金支持维持运营，但这种状态已构成持续经营的重大不确定性。唯一的积极因素是，所有者权益从年初的-2,235,152.78元收窄至-1,900,299.80元，缺口在缩小。')

add_title('⚠ 问题三：其他应付款高达624万元——关联方借款是主要资金来源', level=2)
add_warning('其他应付款6,236,147.42元占负债总额的99.3%，占总资产的142.4%。该公司实质上完全依赖股东/关联方借款维持运营。')
add_para('其他应付款年初余额6,646,329.20元，5个月减少410,181.78元。这意味着公司1-5月赚取的利润大部分用于偿还关联方借款（体现在间接法现金流量表中"经营性应付项目的减少"达456,779.69元）。需要注意的是，如果关联方要求大额还款，公司现金仅236,209.19元，将面临严重的流动性危机。建议与关联方就借款金额、利率、还款期限等签订书面协议。')

add_title('⚠ 问题四：应收、预付、存货余额5个月完全未变', level=2)
add_warning('应收账款92,618元、预付账款188,806.03元、其他应收款6,282.21元、存货481,521.58元——这四个项目的期末余额与年初余额完全一致，连续5个月没有任何变动。')
add_para('可能原因包括：（1）未做往来重分类和账龄分析；（2）未对存货进行实地盘点；（3）应收账款账龄可能较长，存在坏账风险但未计提减值准备。根据《小企业会计准则》，企业应至少每年进行一次全面资产清查。建议尽快对往来款项和存货进行清理核实。')

# ---- Section 4 ----
add_title('四、现金流量分析', level=1)

add_table(
    ['项目', '金额（元）'],
    [
        ['一、经营活动现金流量净额', '6,670.59'],
        ['二、投资活动现金流量净额', '-1,735.00'],
        ['三、筹资活动现金流量净额', '1,735.00'],
        ['四、现金净增加额', '6,670.59'],
        ['加：期初现金余额', '229,538.60'],
        ['五、期末现金余额', '236,209.19'],
    ],
    col_widths=[8.0, 6.0]
)

doc.add_paragraph()

add_table(
    ['间接法：净利润→经营现金流', '金额（元）'],
    [
        ['净利润', '334,852.98'],
        ['加：固定资产折旧', '129,432.30'],
        ['加：长期待摊费用摊销', '900.00'],
        ['减：购建固定资产', '-1,735.00'],
        ['减：经营性应付项目减少', '-456,779.69'],
        ['= 经营活动现金流量净额', '6,670.59'],
    ],
    col_widths=[8.0, 6.0]
)

add_title('⚠ 问题五：经营活动现金流极弱，盈利质量堪忧', level=2)
add_warning('5个月经营活动现金净流入仅6,670.59元，月均仅1,334元。净利润334,852.98元，但现金流仅为利润的2%。')
add_para('间接法揭示了核心原因：公司虽有33.5万元利润，但偿还了45.7万元的应付项目（以其他应付款为主），导致利润几乎全部被营运资金变动吞噬。换句话说，公司"账面上赚钱了"，但钱都拿去还债了，并未形成可用的现金积累。这种盈利质量对公司的短期偿债能力和日常运营资金保障构成严重威胁。')

# ---- Section 5 ----
add_title('五、税务分析', level=1)

add_table(
    ['税种', '计税依据', '金额（元）'],
    [
        ['增值税', '销项税-进项税', '25,239.23'],
        ['城市维护建设税（7%）', '应纳增值税额', '1,766.75'],
        ['教育费附加（3%）', '应纳增值税额', '757.19'],
        ['地方教育附加（2%）', '应纳增值税额', '504.77'],
        ['税金及附加合计', '', '3,028.71'],
        ['企业所得税', '应纳税所得额', '0.00（累计亏损，无需缴纳）'],
    ],
    col_widths=[5.0, 4.5, 5.0]
)

# ---- Section 6 ----
add_title('六、综合评估', level=1)

add_title('🔴 高风险信号', level=2)
add_table(
    ['风险', '严重程度', '详细说明'],
    [
        ['资不抵债', '严重', '所有者权益-190万，资产负债率143%，持续经营能力存在重大不确定性'],
        ['关联方依赖', '严重', '其他应付款624万占负债99%，公司完全依赖关联方资金输血维持运营'],
        ['现金流断裂风险', '较高', '现金仅23.6万，5个月经营净现金流仅6,670元，基本没有"造血"能力'],
        ['5月单月亏损', '中等', '5月单月亏损34.3万，管理费用异常偏高，若趋势持续将加速恶化'],
        ['往来款长期未动', '中等', '应收/预付/存货5个月余额不变，可能存在账龄长、未对账、坏账等问题'],
    ],
    col_widths=[3.5, 2.0, 9.0]
)

add_title('🟢 相对积极的信号', level=2)
add_table(
    ['方面', '说明'],
    [
        ['毛利率39.6%', '主营业务具有一定的盈利能力和竞争力'],
        ['累计盈利33.5万元', '1-5月整体实现净利润为正，企业仍有经营价值'],
        ['净资产缺口收窄', '所有者权益从-223.5万元改善至-190万元，趋势向好'],
        ['固定资产基础', '固定资产净值328万元，具备一定的资产基础和生产经营条件'],
    ],
    col_widths=[4.0, 10.5]
)

# ---- Section 7 ----
add_title('七、建议', level=1)

suggestions = [
    ('与关联方明确债务安排', '624万元其他应付款应补签书面借款协议，明确借款金额、利率、还款计划和担保措施。区分股东投入与外部借款，必要时将部分借款转为增资以改善资本结构。'),
    ('查明5月亏损原因', '逐项核查5月管理费用明细，重点关注是否有一次性大额支出、费用跨期错配、工资集中发放等情况。评估该亏损是否为孤立事件还是趋势性恶化。'),
    ('加强往来款管理', '对应收账款进行账龄分析，对长账龄款项进行催收；对预付账款核实业务进度；对存货进行实地盘点，确保账实相符。'),
    ('编制现金流预算', '现金余额仅23.6万元，建议编制未来3-6个月的月度现金流预算表，提前识别资金缺口并安排融资。'),
    ('考虑股东增资', '从长期看，改善资本结构的根本途径是股东追加投资或引入新投资者。实收资本仅55万元，而累计亏损已达245万元，增资是解决资不抵债问题的最直接方式。'),
    ('规范财务核算', '建议按月进行往来重分类、折旧计提、税费计算，避免期末集中处理；建立固定资产卡片和存货进销存台账。'),
]

for i, (title, desc) in enumerate(suggestions, 1):
    add_para(f'{i}. {title}', bold=True)
    add_para(desc)

# ---- Footer ----
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('— 本报告基于2026年1-5月账面数据编制，未经审计 —')
run.font.name = '宋体'
run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

# ================================================================
# SAVE
# ================================================================
doc.save('2026年1-5月财务报表分析说明.docx')
print("Saved: 2026年1-5月财务报表分析说明.docx")
print("Done!")
