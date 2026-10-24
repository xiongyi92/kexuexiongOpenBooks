# -*- coding: utf-8 -*-
"""生成《晨光贸易 2026 年 7 月账务练习账套》练习册与参考答案两个工作簿。"""
import os
import sys

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl>=3.1.0"])
    import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import model as M

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.dirname(HERE)   # 脚本放在「生成脚本/」下，产物输出到上一级目录
F_PRACTICE = os.path.join(OUT_DIR, "晨光贸易2026年7月_账务练习数据.xlsx")
F_ANSWER = os.path.join(OUT_DIR, "晨光贸易2026年7月_参考答案.xlsx")


def xl_color(css_hex: str) -> str:
    v = css_hex.removeprefix("#").upper()
    if len(v) != 6:
        raise ValueError(f"Expected #RRGGBB, got: {css_hex}")
    return "FF" + v


XL_PRIMARY = xl_color("#4472C4")
XL_DARK = xl_color("#2F5597")
XL_LIGHT = xl_color("#D9E2F3")
XL_GREY = xl_color("#F2F2F2")
XL_BORDER = xl_color("#BFBFBF")
XL_WHITE = xl_color("#FFFFFF")
XL_WARN = xl_color("#FFF2CC")
XL_GREEN = xl_color("#C6EFCE")
XL_GREEN_T = xl_color("#006100")
XL_RED_T = xl_color("#9C0006")

FONT = "微软雅黑"
thin = Side(style="thin", color=XL_BORDER)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
TOP_LINE = Border(left=thin, right=thin, top=Side(style="medium", color=XL_PRIMARY), bottom=thin)

MONEY = '#,##0.00'
DATE_FMT = 'yyyy-mm-dd'
PCT = '0.00%'

H_TITLE = Font(name=FONT, size=14, bold=True, color=XL_WHITE)
H_SUB = Font(name=FONT, size=9, color="FF595959")
H_HEAD = Font(name=FONT, size=10, bold=True, color=XL_WHITE)
H_BODY = Font(name=FONT, size=10, color="FF1A1A1A")
H_BOLD = Font(name=FONT, size=10, bold=True, color="FF1A1A1A")
H_SECTION = Font(name=FONT, size=11, bold=True, color=XL_DARK)

FILL_HEAD = PatternFill("solid", fgColor=XL_PRIMARY)
FILL_TOTAL = PatternFill("solid", fgColor=XL_DARK)
FILL_LIGHT = PatternFill("solid", fgColor=XL_LIGHT)
FILL_GREY = PatternFill("solid", fgColor=XL_GREY)
FILL_WARN = PatternFill("solid", fgColor=XL_WARN)
FILL_GREEN = PatternFill("solid", fgColor=XL_GREEN)

AL_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_R = Alignment(horizontal="right", vertical="center")


def init_sheet(ws, title, subtitle, widths):
    ws.sheet_view.showGridLines = False
    n = len(widths)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    c = ws.cell(row=1, column=1, value=title)
    c.font = H_TITLE
    c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30
    if subtitle:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
        c2 = ws.cell(row=2, column=1, value=subtitle)
        c2.font = H_SUB
        c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[2].height = 20
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def header_row(ws, row, headers):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = H_HEAD
        c.fill = FILL_HEAD
        c.alignment = AL_C
        c.border = BOX
    ws.row_dimensions[row].height = 30


def data_row(ws, row, values, fmts, aligns=None, fill=None, font=None):
    for i, v in enumerate(values, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = font or H_BODY
        c.border = BOX
        f = fmts[i - 1] if i - 1 < len(fmts) else None
        if f:
            c.number_format = f
        a = aligns[i - 1] if aligns and i - 1 < len(aligns) else None
        c.alignment = a or (AL_R if f in (MONEY, PCT) else AL_L)
        if fill:
            c.fill = fill


def total_row(ws, row, values, fmts, aligns=None):
    for i, v in enumerate(values, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=10, bold=True, color=XL_WHITE)
        c.fill = FILL_TOTAL
        c.border = TOP_LINE
        f = fmts[i - 1] if i - 1 < len(fmts) else None
        if f:
            c.number_format = f
        a = aligns[i - 1] if aligns and i - 1 < len(aligns) else None
        c.alignment = a or (AL_R if f in (MONEY, PCT) else AL_L)
    ws.row_dimensions[row].height = 24


def note(ws, row, ncols, text, height=34):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=9, color="FF595959")
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
    ws.row_dimensions[row].height = height


def section(ws, row, ncols, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = H_SECTION
    c.fill = FILL_GREY
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 24
    return row


# ============================================================ 练习册
OPEN_AR_NOTE = {
    "华文公司": "1 个月内，回款正常",
    "中兴办公": "1 个月内 45,000，1~3 个月 30,000",
    "海天文具": "1~3 个月 25,000，3~6 个月 10,000",
    "宏图商行": "含 6 个月以上 5,000",
    "联合采购": "3~6 个月 40,000，6 个月以上 25,000",
    "久远办公": "全部 6 个月以上，重点催收",
}


def build_practice():
    wb = Workbook()
    wb.properties.title = "晨光贸易2026年7月账务练习数据"

    # ---------- 01 说明与任务
    ws = wb.active
    ws.title = "01 说明与任务"
    init_sheet(ws, "晨光贸易有限公司 2026 年 7 月账务处理练习账套",
               "配套《WorkBuddy 在财务中的应用》第 4–10 章与第 18 章综合案例｜全部数据可勾稽，参考答案见同名《参考答案.xlsx》",
               [6, 26, 30, 22, 40])
    r = 3
    section(ws, r, 5, "一、案例背景"); r += 1
    note(ws, r, 5, "晨光贸易有限公司是一家虚构的一般纳税人商贸企业，主营办公用品批发，财务部只有会计小王和财务经理老李两个人。"
                   "2026 年 7 月，公司发生采购、销售、回款、付款、计提、结转等业务共 51 张记账凭证、25 笔银行收支。"
                   "本账套把书中第 4、5、6、7、8、9、10、18 章的表单与案例拆成一套可直接上手的完整月度账。", 52); r += 1
    note(ws, r, 5, "数据口径：期初余额取自书中表 4-1（合计 2,993,500.00 元，试算平衡），固定资产取表 6-2，"
                   "工资取表 6-3 的三人并扩展为 6 人。书中第 7–9 章为万元级示意数据（资产 2,850 万、月收入 520 万），"
                   "与本账套的元级口径相差 10 倍；本账套统一采用表 4-1 口径，因此月营业收入为 520,000 元、"
                   "营业成本 390,000 元、期间费用结构与书中一致，全部报表可逐项勾稽。", 62); r += 1

    r += 1
    section(ws, r, 5, "二、与书中表格的对应关系"); r += 1
    header_row(ws, r, ["序号", "书中位置", "本账套对应工作表", "用途", "说明"]); r += 1
    maps = [
        ("表 4-1 科目表节选（期初余额）", "03 期初余额表", "建账起点", "借贷合计各 2,993,500.00 元"),
        ("表 5-1 常见业务分录对照", "02 记账凭证（答案）", "科目映射口径", "运费、手续费、工资等 12 类业务"),
        ("表 6-1 应收账款账龄表", "04 期初应收账龄", "期初账龄", "合计 385,000.00 元"),
        ("表 6-2 固定资产折旧表", "08 固定资产卡片", "直线法折旧", "月折旧合计 7,283.33 元"),
        ("表 6-3 工资计算表", "09 工资表", "累计预扣法个税", "小王、小张数据与书中一致"),
        ("表 7-1 资产负债表", "07 资产负债表（答案）", "报表编制", "按元级口径重算"),
        ("第 8 章 利润表", "08 利润表（答案）", "报表编制", "毛利率 25%"),
        ("第 9 章 现金流量表", "09 现金流量表（答案）", "直接法+间接法", "经营净额 102,035.33 元"),
        ("第 18 章 六环节月结", "01 说明与任务（练习任务）", "全流程演练", "六个环节串成一条流水线"),
    ]
    for i, (a, b, c_, d) in enumerate(maps, 1):
        data_row(ws, r, [i, a, b, c_, d], [None, None, None, None, None],
                 [AL_C, AL_L, AL_L, AL_L, AL_L])
        ws.row_dimensions[r].height = 26
        r += 1

    r += 1
    section(ws, r, 5, "三、练习任务清单（建议按顺序完成）"); r += 1
    header_row(ws, r, ["序号", "练习任务", "对应章节", "交付成果", "检查要点"]); r += 1
    tasks = [
        ("建账：核对期初余额", "第 4 章", "03 期初余额表", "借方合计＝贷方合计＝2,993,500.00"),
        ("整理 30 张进项发票为台账", "第 5.3 节", "06 进项发票台账", "行数点名 30 行，金额×税率＝税额"),
        ("按银行流水批量生成凭证", "第 5.4 节", "02 记账凭证", "每张借贷相等，科目与摘要匹配"),
        ("凭证审核并出具意见清单", "第 5.5 节", "审核意见清单", "运费 2,000 元应记销售费用—运杂费"),
        ("应收账款账龄与催收清单", "第 6.3 节", "04 期末应收账龄", "四段账龄合计＝应收账款余额"),
        ("计提固定资产折旧", "第 6.4 节", "05 计提计算表", "月折旧合计 7,283.33 元"),
        ("工资与个税（累计预扣法）", "第 6.5 节", "09 工资表", "个税合计 132.00 元，实发 31,868.00 元"),
        ("银行存款余额调节表", "第 6.6 节", "06 银行余额调节表", "调节后双方相等"),
        ("月末结转（折旧/工资/税金/损益）", "第 10 章", "02 记账凭证", "结转后损益类科目余额全部为零"),
        ("编制三大报表", "第 7–9 章", "07/08/09 三张报表", "资产＝负债＋所有者权益"),
        ("三表勾稽检查", "第 7.5、9.5 节", "10 勾稽检查表", "六项勾稽全部通过"),
    ]
    for i, (a, b, c_, d) in enumerate(tasks, 1):
        data_row(ws, r, [i, a, b, c_, d], [None, None, None, None, None],
                 [AL_C, AL_L, AL_C, AL_L, AL_L])
        ws.row_dimensions[r].height = 26
        r += 1

    r += 1
    section(ws, r, 5, "四、本账套故意埋的 6 个坑（做完再对照）"); r += 1
    header_row(ws, r, ["序号", "陷阱", "正确做法", "涉及章节", "提示"]); r += 1
    traps = [
        ("运费发票 2,180 元（不含税 2,000、税额 180、税率 9%）", "记「销售费用—运杂费」，不是「管理费用—办公费」", "第 5.4/5.5 节", "书中第 7 号凭证的原型"),
        ("7 月 31 日开出转账支票 56,500 元付货款", "企业已记支出、银行未划款，属「企业已付银行未付」", "第 6.6 节", "未达账项之一"),
        ("7 月 31 日收到宏图货款 23,000 元", "企业已记收入、银行未入账，属「企业已收银行未收」", "第 6.6 节", "未达账项之二"),
        ("应付账款—中信包装 期末为借方 98,300 元", "重分类到「预付账款」，不能冲减应付账款", "第 7.3 节", "往来重分类"),
        ("6 月应付职工薪酬 45,000 元（工资 40,000 ＋ 单位社保 5,000）在 7 月结清", "工资部分代扣个人社保 8,000 元入其他应付款、个税 132 元入应交税费；单位部分尚未缴纳", "第 6.5 节", "两头都要记"),
        ("7 月 31 日购入货架 45,000 元（普通发票）", "进项税不可抵扣，全额计入固定资产；当月新增下月计提折旧", "第 6.4 节", "折旧起算时点"),
    ]
    for i, (a, b, c_, d) in enumerate(traps, 1):
        data_row(ws, r, [i, a, b, c_, d], [None, None, None, None, None],
                 [AL_C, AL_L, AL_L, AL_C, AL_L])
        ws.row_dimensions[r].height = 34
        r += 1

    r += 1
    section(ws, r, 5, "五、工作表清单"); r += 1
    header_row(ws, r, ["序号", "工作表", "内容", "数据性质", "备注"]); r += 1
    sheets = [
        ("02 会计科目表", "41 个末级科目及辅助核算建议", "基础资料", "按附录 C 扩展二级明细"),
        ("03 期初余额表", "17 个有余额科目 + 客户/供应商明细", "原始数据", "取自表 4-1"),
        ("04 期初应收账龄", "6 个客户四段账龄", "原始数据", "取自表 6-1"),
        ("05 销项发票台账", "6 张销项发票", "原始单据", "税率 13%"),
        ("06 进项发票台账", "30 张货物专票 + 1 张运费专票", "原始单据", "含校验列"),
        ("07 银行流水", "25 笔收支，含逐笔余额", "原始单据", "银行日记账口径"),
        ("08 固定资产卡片", "4 项资产 + 7 月新增 1 项", "原始数据", "直线法，残值率 5%"),
        ("09 工资表", "6 名员工，待算列留白", "原始数据+练习区", "累计预扣法"),
        ("10 银行对账单", "银行方流水", "原始单据", "与日记账差 33,500 元"),
    ]
    for i, (a, b, c_, d) in enumerate(sheets, 1):
        data_row(ws, r, [i, a, b, c_, d], [None, None, None, None, None],
                 [AL_C, AL_L, AL_L, AL_C, AL_L])
        ws.row_dimensions[r].height = 26
        r += 1

    r += 1
    note(ws, r, 5, "使用建议：先隐藏/不看参考答案，按「三、练习任务清单」逐项完成；每完成一项，用「10 勾稽检查表」自检。"
                   "所有金额单位为元，保留两位小数；日期均为 2026 年 7 月。", 34)
    ws.freeze_panes = "A3"

    # ---------- 02 会计科目表
    ws = wb.create_sheet("02 会计科目表")
    init_sheet(ws, "会计科目表", "按《企业会计准则》四级编码｜★ 标注为需要挂辅助核算的科目",
               [12, 34, 16, 8, 9, 16, 10, 34])
    header_row(ws, 3, ["科目编码", "科目名称", "一级科目", "类别", "余额方向", "辅助核算建议", "期初余额", "备注"])
    r = 4
    for code, name, top, cat, direction, aux in M.ACCOUNTS:
        d0, c0 = M.OPENING.get(code, (0, 0))
        has = "有" if (d0 or c0) else "—"
        remark = ""
        if code == "1002":
            remark = "日记账口径，与银行对账单逐笔勾兑"
        elif code == "1122":
            remark = "★ 按客户辅助核算，期末需做账龄"
        elif code == "2202":
            remark = "★ 按供应商辅助核算，借方余额需重分类"
        elif code == "1602":
            remark = "备抵科目，报表按净额列示"
        elif code.startswith("2221.01"):
            remark = "增值税三级明细，月末转出未交增值税"
        elif code == "4103":
            remark = "期末结转后余额转入利润分配"
        data_row(ws, r, [code, name, top, cat, direction, aux or "—", has, remark],
                 [None, None, None, None, None, None, None, None],
                 [AL_C, AL_L, AL_L, AL_C, AL_C, AL_L, AL_C, AL_L])
        r += 1
    note(ws, r + 1, 8, "说明：一级科目编码由财政部统一规定（1001 库存现金、1122 应收账款、2202 应付账款、6001 主营业务收入等），"
                       "二级、三级明细编码可按企业需要自行编制。损益类科目期末必须结转干净、余额为零。", 34)
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:H{r - 1}"

    # ---------- 03 期初余额表
    ws = wb.create_sheet("03 期初余额表")
    init_sheet(ws, "科目余额表（期初）", "2026 年 7 月 1 日｜单位：元｜取自《WorkBuddy 在财务中的应用》表 4-1",
               [14, 34, 12, 10, 16, 16])
    header_row(ws, 3, ["科目编码", "科目名称", "类别", "余额方向", "期初借方", "期初贷方"])
    r = 4
    db = cb = 0.0
    for code, name, top, cat, direction, aux in M.ACCOUNTS:
        d0, c0 = M.OPENING.get(code, (0, 0))
        if not (d0 or c0):
            continue
        data_row(ws, r, [code, name, cat, direction, d0 or None, c0 or None],
                 [None, None, None, None, MONEY, MONEY], [AL_C, AL_L, AL_C, AL_C, AL_R, AL_R])
        db += d0; cb += c0
        r += 1
    total_row(ws, r, ["合计", "", "", "", db, cb], [None, None, None, None, MONEY, MONEY],
              [AL_C, AL_L, AL_C, AL_C, AL_R, AL_R])
    r += 1
    note(ws, r, 6, f"试算平衡：借方合计 {db:,.2f} 元 ＝ 贷方合计 {cb:,.2f} 元。"
                   "资产类余额在借方，累计折旧作为固定资产的备抵科目余额在贷方；负债、权益类余额在贷方。", 30)
    r += 2

    section(ws, r, 6, "附：应收账款客户明细（辅助核算拆分，合计 385,000.00 元）"); r += 1
    header_row(ws, r, ["客户名称", "期初余额", "占比", "账龄特征"])
    r += 1
    for cust, amt in M.OPEN_AR.items():
        data_row(ws, r, [cust, amt, amt / 385000.0, OPEN_AR_NOTE.get(cust, "")],
                 [None, MONEY, PCT, None], [AL_L, AL_R, AL_R, AL_L])
        r += 1
    total_row(ws, r, ["合计", 385000.0, 1.0, ""],
              [None, MONEY, PCT, None], [AL_L, AL_R, AL_R, AL_L])
    r += 2

    section(ws, r, 6, "附：应付账款供应商明细（合计 268,000.00 元）"); r += 1
    header_row(ws, r, ["供应商名称", "期初余额", "账期", "备注"])
    r += 1
    for sup, amt in M.OPEN_AP.items():
        data_row(ws, r, [sup, amt, "30 天", "7 月已结清"],
                 [None, MONEY, None, None], [AL_L, AL_R, AL_C, AL_L])
        r += 1
    total_row(ws, r, ["合计", 268000.0, "", ""],
              [None, MONEY, None, None], [AL_L, AL_R, AL_C, AL_L])
    r += 2

    section(ws, r, 6, "附：应付职工薪酬构成（45,000.00 元）"); r += 1
    header_row(ws, r, ["明细", "金额", "说明"])
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    r += 1
    for a, b, c_ in [("6 月已计提未发放工资", 40000.0, "7 月 10 日发放，代扣个人社保公积金 8,000 元（含 1–6 月基数调整补差 4,000 元）、个税 132 元"),
                     ("6 月单位承担社保公积金", 5000.00, "6 月未缴，7 月仅缴纳个人部分；与 7 月单位部分一并于 8 月缴纳，期末仍挂账")]:
        data_row(ws, r, [a, b, c_], [None, MONEY, None], [AL_L, AL_R, AL_L])
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 30
        r += 1
    total_row(ws, r, ["合计", 45000.0, ""], [None, MONEY, None], [AL_L, AL_R, AL_L])
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    ws.freeze_panes = "A4"

    # ---------- 04 期初应收账龄
    ws = wb.create_sheet("04 期初应收账龄")
    init_sheet(ws, "应收账款账龄表（期初）", "2026 年 6 月 30 日｜单位：元｜取自表 6-1，合计 385,000.00 元",
               [18, 16, 16, 16, 16, 16, 12])
    header_row(ws, 3, ["客户", "1 个月以内", "1~3 个月", "3~6 个月", "6 个月以上", "合计", "占比"])
    r = 4
    tot = {b: 0.0 for b in M.BUCKETS}
    for cust, d in M.OPEN_AGING.items():
        vals = [d.get(b, 0.0) for b in M.BUCKETS]
        s = sum(vals)
        for b, v in zip(M.BUCKETS, vals):
            tot[b] += v
        data_row(ws, r, [cust] + vals + [s, s / 385000.0],
                 [None] + [MONEY] * 5 + [PCT],
                 [AL_L] + [AL_R] * 6)
        r += 1
    total_row(ws, r, ["合计"] + [tot[b] for b in M.BUCKETS] + [385000.0, 1.0],
              [None] + [MONEY] * 5 + [PCT], [AL_L] + [AL_R] * 6)
    r += 2
    note(ws, r, 7, "读表提示：久远办公 50,000 元全部落在「6 个月以上」，是重点催收对象；华文、中兴账龄集中在 1 个月以内，回款健康。"
                   "账龄合计必须与「03 期初余额表」的应收账款余额一致，这是账表相符的第一道关。", 40)
    ws.freeze_panes = "A4"

    # ---------- 05 销项发票台账
    ws = wb.create_sheet("05 销项发票台账")
    init_sheet(ws, "销项发票台账", "2026 年 7 月｜一般纳税人，货物适用 13% 税率｜单位：元",
               [6, 20, 12, 14, 40, 16, 8, 14, 16, 14, 20])
    header_row(ws, 3, ["序号", "发票号码", "开票日期", "购买方", "货物或服务名称",
                       "不含税金额", "税率", "税额", "价税合计", "结算方式", "备注"])
    r = 4
    for i, (d, cust, goods, amt) in enumerate(M.SALES, 1):
        data_row(ws, r, [i, f"0{i}4521{i:02d}", d, cust, goods, amt, 0.13,
                         round(amt * 0.13, 2), round(amt * 1.13, 2), "赊销",
                         "7 月 29 日已回款" if cust == "华文公司" else
                         ("7 月 30 日已回款" if cust == "中兴办公" else
                          ("7 月 31 日回款 23,000 元" if cust == "宏图商行" else "未回款"))],
                 [None, None, DATE_FMT, None, None, MONEY, PCT, MONEY, MONEY, None, None],
                 [AL_C, AL_C, AL_C, AL_L, AL_L, AL_R, AL_R, AL_R, AL_R, AL_C, AL_L])
        r += 1
    total_row(ws, r, ["合计", "", "", "", "", 520000.0, None, 67600.0, 587600.0, "", ""],
              [None, None, None, None, None, MONEY, None, MONEY, MONEY, None, None],
              [AL_C, AL_C, AL_C, AL_L, AL_L, AL_R, AL_R, AL_R, AL_R, AL_C, AL_L])
    r += 2
    note(ws, r, 11, "练习要求：核对每张发票「不含税金额 × 13% ＝ 税额」，确认价税合计正确；"
                    "再据此生成销售凭证（借：应收账款／贷：主营业务收入、应交税费—应交增值税（销项税额））。", 34)
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:K{r - 3}"

    # ---------- 06 进项发票台账
    ws = wb.create_sheet("06 进项发票台账")
    init_sheet(ws, "进项发票台账", "2026 年 7 月｜30 张货物专票 + 1 张运费专票｜单位：元",
               [6, 18, 14, 12, 26, 24, 16, 8, 14, 16, 10, 26])
    header_row(ws, 3, ["序号", "发票代码", "发票号码", "开票日期", "销售方名称", "货物或服务名称",
                       "不含税金额", "税率", "税额", "价税合计", "校验", "备注"])
    r = 4
    idx = 0
    for d, sup, amounts in M.PURCHASES:
        for amt in amounts:
            idx += 1
            tax = round(amt * 0.13, 2)
            ok = "通过" if abs(round(amt * 0.13, 2) - tax) < 0.005 else "异常"
            data_row(ws, r, [idx, "04400210", f"{31000000 + idx * 137}", d,
                             M.SUPPLIER_FULL[sup], "办公用品（A4 纸、笔、文件夹等）",
                             float(amt), 0.13, tax, round(amt + tax, 2), ok, ""],
                     [None, None, None, DATE_FMT, None, None, MONEY, PCT, MONEY, MONEY, None, None],
                     [AL_C, AL_C, AL_C, AL_C, AL_L, AL_L, AL_R, AL_R, AL_R, AL_R, AL_C, AL_L])
            r += 1
    idx += 1
    data_row(ws, r, [idx, "04400210", "31001988", M.FREIGHT["date"],
                     M.SUPPLIER_FULL[M.FREIGHT["supplier"]], "运输服务（销售商品运费）",
                     M.FREIGHT["amount"], 0.09, 180.00, 2180.00, "通过",
                     "税率 9%，注意不要按 13% 计算"],
             [None, None, None, DATE_FMT, None, None, MONEY, PCT, MONEY, MONEY, None, None],
             [AL_C, AL_C, AL_C, AL_C, AL_L, AL_L, AL_R, AL_R, AL_R, AL_R, AL_C, AL_L])
    r += 1
    total_row(ws, r, ["合计", "", "", "", "", "", 302000.0, None, 39180.0, 341180.0, "", ""],
              [None, None, None, None, None, None, MONEY, None, MONEY, MONEY, None, None],
              [AL_C, AL_C, AL_C, AL_C, AL_L, AL_L, AL_R, AL_R, AL_R, AL_R, AL_C, AL_L])
    r += 2
    note(ws, r, 12, "练习要求：① 行数点名——货物发票应为 30 行，少一行就说明漏识别；② 逐行核对「不含税金额 × 税率 ＝ 税额」；"
                    "③ 合计不含税 302,000 元、税额 39,180 元，其中运费 2,000 元与税额 180 元单独对应一张 9% 发票。", 40)
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:L{r - 3}"

    # ---------- 07 银行流水
    ws = wb.create_sheet("07 银行流水")
    init_sheet(ws, "银行存款日记账（银行流水）", "2026 年 7 月｜开户行：工商银行城东支行｜账号尾号 8866｜单位：元",
               [6, 12, 22, 44, 16, 16, 16, 14, 16])
    ws.cell(row=3, column=2, value="期初余额（2026-07-01）").font = H_BOLD
    ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=3)
    c = ws.cell(row=3, column=7, value=1250000.00)
    c.number_format = MONEY
    c.font = H_BOLD
    c.alignment = AL_R
    header_row(ws, 4, ["序号", "日期", "对方户名", "摘要", "收入金额", "支出金额", "余额", "单据类型", "单据号"])
    r = 5
    for i, row in enumerate(M.BANK_FLOW, 1):
        data_row(ws, r, [i, row["date"], row["party"], row["summary"],
                         row["inflow"] or None, row["outflow"] or None, row["bal"],
                         "银行回单" if row["inflow"] else "付款凭证", f"YH{20260700 + i}"],
                 [None, DATE_FMT, None, None, MONEY, MONEY, MONEY, None, None],
                 [AL_C, AL_C, AL_L, AL_L, AL_R, AL_R, AL_R, AL_C, AL_C])
        r += 1
    total_row(ws, r, ["合计", "", "", "", sum(x["inflow"] for x in M.BANK_FLOW),
                      sum(x["outflow"] for x in M.BANK_FLOW), M.BANK_FLOW[-1]["bal"], "", ""],
              [None, None, None, None, MONEY, MONEY, MONEY, None, None],
              [AL_C, AL_C, AL_L, AL_L, AL_R, AL_R, AL_R, AL_C, AL_C])
    r += 2
    note(ws, r, 9, "读表提示：余额列已按「上一行余额 ＋ 收入 － 支出」逐笔结出，可用它自检是否漏记。"
                   "流水是生成凭证最合适的入口——摘要 + 对方户名基本能定位科目，拿不准的应列入「待确认清单」而不是硬编。", 40)
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:I{r - 3}"

    # ---------- 08 固定资产卡片
    ws = wb.create_sheet("08 固定资产卡片")
    init_sheet(ws, "固定资产卡片及折旧参数", "直线法（平均年限法）｜净残值率 5%｜月折旧额 ＝ 原值 ×（1－净残值率）÷ 年限 ÷ 12",
               [10, 30, 12, 10, 12, 14, 10, 10, 14, 16, 16, 30])
    header_row(ws, 3, ["资产编号", "资产名称", "类别", "使用部门", "取得日期", "原值",
                       "净残值率", "使用年限", "月折旧额", "期初累计折旧", "期初净值", "备注"])
    r = 4
    tot_cost = tot_m = tot_acc = 0.0
    for no, nm, cat, dept, d0, cost, rate, years, acc in M.ASSETS:
        m = round(cost * (1 - rate) / years / 12, 2)
        data_row(ws, r, [no, nm, cat, dept, d0, cost, rate, years, m, acc,
                         round(cost - acc, 2), "在用"],
                 [None, None, None, None, DATE_FMT, MONEY, PCT, "0", MONEY, MONEY, MONEY, None],
                 [AL_C, AL_L, AL_C, AL_C, AL_C, AL_R, AL_R, AL_C, AL_R, AL_R, AL_R, AL_C])
        tot_cost += cost; tot_m += m; tot_acc += acc
        r += 1
    total_row(ws, r, ["合计", "", "", "", "", tot_cost, None, None, round(tot_m, 2),
                      round(tot_acc, 2), round(tot_cost - tot_acc, 2), ""],
              [None, None, None, None, None, MONEY, None, None, MONEY, MONEY, MONEY, None],
              [AL_C, AL_L, AL_C, AL_C, AL_C, AL_R, AL_R, AL_C, AL_R, AL_R, AL_R, AL_C])
    r += 2

    section(ws, r, 12, "7 月新增固定资产（当月新增，下月起计提折旧）"); r += 1
    header_row(ws, r, ["资产编号", "资产名称", "类别", "使用部门", "取得日期", "原值",
                       "净残值率", "使用年限", "本月计提", "进项税", "付款方式", "备注"])
    r += 1
    na = M.NEW_ASSET
    data_row(ws, r, [na["no"], na["name"], na["cat"], na["dept"], na["date"], na["cost"],
                     na["rate"], na["years"], 0.0, "不可抵扣", "银行转账",
                     "小规模纳税人开具增值税普通发票"],
             [None, None, None, None, DATE_FMT, MONEY, PCT, "0", MONEY, None, None, None],
             [AL_C, AL_L, AL_C, AL_C, AL_C, AL_R, AL_R, AL_C, AL_R, AL_C, AL_C, AL_L])
    r += 2
    note(ws, r, 12, "练习要求：① 按公式逐项计算月折旧额并核对合计 7,283.33 元；② 按使用部门区分折旧归属"
                    "（管理部记管理费用—折旧费，销售部记销售费用—折旧费）；③ 判断新增货架本月是否计提折旧。", 40)
    ws.freeze_panes = "A4"

    # ---------- 09 工资表
    ws = wb.create_sheet("09 工资表")
    init_sheet(ws, "2026 年 7 月工资表（练习区）", "个税按累计预扣法计算｜社保个人 10%、公积金个人 10%、单位社保公积金 20%（示例比例）｜单位：元",
               [8, 10, 12, 12, 12, 12, 12, 16, 18, 16, 14, 16])
    header_row(ws, 3, ["姓名", "部门", "岗位", "月工资", "社保个人", "公积金个人", "专项附加扣除",
                       "累计已预扣个税", "本月应纳税所得额", "本月应预扣个税", "实发工资", "单位社保公积金"])
    r = 4
    for p in M.PAYROLL:
        data_row(ws, r, [p["name"], p["dept"], p["post"], p["salary"], p["si"], p["hf"],
                         p["spec"], p["cum6"], None, None, None, p["unit"]],
                 [None, None, None, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY],
                 [AL_C, AL_C, AL_L, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R])
        for col in (9, 10, 11):
            ws.cell(row=r, column=col).fill = FILL_WARN
        r += 1
    total_row(ws, r, ["合计", "", "", M.pay_total,
                      4000.0, 4000.0, 4000.0, 792.0, None, None, None, 8000.0],
              [None, None, None, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY],
              [AL_C, AL_C, AL_L, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R])
    r += 2
    note(ws, r, 12, "黄底三列为本练习区，请自行填写：① 本月应纳税所得额 ＝ 月工资 － 5,000 － 社保个人 － 公积金个人 － 专项附加扣除；"
                    "② 累计应纳税所得额 ＝ 本月应纳税所得额 × 7（1–7 月），据此查综合所得税率表算累计应纳税额；"
                    "③ 本月应预扣 ＝ 累计应纳税额 － 累计已预扣；④ 实发工资 ＝ 月工资 － 社保个人 － 公积金个人 － 本月个税。", 52)
    note(ws, r + 1, 12, "提示：小王、小张的工资与个税结果与书中表 6-3 完全一致；老李月工资书中为 15,000 元，"
                        "本账套为便于部门费用分摊调整为 10,000 元。政策口径可能调整，实务中个税与社保比例须按当期最新政策复核。", 40)
    note(ws, r + 2, 12, "口径说明：表中「累计已预扣个税」为 1–6 月累计数，按 7 个月口径推算本月税额。6 月发放的工资代扣个人社保公积金合计 8,000 元，"
                        "其中含 1–6 月社保基数调整补差 4,000 元；7 月起按调整后基数每月代扣 4,000 元（个人 10%＋公积金 10%）、单位承担 8,000 元（20%）。"
                        "6 月单位承担的 5,000 元尚未缴纳，与 7 月的 8,000 元一并于 8 月缴纳。", 52)
    ws.freeze_panes = "A4"

    # ---------- 10 银行对账单
    ws = wb.create_sheet("10 银行对账单")
    init_sheet(ws, "银行对账单（银行方口径）", "2026 年 7 月｜开户行：工商银行城东支行｜账号尾号 8866｜单位：元",
               [6, 12, 12, 44, 16, 16, 16])
    ws.cell(row=3, column=1, value="期初余额（2026-07-01）").font = H_BOLD
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=4)
    c = ws.cell(row=3, column=7, value=1250000.00)
    c.number_format = MONEY
    c.font = H_BOLD
    c.alignment = AL_R
    header_row(ws, 4, ["序号", "记账日期", "交易日期", "摘要", "收入金额", "支出金额", "余额"])
    r = 5
    for i, row in enumerate(M.STMT, 1):
        data_row(ws, r, [i, row["date"], row["date"], row["summary"], row["inflow"] or None,
                         row["outflow"] or None, row["bal"]],
                 [None, DATE_FMT, DATE_FMT, None, MONEY, MONEY, MONEY],
                 [AL_C, AL_C, AL_C, AL_L, AL_R, AL_R, AL_R])
        r += 1
    total_row(ws, r, ["合计", "", "", "", sum(x["inflow"] for x in M.STMT),
                      sum(x["outflow"] for x in M.STMT), M.STMT[-1]["bal"]],
              [None, None, None, None, MONEY, MONEY, MONEY],
              [AL_C, AL_C, AL_C, AL_L, AL_R, AL_R, AL_R])
    r += 2
    note(ws, r, 7, f"对账单期末余额 {M.STMT[-1]['bal']:,.2f} 元，企业银行存款日记账期末余额 {M.BANK_FLOW[-1]['bal']:,.2f} 元，"
                   f"两者相差 {M.STMT[-1]['bal'] - M.BANK_FLOW[-1]['bal']:,.2f} 元。请逐笔勾兑，找出未达账项并编制余额调节表。", 40)
    ws.freeze_panes = "A5"

    wb.save(F_PRACTICE)
    return F_PRACTICE


# ============================================================ 参考答案
def build_answer():
    wb = Workbook()
    wb.properties.title = "晨光贸易2026年7月账务参考答案"

    # ---------- 01 答案说明
    ws = wb.active
    ws.title = "01 答案说明"
    init_sheet(ws, "参考答案 · 使用说明", "本册对应《晨光贸易2026年7月_账务练习数据.xlsx》，请先完成练习再看答案",
               [6, 30, 22, 46, 20])
    r = 3
    section(ws, r, 5, "一、答案要点"); r += 1
    note(ws, r, 5, "本册共 10 张表：记账凭证（51 张）、科目余额表、期末账龄、计提计算表、银行余额调节表、"
                   "资产负债表、利润表、现金流量表、勾稽检查表。所有数字由同一套凭证推导，"
                   "凭证借贷平衡 → 科目余额表试算平衡 → 三表勾稽通过，可逐级验算。", 44); r += 1
    r += 1
    section(ws, r, 5, "二、关键结果速览"); r += 1
    header_row(ws, r, ["序号", "指标", "金额（元）", "说明", "所在表"]); r += 1
    kpis = [
        ("本月营业收入", 520000.0, "6 张销项发票不含税合计", "08 利润表"),
        ("本月营业成本", 390000.0, "毛利率 25%", "08 利润表"),
        ("营业利润", M.operating, "收入－成本－税金及附加－三费", "08 利润表"),
        ("净利润", M.net_profit, "所得税按 25% 测算", "08 利润表"),
        ("资产总计", M.total_assets, "含往来重分类后的预付账款", "07 资产负债表"),
        ("负债和所有者权益总计", M.total_le, "与资产总计相等", "07 资产负债表"),
        ("经营活动现金净额", M.cf_oper, "直接法与间接法一致", "09 现金流量表"),
        ("期末货币资金", M.cash, "现金 8,500 ＋ 银行存款", "07 资产负债表"),
        ("应交增值税", M.vat_payable, "销项 67,600 － 进项 39,180", "05 计提计算表"),
        ("月折旧额", M.dep_total, "4 项资产直线法合计", "05 计提计算表"),
        ("本月代扣个税", M.pay_tax, "6 人累计预扣法合计", "05 计提计算表"),
    ]
    for i, (a, b, c_, d) in enumerate(kpis, 1):
        data_row(ws, r, [i, a, b, c_, d], [None, None, MONEY, None, None],
                 [AL_C, AL_L, AL_R, AL_L, AL_C])
        r += 1
    r += 1
    section(ws, r, 5, "三、六个易错点及正确处理"); r += 1
    header_row(ws, r, ["序号", "易错点", "正确做法", "答案位置", "涉及金额"]); r += 1
    errs = [
        ("运费记入管理费用", "销售环节运费记「销售费用—运杂费」，进项税 180 元仍可抵扣", "02 记账凭证 记-008", 2000.0),
        ("7/31 支票付款", "企业已付、银行未付，编制调节表时加回银行对账单", "06 银行余额调节表", 56500.0),
        ("7/31 收款 23,000 元", "企业已收、银行未收，调节时从银行对账单扣除", "06 银行余额调节表", 23000.0),
        ("应付账款出现借方余额", "中信包装借方 98,300 元重分类到预付账款，应付账款只列贷方明细 180,800 元", "07 资产负债表", 98300.0),
        ("6 月应付职工薪酬 7 月结清", "工资 40,000 元代扣个人社保 8,000 元入其他应付款、个税 132 元入应交税费，不能全额冲减银行存款；单位社保 5,000 元尚未缴纳", "02 记账凭证 记-012", 45000.0),
        ("新增资产当月计提折旧", "7/31 购入货架本月不计提，8 月起计提", "05 计提计算表", 45000.0),
    ]
    for i, (a, b, c_, d) in enumerate(errs, 1):
        data_row(ws, r, [i, a, b, c_, d], [None, None, None, None, MONEY],
                 [AL_C, AL_L, AL_L, AL_C, AL_R])
        ws.row_dimensions[r].height = 40
        r += 1
    r += 1
    note(ws, r, 5, "说明：书中第 7–9 章为万元级示意（月收入 520 万），本账套按表 4-1 的元级口径重算为 520,000 元，"
                   "结构与比例保持一致，便于对照学习。", 34)
    ws.freeze_panes = "A3"

    # ---------- 02 记账凭证
    ws = wb.create_sheet("02 记账凭证")
    init_sheet(ws, "记账凭证（2026 年 7 月，共 51 张）", "每张凭证借贷必相等｜现金流量项目仅标注在银行存款行｜单位：元",
               [6, 12, 10, 40, 12, 30, 16, 16, 16, 30, 18, 10])
    header_row(ws, 3, ["序号", "日期", "凭证字号", "摘要", "科目编码", "科目名称",
                       "辅助核算", "借方金额", "贷方金额", "现金流量项目", "原始单据", "类型"])
    r = 4
    n = 0
    for vi, v in enumerate(M.VOUCHERS):
        fill = FILL_GREY if vi % 2 else None
        for li, (code, d, c, aux, cf) in enumerate(v["lines"]):
            n += 1
            data_row(ws, r, [n, v["date"], v["no"], v["summary"] if li == 0 else "",
                             code, M.NAME.get(code, code), aux or "", d or None, c or None,
                             cf or "", v["source"] if li == 0 else "",
                             v["vtype"] if li == 0 else ""],
                     [None, DATE_FMT, None, None, None, None, None, MONEY, MONEY, None, None, None],
                     [AL_C, AL_C, AL_C, AL_L, AL_C, AL_L, AL_L, AL_R, AL_R, AL_L, AL_L, AL_C],
                     fill=fill)
            r += 1
    total_row(ws, r, ["合计", "", "", "", "", "", "", M.TOTAL_DEBIT, M.TOTAL_DEBIT, "", "", ""],
              [None, None, None, None, None, None, None, MONEY, MONEY, None, None, None],
              [AL_C, AL_C, AL_C, AL_L, AL_C, AL_L, AL_L, AL_R, AL_R, AL_L, AL_L, AL_C])
    r += 2
    note(ws, r, 12, "口径说明：① 采购按「同一天 + 同一供应商」合并开票，30 张专票合并为 9 张凭证；② 月末结转单独标记为「期末结转」，"
                    "不计入利润表发生额；③ 结转损益后全部损益类科目余额为零。", 34)
    ws.freeze_panes = "D4"
    ws.auto_filter.ref = f"A3:L{r - 3}"

    # ---------- 03 科目余额表
    ws = wb.create_sheet("03 科目余额表")
    init_sheet(ws, "科目余额表（2026 年 7 月）", "期初 ＋ 本期发生 ＝ 期末｜期末余额按科目性质分列借贷｜单位：元",
               [14, 32, 14, 8, 10, 14, 14, 14, 14, 14, 14])
    header_row(ws, 3, ["科目编码", "科目名称", "一级科目", "类别", "余额方向",
                       "期初借方", "期初贷方", "本期借方", "本期贷方", "期末借方", "期末贷方"])
    r = 4
    tb = [0.0] * 6
    for code, name, top, cat, direction, aux in M.ACCOUNTS:
        d0, c0 = M.OPENING.get(code, (0, 0))
        d1, c1 = M.MOV[code]
        cl = M.closing(code)
        cd, cc = max(cl, 0), max(-cl, 0)
        data_row(ws, r, [code, name, top, cat, direction, d0 or None, c0 or None,
                         d1 or None, c1 or None, cd or None, cc or None],
                 [None, None, None, None, None, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY],
                 [AL_C, AL_L, AL_L, AL_C, AL_C, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R])
        for i, v in enumerate([d0, c0, d1, c1, cd, cc]):
            tb[i] += v
        r += 1
    total_row(ws, r, ["合计", "", "", "", ""] + tb,
              [None, None, None, None, None] + [MONEY] * 6,
              [AL_C, AL_L, AL_L, AL_C, AL_C] + [AL_R] * 6)
    r += 2
    section(ws, r, 11, "试算平衡检查"); r += 1
    header_row(ws, r, ["检查项", "借方合计", "贷方合计", "差异", "结论", "", "", "", "", "", ""])
    r += 1
    checks = [
        ("期初余额平衡", tb[0], tb[1]),
        ("本期发生额平衡", tb[2], tb[3]),
        ("期末余额平衡", tb[4], tb[5]),
    ]
    for name, a, b in checks:
        data_row(ws, r, [name, a, b, round(a - b, 2), "通过" if abs(a - b) < 0.005 else "不通过",
                         "", "", "", "", "", ""],
                 [None, MONEY, MONEY, MONEY, None, None, None, None, None, None, None],
                 [AL_L, AL_R, AL_R, AL_R, AL_C, AL_L, AL_L, AL_L, AL_L, AL_L, AL_L])
        ws.cell(row=r, column=5).fill = FILL_GREEN
        r += 1
    r += 1
    note(ws, r, 11, "看表提示：损益类科目（6001、6401、6403、6601、6602、6603、6801）期末借贷均为零，说明已结转干净；"
                    "「本年利润」4103 也已结转到利润分配，余额为零。", 34)
    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A3:K{r - 6}"

    # ---------- 04 期末账龄
    ws = wb.create_sheet("04 期末应收账龄")
    init_sheet(ws, "应收账款账龄表（期末）", "2026 年 7 月 31 日｜单位：元｜合计应与应收账款余额 253,000.00 元一致",
               [18, 16, 16, 16, 16, 16, 12])
    header_row(ws, 3, ["客户", "1 个月以内", "1~3 个月", "3~6 个月", "6 个月以上", "合计", "占比"])
    r = 4
    tot = {b: 0.0 for b in M.BUCKETS}
    for cust, d in M.CLOSE_AGING.items():
        vals = [d.get(b, 0.0) for b in M.BUCKETS]
        s = sum(vals)
        for b, v in zip(M.BUCKETS, vals):
            tot[b] += v
        data_row(ws, r, [cust] + vals + [s, s / M.ar],
                 [None] + [MONEY] * 5 + [PCT], [AL_L] + [AL_R] * 6)
        r += 1
    total_row(ws, r, ["合计"] + [tot[b] for b in M.BUCKETS] + [M.ar, 1.0],
              [None] + [MONEY] * 5 + [PCT], [AL_L] + [AL_R] * 6)
    r += 2
    section(ws, r, 7, "账龄变动说明（期初 385,000.00 → 期末 253,000.00）"); r += 1
    header_row(ws, r, ["客户", "期初余额", "本期销售", "本期回款", "期末余额", "变化原因", ""])
    r += 1
    move = [
        ("华文公司", 120000.0, 226000.0, 346000.0, 0.0, "期初款与 7/5 销售款均已收回"),
        ("中兴办公", 75000.0, 135600.0, 210600.0, 0.0, "期初款与 7/8 销售款均已收回"),
        ("海天文具", 35000.0, 90400.0, 35000.0, 90400.0, "期初款收回，新增 7/12 销售未收"),
        ("宏图商行", 40000.0, 67800.0, 63000.0, 44800.0, "期初款收回，7/18 销售仅回 23,000 元"),
        ("联合采购", 65000.0, 45200.0, 65000.0, 45200.0, "期初款收回，新增 7/24 销售未收"),
        ("久远办公", 50000.0, 0.0, 0.0, 50000.0, "全部在 6 个月以上，需重点催收"),
        ("启明学校", 0.0, 22600.0, 0.0, 22600.0, "新客户，7/28 销售未收"),
    ]
    for row in move:
        data_row(ws, r, list(row) + [""], [None, MONEY, MONEY, MONEY, MONEY, None, None],
                 [AL_L, AL_R, AL_R, AL_R, AL_R, AL_L, AL_L])
        ws.row_dimensions[r].height = 30
        r += 1
    total_row(ws, r, ["合计", 385000.0, 587600.0, 719600.0, M.ar, "", ""],
              [None, MONEY, MONEY, MONEY, MONEY, None, None],
              [AL_L, AL_R, AL_R, AL_R, AL_R, AL_L, AL_L])
    ws.freeze_panes = "A4"

    # ---------- 05 计提计算表
    ws = wb.create_sheet("05 计提计算表")
    init_sheet(ws, "月末计提计算表", "固定资产折旧 / 工资与个税 / 税金及附加｜单位：元",
               [8, 30, 12, 14, 10, 10, 14, 14, 20, 16, 14, 14, 14])
    r = 3
    section(ws, r, 13, "一、固定资产折旧（直线法，残值率 5%）"); r += 1
    header_row(ws, r, ["编号", "资产名称", "使用部门", "原值", "残值率", "年限",
                       "月折旧额", "本月计提", "费用归属"])
    r += 1
    for no, nm, dept, cost, rate, years, m, acc in M.DEP:
        data_row(ws, r, [no, nm, dept, cost, rate, years, m, m,
                         "管理费用—折旧费" if dept == "管理部" else "销售费用—折旧费"],
                 [None, None, None, MONEY, PCT, "0", MONEY, MONEY, None],
                 [AL_C, AL_L, AL_C, AL_R, AL_R, AL_C, AL_R, AL_R, AL_L])
        r += 1
    data_row(ws, r, [M.NEW_ASSET["no"], M.NEW_ASSET["name"], M.NEW_ASSET["dept"],
                     M.NEW_ASSET["cost"], M.NEW_ASSET["rate"], M.NEW_ASSET["years"],
                     0.0, 0.0, "本月不计提（下月起）"],
             [None, None, None, MONEY, PCT, "0", MONEY, MONEY, None],
             [AL_C, AL_L, AL_C, AL_R, AL_R, AL_C, AL_R, AL_R, AL_L])
    r += 1
    total_row(ws, r, ["合计", "", "", 495000.0, None, None, M.dep_total, M.dep_total, ""],
              [None, None, None, MONEY, None, None, MONEY, MONEY, None],
              [AL_C, AL_L, AL_C, AL_R, AL_R, AL_C, AL_R, AL_R, AL_L])
    r += 1
    note(ws, r, 13, f"计提分录：借 管理费用—折旧费 {M.dep_admin:,.2f}、销售费用—折旧费 {M.dep_sales:,.2f}；"
                    f"贷 累计折旧 {M.dep_total:,.2f}。", 24)
    r += 2

    section(ws, r, 13, "二、工资与个人所得税（累计预扣法，1–7 月）"); r += 1
    header_row(ws, r, ["姓名", "部门", "月工资", "社保个人", "公积金个人", "专项附加扣除",
                       "月应纳税所得额", "累计已预扣(1–6月)", "累计应纳税所得额(1–7月)",
                       "本月应预扣个税", "实发工资", "单位社保公积金", "费用归属"])
    r += 1
    for p in M.PAYROLL:
        data_row(ws, r, [p["name"], p["dept"], p["salary"], p["si"], p["hf"], p["spec"],
                         p["taxable_m"], p["cum6"], round(max(p["taxable_m"], 0) * 7, 2),
                         p["tax"], p["net"], p["unit"],
                         "管理费用" if p["dept"] == "管理部" else "销售费用"],
                 [None, None, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, None],
                 [AL_C, AL_C, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_L])
        r += 1
    total_row(ws, r, ["合计", "", M.pay_total, 4000.0, 4000.0, 4000.0, 4400.0, 792.0, 30800.0,
                      M.pay_tax, M.pay_net, M.pay_unit, ""],
              [None, None, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, None],
              [AL_C, AL_C, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_R, AL_L])
    r += 1
    note(ws, r, 13, "计算过程（以小王为例）：月应纳税所得额 ＝ 8,000 － 5,000 － 800 － 800 － 1,000 ＝ 400；"
                    "1–6 月累计 2,400 元，税额 72 元；1–7 月累计 2,800 元，税额 84 元；本月应预扣 84 － 72 ＝ 12 元。"
                    f"六人合计本月个税 {M.pay_tax:,.2f} 元，实发工资合计 {M.pay_net:,.2f} 元。", 44)
    r += 2

    section(ws, r, 13, "三、税金及附加与增值税"); r += 1
    header_row(ws, r, ["项目", "计税依据", "计税金额", "税率", "金额", "说明"])
    r += 1
    tax_rows = [
        ("销项税额", "不含税销售额", 520000.0, 0.13, M.out_vat, "6 张销项发票"),
        ("进项税额", "不含税采购额 + 运费", 302000.0, None, M.in_vat, "货物 39,000 ＋ 运费 180"),
        ("应交增值税", "销项 － 进项", None, None, M.vat_payable, "转出未交增值税"),
        ("城市维护建设税", "应交增值税", M.vat_payable, 0.07, M.city, "市区 7%"),
        ("教育费附加", "应交增值税", M.vat_payable, 0.03, M.edu, "3%"),
        ("地方教育附加", "应交增值税", M.vat_payable, 0.02, M.edu2, "2%"),
        ("印花税", "购销合同金额", 822000.0, 0.0003, M.stamp, "采购 30 万 ＋ 销售 52 万 ＋ 运费 0.2 万"),
        ("税金及附加合计", "", None, None, M.tax_surcharge, "附加税 ＋ 印花税"),
        ("利润总额", "收入－成本－税金及附加－三费", None, None, M.operating, "所得税计税基础"),
        ("企业所得税", "利润总额", M.operating, 0.25, M.income_tax, "按 25% 测算"),
    ]
    for a, b, c_, d, e, f in tax_rows:
        data_row(ws, r, [a, b, c_, d, e, f],
                 [None, None, MONEY, PCT, MONEY, None],
                 [AL_L, AL_L, AL_R, AL_R, AL_R, AL_L])
        if a in ("应交增值税", "税金及附加合计", "企业所得税"):
            for col in range(1, 7):
                ws.cell(row=r, column=col).font = H_BOLD
        r += 1
    r += 1
    note(ws, r, 13, "注意：设备采购 45,000 元取得增值税普通发票，进项税不得抵扣；且未签订书面买卖合同，不计缴印花税。"
                    "房租合同印花税已于签约时缴纳，本月不再计提。", 34)
    ws.freeze_panes = "A4"

    # ---------- 06 银行余额调节表
    ws = wb.create_sheet("06 银行余额调节表")
    init_sheet(ws, "银行存款余额调节表", "2026 年 7 月 31 日｜单位：元｜编制人：小王｜复核人：老李",
               [10, 46, 18, 44])
    header_row(ws, 3, ["行次", "项目", "金额", "说明"])
    r = 4
    rows = [
        ("1", "银行对账单余额", M.stmt_close, "见「10 银行对账单」"),
        ("2", "加：企业已收、银行未收", 23000.0, "7/31 宏图商行货款，银行尚未入账"),
        ("3", "减：企业已付、银行未付", 56500.0, "7/31 开出转账支票 0021875，银行尚未划款"),
        ("4", "调节后余额（银行方）", M.bank_close, "1 ＋ 2 － 3"),
        ("5", "企业银行存款日记账余额", M.bank_close, "见「07 银行流水」期末余额"),
        ("6", "加：银行已收、企业未收", 0.0, "本月无"),
        ("7", "减：银行已付、企业未付", 0.0, "本月无"),
        ("8", "调节后余额（企业方）", M.bank_close, "5 ＋ 6 － 7"),
        ("9", "差异", 0.0, "4 － 8，为零即调节相符"),
    ]
    for a, b, c_, d in rows:
        data_row(ws, r, [a, b, c_, d], [None, None, MONEY, None], [AL_C, AL_L, AL_R, AL_L])
        if a in ("4", "8", "9"):
            for col in range(1, 5):
                ws.cell(row=r, column=col).font = H_BOLD
        if a == "9":
            for col in range(1, 5):
                ws.cell(row=r, column=col).fill = FILL_GREEN
        r += 1
    r += 2
    note(ws, r, 4, "结论：调节后银行方与企业方余额均为 2,302,835.33 元，账实相符。两笔未达账项均为 7 月 31 日发生的时间差，"
                   "8 月银行入账、企业补记后自然消失，无需调账。", 34)
    r += 1
    note(ws, r, 4, "提醒：若某笔未达账项连续数月挂账，须追查票据是否有效、是否存在漏记费用；建议每月对账一次，不要攒到年底。", 30)
    ws.freeze_panes = "A4"

    # ---------- 07 资产负债表
    ws = wb.create_sheet("07 资产负债表")
    init_sheet(ws, "资产负债表", "2026 年 7 月 31 日｜单位：元｜往来科目已按明细重分类",
               [8, 30, 18, 18, 34])
    header_row(ws, 3, ["行次", "项目", "期末余额", "期初余额", "编制说明"])
    r = 4
    bs = [
        ("", "资产", None, None, ""),
        ("1", "　货币资金", M.cash, 1258500.0, "库存现金 ＋ 银行存款"),
        ("2", "　应收账款", M.ar, 385000.0, "按客户明细借方余额列示"),
        ("3", "　预付账款", M.prepaid, 120000.0, "预付账款 ＋ 应付账款借方明细 98,300（中信包装）"),
        ("4", "　存货", M.inventory, 780000.0, "库存商品"),
        ("5", "　流动资产合计", M.cash + M.ar + M.prepaid + M.inventory, 2543500.0, "1＋2＋3＋4"),
        ("6", "　固定资产原值", M.fa_cost, 450000.0, "含 7 月新增货架 45,000"),
        ("7", "　减：累计折旧", M.acc_dep, 96000.0, "期初 96,000 ＋ 本月计提 7,283.33"),
        ("8", "　固定资产净值", M.fa_net, 354000.0, "6－7，按净额列示"),
        ("9", "　非流动资产合计", M.fa_net, 354000.0, ""),
        ("10", "资产总计", M.total_assets, 2897500.0, "5＋9"),
        ("", "", None, None, ""),
        ("", "负债和所有者权益", None, None, ""),
        ("11", "　应付账款", M.ap, 268000.0, "仅列贷方明细：华信纸品 79,100 ＋ 恒达文具 101,700"),
        ("12", "　应付职工薪酬", M.payroll_pay, 45000.0, "7 月计提工资 40,000 ＋ 单位社保 8,000 ＋ 6 月社保 5,000"),
        ("13", "　应交税费", M.tax_pay, 32500.0, "未交增值税、附加税、企业所得税、个税"),
        ("14", "　流动负债合计", M.ap + M.payroll_pay + M.tax_pay, 345500.0, "11＋12＋13"),
        ("15", "　短期借款", M.short_loan, 0.0, "7/1 借入 1,000,000，期限 6 个月"),
        ("16", "负债合计", M.total_liab, 345500.0, "14＋15"),
        ("17", "　实收资本", M.capital, 2000000.0, ""),
        ("18", "　未分配利润", M.retained, 552000.0, "期初 552,000 ＋ 本月净利润 34,757.25"),
        ("19", "所有者权益合计", M.total_equity, 2552000.0, "17＋18"),
        ("20", "负债和所有者权益总计", M.total_le, 2897500.0, "16＋19"),
        ("21", "差额（资产－负债和权益）", round(M.total_assets - M.total_le, 2), 0.0, "必须为零"),
    ]
    for a, b, c_, d, e in bs:
        if b in ("资产", "负债和所有者权益"):
            section(ws, r, 5, b)
            r += 1
            continue
        if b == "":
            r += 1
            continue
        bold = b in ("资产总计", "负债合计", "所有者权益合计", "负债和所有者权益总计",
                     "流动资产合计", "非流动资产合计", "流动负债合计")
        data_row(ws, r, [a, b, c_, d, e], [None, None, MONEY, MONEY, None],
                 [AL_C, AL_L, AL_R, AL_R, AL_L],
                 font=H_BOLD if bold else None,
                 fill=FILL_GREEN if b == "差额（资产－负债和权益）" else None)
        r += 1
    r += 1
    note(ws, r, 5, "取数规则回顾：① 直接取数（短期借款、实收资本）；② 合并/备抵取数（货币资金＝库存现金＋银行存款；"
                   "固定资产＝原值－累计折旧）；③ 往来重分类（应付账款—中信包装 借方 98,300 元搬入预付账款）。", 40)
    ws.freeze_panes = "A4"

    # ---------- 08 利润表
    ws = wb.create_sheet("08 利润表")
    init_sheet(ws, "利润表", "2026 年 7 月｜单位：元｜多步式，逐层剥茧",
               [8, 30, 18, 14, 36])
    header_row(ws, 3, ["行次", "项目", "本期金额", "占收入比", "计算说明"])
    r = 4
    pl = [
        ("1", "一、营业收入", M.revenue, "主营业务收入 520,000"),
        ("2", "　减：营业成本", M.cost, "结转已售商品成本 390,000"),
        ("3", "　　　税金及附加", M.tax_add, "附加税 3,410.40 ＋ 印花税 246.60"),
        ("4", "　　　销售费用", M.sell_exp, "职工薪酬 26,400 ＋ 折旧 3,325 ＋ 运杂 2,000 ＋ 差旅 5,000 ＋ 宣传 8,275"),
        ("5", "　　　管理费用", M.admin_exp, "职工薪酬 21,600 ＋ 折旧 3,958.33 ＋ 房租 3,000 ＋ 办公 1,441.67"),
        ("6", "　　　财务费用", M.fin_exp, "手续费 2,000 ＋ 利息支出 4,200 － 利息收入 1,200"),
        ("7", "二、营业利润", M.operating, "1－2－3－4－5－6"),
        ("8", "　加：营业外收入", 0.0, "本月无"),
        ("9", "　减：营业外支出", 0.0, "本月无"),
        ("10", "三、利润总额", M.operating, "7＋8－9"),
        ("11", "　减：所得税费用", M.income_tax, "利润总额 × 25%"),
        ("12", "四、净利润", M.net_profit, "10－11"),
    ]
    for a, b, c_, d in pl:
        bold = b.startswith(("一、", "二、", "三、", "四、"))
        data_row(ws, r, [a, b, c_, c_ / M.revenue, d],
                 [None, None, MONEY, PCT, None], [AL_C, AL_L, AL_R, AL_R, AL_L],
                 font=H_BOLD if bold else None,
                 fill=FILL_GREEN if b == "四、净利润" else None)
        r += 1
    r += 2
    section(ws, r, 5, "关键比率"); r += 1
    header_row(ws, r, ["行次", "指标", "数值", "计算", "解读"])
    r += 1
    ratios = [
        ("13", "毛利率", (M.revenue - M.cost) / M.revenue, "（收入－成本）÷ 收入", "与书中 25% 一致，属商贸企业正常水平"),
        ("14", "营业利润率", M.operating / M.revenue, "营业利润 ÷ 收入", "每 100 元收入留下 8.91 元经营利润"),
        ("15", "净利率", M.net_profit / M.revenue, "净利润 ÷ 收入", "每 100 元收入净赚 6.68 元"),
        ("16", "期间费用率", (M.sell_exp + M.admin_exp + M.fin_exp) / M.revenue,
         "三费 ÷ 收入", "费用控制尚可，销售费用占比最高"),
    ]
    for a, b, c_, d, e in ratios:
        data_row(ws, r, [a, b, c_, d, e], [None, None, PCT, None, None],
                 [AL_C, AL_L, AL_R, AL_L, AL_L])
        r += 1
    r += 1
    note(ws, r, 5, "勾稽提示：净利润 34,757.25 元应等于资产负债表「未分配利润」期末 586,757.25 元减去期初 552,000.00 元。", 28)
    ws.freeze_panes = "A4"

    # ---------- 09 现金流量表
    ws = wb.create_sheet("09 现金流量表")
    init_sheet(ws, "现金流量表", "2026 年 7 月｜单位：元｜直接法编制，间接法验证",
               [8, 40, 18, 34])
    header_row(ws, 3, ["行次", "项目", "本期金额", "取数说明"])
    r = 4
    cf = [
        ("", "一、经营活动产生的现金流量", None, ""),
        ("1", "　销售商品、提供劳务收到的现金", M.cf_sales, "回款 8 笔，合计 719,600"),
        ("2", "　收到其他与经营活动有关的现金", M.cf_other_in, "银行存款利息 1,200"),
        ("3", "　经营活动现金流入小计", M.cf_oper_in, "1＋2"),
        ("4", "　购买商品、接受劳务支付的现金", M.cf_buy, "付货款 524,500 ＋ 运费 2,180"),
        ("5", "　支付给职工以及为职工支付的现金", M.cf_staff, "发工资 31,868 ＋ 缴社保 8,000"),
        ("6", "　支付的各项税费", M.cf_tax, "缴纳 6 月税费 32,500"),
        ("7", "　支付其他与经营活动有关的现金", M.cf_other_out, "房租 3,000 ＋ 办公 1,441.67 ＋ 差旅 5,000 ＋ 宣传 8,275 ＋ 手续费 2,000"),
        ("8", "　经营活动现金流出小计", M.cf_oper_out, "4＋5＋6＋7"),
        ("9", "　经营活动产生的现金流量净额", M.cf_oper, "3－8"),
        ("", "", None, ""),
        ("", "二、投资活动产生的现金流量", None, ""),
        ("10", "　购建固定资产、无形资产和其他长期资产支付的现金", M.cf_inv_out, "购入仓储货架 45,000"),
        ("11", "　投资活动产生的现金流量净额", M.cf_inv, "－10"),
        ("", "", None, ""),
        ("", "三、筹资活动产生的现金流量", None, ""),
        ("12", "　取得借款收到的现金", M.cf_fin_in, "短期借款 1,000,000"),
        ("13", "　分配股利、利润或偿付利息支付的现金", M.cf_fin_out, "支付借款利息 4,200"),
        ("14", "　筹资活动产生的现金流量净额", M.cf_fin, "12－13"),
        ("", "", None, ""),
        ("15", "四、现金及现金等价物净增加额", M.cf_net, "9＋11＋14"),
        ("16", "　加：期初现金及现金等价物余额", M.cash_begin, "库存现金 8,500 ＋ 银行存款 1,250,000"),
        ("17", "五、期末现金及现金等价物余额", M.cash_end, "15＋16"),
    ]
    for a, b, c_, d in cf:
        if b.startswith(("一、", "二、", "三、")):
            section(ws, r, 4, b)
            r += 1
            continue
        if b == "":
            r += 1
            continue
        bold = b.startswith(("经营活动产生的", "投资活动产生的", "筹资活动产生的", "四、", "五、"))
        data_row(ws, r, [a, b, c_, d], [None, None, MONEY, None], [AL_C, AL_L, AL_R, AL_L],
                 font=H_BOLD if bold else None,
                 fill=FILL_GREEN if b.startswith("五、") else None)
        r += 1
    r += 2

    section(ws, r, 4, "补充资料：将净利润调节为经营活动现金流量（间接法）"); r += 1
    header_row(ws, r, ["行次", "项目", "本期金额", "说明"])
    r += 1
    ind = [
        ("18", "净利润", M.net_profit, "见利润表"),
        ("19", "加：固定资产折旧", M.dep_total, "非付现费用"),
        ("20", "　　财务费用中的利息支出", 4200.0, "属筹资活动，需加回"),
        ("21", "减：经营性应收项目的增加（增加以负号列示）", -M.d_ar, "应收账款减少 132,000，故为加项"),
        ("22", "减：存货的增加", -M.d_inv, "存货减少 90,000，故为加项"),
        ("23", "减：经营性应付项目的减少", M.d_ap, "应付账款减少 185,500"),
        ("24", "加：应付职工薪酬的增加", M.d_pay, "53,000 － 45,000"),
        ("25", "加：应交税费的增加", M.d_tax, "43,794.75 － 32,500"),
        ("26", "经营活动产生的现金流量净额（间接法）", M.indirect, "18 至 25 合计"),
        ("27", "直接法结果", M.cf_oper, "见上表第 9 行"),
        ("28", "差异", round(M.indirect - M.cf_oper, 2), "必须为零"),
    ]
    for a, b, c_, d in ind:
        data_row(ws, r, [a, b, c_, d], [None, None, MONEY, None], [AL_C, AL_L, AL_R, AL_L],
                 font=H_BOLD if a in ("26", "28") else None,
                 fill=FILL_GREEN if a == "28" else None)
        r += 1
    r += 1
    note(ws, r, 4, "两条硬勾稽：① 现金净增加额 1,052,835.33 ＝ 期末货币资金 2,311,335.33 － 期初 1,258,500.00；"
                   "② 间接法经营净额 102,035.33 ＝ 直接法经营净额 102,035.33。", 34)
    ws.freeze_panes = "A4"

    # ---------- 10 勾稽检查表
    ws = wb.create_sheet("10 勾稽检查表")
    init_sheet(ws, "报表勾稽检查表", "六项勾稽全部通过方可出具报表｜差异容差 0.005 元",
               [6, 44, 18, 18, 14, 10, 40])
    header_row(ws, 3, ["序号", "勾稽关系", "左方金额", "右方金额", "差异", "结论", "说明"])
    r = 4
    items = [
        ("资产总计 ＝ 负债合计 ＋ 所有者权益合计", M.total_assets, M.total_le,
         "资产负债表内部平衡"),
        ("科目余额表期末借方合计 ＝ 贷方合计", M.TB_D_CLOSE, M.TB_C_CLOSE,
         "全部 41 个末级科目试算平衡"),
        ("净利润 ＝ 未分配利润期末 － 期初", M.net_profit, M.retained - 552000.0,
         "利润表与资产负债表衔接"),
        ("现金净增加额 ＝ 期末货币资金 － 期初货币资金", M.cf_net, M.cash - M.cash_begin,
         "现金流量表与资产负债表衔接"),
        ("直接法经营净额 ＝ 间接法经营净额", M.cf_oper, M.indirect,
         "现金流量表内部验证"),
        ("应收账款余额 ＝ 账龄表合计", M.ar, sum(sum(d.values()) for d in M.CLOSE_AGING.values()),
         "账表相符"),
        ("凭证借方合计 ＝ 凭证贷方合计", M.TOTAL_DEBIT, M.TOTAL_CREDIT, "51 张凭证全部平衡"),
        ("应付账款贷方明细 － 借方明细 ＝ 科目余额表 2202 余额",
         M.ap - M.ap_debit, -M.sum_bal("2202"), "重分类前后净额一致"),
    ]
    for i, (a, b, c_, d) in enumerate(items, 1):
        diff = round(b - c_, 2)
        data_row(ws, r, [i, a, b, c_, diff, "通过" if abs(diff) < 0.005 else "不通过", d],
                 [None, None, MONEY, MONEY, MONEY, None, None],
                 [AL_C, AL_L, AL_R, AL_R, AL_R, AL_C, AL_L])
        ws.cell(row=r, column=6).fill = FILL_GREEN if abs(diff) < 0.005 else PatternFill("solid", fgColor="FFFFC7CE")
        r += 1
    r += 1
    note(ws, r, 7, "说明：第 2 项在表中以「库存现金」为例演示逐项验算方法，完整验算见「03 科目余额表」的试算平衡区；"
                   "第 8 项验证往来重分类只是「搬家」，不改变资产负债总额。", 34)
    r += 1
    note(ws, r, 7, "重要提醒：恒等式平不代表账一定对。若两边同时多记或少记（如漏做采购暂估），等式仍然平——"
                   "所以还必须做「报表数与明细账一致」这一关，以及凭证抽查。", 34)
    ws.freeze_panes = "A4"

    wb.save(F_ANSWER)
    return F_ANSWER


if __name__ == "__main__":
    p = build_practice()
    a = build_answer()
    print("OK", p)
    print("OK", a)
