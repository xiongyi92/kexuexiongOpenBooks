# -*- coding: utf-8 -*-
"""晨光贸易 2026 年 7 月账务练习账套 —— 数据模型与勾稽计算。

全部数据以复式记账为唯一事实来源：凭证 -> 发生额 -> 科目余额表 -> 三大报表。
生成脚本末尾对借贷平衡、报表平衡、现金流勾稽做断言，任何一处不平立即失败。
"""
from collections import defaultdict
from datetime import date

# ---------------------------------------------------------------- 会计科目
# (编码, 名称, 一级科目, 类别, 余额方向, 辅助核算)
ACCOUNTS = [
    ("1001", "库存现金", "库存现金", "资产", "借", ""),
    ("1002", "银行存款", "银行存款", "资产", "借", "开户行"),
    ("1122", "应收账款", "应收账款", "资产", "借", "客户"),
    ("1123", "预付账款", "预付账款", "资产", "借", "供应商"),
    ("1405", "库存商品", "库存商品", "资产", "借", "商品品类"),
    ("1601", "固定资产", "固定资产", "资产", "借", "资产类别"),
    ("1602", "累计折旧", "累计折旧", "资产", "贷", "资产类别"),
    ("2001", "短期借款", "短期借款", "负债", "贷", "贷款机构"),
    ("2202", "应付账款", "应付账款", "负债", "贷", "供应商"),
    ("2211.01", "应付职工薪酬—工资", "应付职工薪酬", "负债", "贷", "部门"),
    ("2211.02", "应付职工薪酬—社保公积金", "应付职工薪酬", "负债", "贷", "部门"),
    ("2221.01.01", "应交税费—应交增值税（进项税额）", "应交税费", "负债", "借", ""),
    ("2221.01.02", "应交税费—应交增值税（销项税额）", "应交税费", "负债", "贷", ""),
    ("2221.01.03", "应交税费—应交增值税（转出未交增值税）", "应交税费", "负债", "借", ""),
    ("2221.02", "应交税费—未交增值税", "应交税费", "负债", "贷", ""),
    ("2221.03", "应交税费—应交城市维护建设税", "应交税费", "负债", "贷", ""),
    ("2221.04", "应交税费—应交教育费附加", "应交税费", "负债", "贷", ""),
    ("2221.05", "应交税费—应交地方教育附加", "应交税费", "负债", "贷", ""),
    ("2221.06", "应交税费—应交印花税", "应交税费", "负债", "贷", ""),
    ("2221.07", "应交税费—应交企业所得税", "应交税费", "负债", "贷", ""),
    ("2221.08", "应交税费—应交个人所得税", "应交税费", "负债", "贷", ""),
    ("2241.01", "其他应付款—社保公积金（个人）", "其他应付款", "负债", "贷", ""),
    ("4001", "实收资本", "实收资本", "权益", "贷", "股东"),
    ("4103", "本年利润", "本年利润", "权益", "贷", ""),
    ("4104.01", "利润分配—未分配利润", "利润分配", "权益", "贷", ""),
    ("6001", "主营业务收入", "主营业务收入", "损益", "贷", "客户/商品"),
    ("6401", "主营业务成本", "主营业务成本", "损益", "借", "商品品类"),
    ("6403", "税金及附加", "税金及附加", "损益", "借", ""),
    ("6601.01", "销售费用—职工薪酬", "销售费用", "损益", "借", "部门"),
    ("6601.02", "销售费用—折旧费", "销售费用", "损益", "借", "部门"),
    ("6601.03", "销售费用—运杂费", "销售费用", "损益", "借", "部门"),
    ("6601.04", "销售费用—差旅费", "销售费用", "损益", "借", "部门"),
    ("6601.05", "销售费用—业务宣传费", "销售费用", "损益", "借", "部门"),
    ("6602.01", "管理费用—职工薪酬", "管理费用", "损益", "借", "部门"),
    ("6602.02", "管理费用—折旧费", "管理费用", "损益", "借", "部门"),
    ("6602.03", "管理费用—房租", "管理费用", "损益", "借", "部门"),
    ("6602.04", "管理费用—办公费", "管理费用", "损益", "借", "部门"),
    ("6603.01", "财务费用—手续费", "财务费用", "损益", "借", ""),
    ("6603.02", "财务费用—利息支出", "财务费用", "损益", "借", ""),
    ("6603.03", "财务费用—利息收入", "财务费用", "损益", "贷", ""),
    ("6801", "所得税费用", "所得税费用", "损益", "借", ""),
]
NAME = {c[0]: c[1] for c in ACCOUNTS}
CAT = {c[0]: c[3] for c in ACCOUNTS}
TOP = {c[0]: c[2] for c in ACCOUNTS}

# ---------------------------------------------------------------- 期初余额（表4-1）
OPENING = {
    "1001": (8500.00, 0), "1002": (1250000.00, 0), "1122": (385000.00, 0),
    "1123": (120000.00, 0), "1405": (780000.00, 0), "1601": (450000.00, 0),
    "1602": (0, 96000.00), "2202": (0, 268000.00),
    "2211.01": (0, 40000.00), "2211.02": (0, 5000.00),
    "2221.02": (0, 20000.00), "2221.03": (0, 1400.00), "2221.04": (0, 600.00),
    "2221.05": (0, 400.00), "2221.07": (0, 10100.00),
    "4001": (0, 2000000.00), "4104.01": (0, 552000.00),
}

# 期初往来明细（表4-1 辅助核算拆分）
OPEN_AR = {"华文公司": 120000.00, "中兴办公": 75000.00, "海天文具": 35000.00,
           "宏图商行": 40000.00, "联合采购": 65000.00, "久远办公": 50000.00}
OPEN_AP = {"华信纸品": 168000.00, "恒达文具": 100000.00}

# 期初账龄（表6-1）
BUCKETS = ["1 个月以内", "1~3 个月", "3~6 个月", "6 个月以上"]
OPEN_AGING = {
    "华文公司": {"1 个月以内": 120000.00},
    "中兴办公": {"1 个月以内": 45000.00, "1~3 个月": 30000.00},
    "海天文具": {"1~3 个月": 25000.00, "3~6 个月": 10000.00},
    "宏图商行": {"1 个月以内": 20000.00, "3~6 个月": 15000.00, "6 个月以上": 5000.00},
    "联合采购": {"3~6 个月": 40000.00, "6 个月以上": 25000.00},
    "久远办公": {"6 个月以上": 50000.00},
}

# ---------------------------------------------------------------- 进项发票（30 张）
PURCHASES = [
    (date(2026, 7, 2), "华信纸品", [8600, 12400, 6800]),
    (date(2026, 7, 3), "恒达文具", [9800, 10200, 8600]),
    (date(2026, 7, 6), "中信包装", [10500, 9700, 11200]),
    (date(2026, 7, 9), "华信纸品", [9200, 11500, 7300]),
    (date(2026, 7, 11), "恒达文具", [11400, 9600, 10800]),
    (date(2026, 7, 14), "中信包装", [8800, 10600, 9900]),
    (date(2026, 7, 16), "华信纸品", [10800, 8900, 9500]),
    (date(2026, 7, 19), "恒达文具", [9900, 10300, 9400]),
    (date(2026, 7, 21), "中信包装", [9500, 10100, 9700]),
    (date(2026, 7, 23), "华信纸品", [12000, 11300, 11700]),
]
SUPPLIER_FULL = {"华信纸品": "华信纸品有限公司",
                 "恒达文具": "恒达文具制造有限公司",
                 "中信包装": "中信包装材料有限公司",
                 "远洋物流": "远洋物流有限公司"}
FREIGHT = dict(date=date(2026, 7, 8), supplier="远洋物流",
               amount=2000.00, rate=0.09)

# ---------------------------------------------------------------- 销项发票（6 张）
SALES = [
    (date(2026, 7, 5), "华文公司", "A4 复印纸、中性笔等办公用品", 200000.00),
    (date(2026, 7, 8), "中兴办公", "文件夹、档案盒等办公用品", 120000.00),
    (date(2026, 7, 12), "海天文具", "打印耗材、便签纸等办公用品", 80000.00),
    (date(2026, 7, 18), "宏图商行", "办公家具配件、白板等", 60000.00),
    (date(2026, 7, 24), "联合采购", "学生文具套装", 40000.00),
    (date(2026, 7, 28), "启明学校", "教学用办公用品", 20000.00),
]

# 回款（客户，金额）
COLLECTIONS = [
    (date(2026, 7, 6), "华文公司", 120000.00),
    (date(2026, 7, 12), "中兴办公", 75000.00),
    (date(2026, 7, 18), "海天文具", 35000.00),
    (date(2026, 7, 22), "宏图商行", 40000.00),
    (date(2026, 7, 26), "联合采购", 65000.00),
    (date(2026, 7, 29), "华文公司", 226000.00),
    (date(2026, 7, 30), "中兴办公", 135600.00),
    (date(2026, 7, 31), "宏图商行", 23000.00),      # 未达账项：银行未入账
]

# ---------------------------------------------------------------- 工资
STAFF = [
    # 姓名, 部门, 岗位, 月工资, 专项附加扣除
    ("老李", "管理部", "财务经理", 10000.00, 2000.00),
    ("小王", "管理部", "会计", 8000.00, 1000.00),
    ("小张", "销售部", "销售主管", 10000.00, 0.00),
    ("小周", "销售部", "销售代表", 5000.00, 1000.00),
    ("小吴", "销售部", "销售代表", 4000.00, 0.00),
    ("小郑", "销售部", "销售代表", 3000.00, 0.00),
]
SI_RATE, HF_RATE, UNIT_RATE = 0.10, 0.10, 0.20   # 个人社保/公积金、单位社保公积金
TAX_BRACKETS = [(0, 0.03, 0), (36000, 0.10, 2520), (144000, 0.20, 16920),
                (300000, 0.25, 31920), (420000, 0.30, 52920),
                (660000, 0.35, 85920), (960000, 0.45, 181920)]

# ---------------------------------------------------------------- 固定资产
ASSETS = [
    # 编号, 名称, 类别, 使用部门, 取得日期, 原值, 残值率, 年限, 期初累计折旧
    ("GD-001", "办公室设备（电脑、打印机等）", "电子设备", "管理部", date(2024, 7, 15), 90000.00, 0.05, 3, 28500.00),
    ("GD-002", "运输货车（厢式）", "运输设备", "销售部", date(2024, 7, 20), 120000.00, 0.05, 5, 22800.00),
    ("GD-003", "办公家具", "家具用具", "管理部", date(2024, 7, 1), 60000.00, 0.05, 3, 19050.00),
    ("GD-004", "仓储设备（货架、叉车）", "机器设备", "销售部", date(2024, 1, 10), 180000.00, 0.05, 10, 25650.00),
]
NEW_ASSET = dict(no="GD-005", name="仓储货架（新增）", cat="机器设备", dept="销售部",
                 date=date(2026, 7, 31), cost=45000.00, rate=0.05, years=5)


def tax_of(amount):
    """按 3% / 10% / 20% ... 速算的年度综合所得税额"""
    rate, quick = 0.03, 0
    for lower, r, q in TAX_BRACKETS:
        if amount > lower:
            rate, quick = r, q
    return max(amount * rate - quick, 0)


def monthly_tax(monthly_taxable, months):
    return round(tax_of(max(monthly_taxable, 0) * months), 2)


# ---------------------------------------------------------------- 凭证
VOUCHERS = []


def V(d, summary, lines, source="", vtype="日常业务"):
    VOUCHERS.append(dict(date=d, summary=summary, lines=lines,
                         source=source, vtype=vtype))


def money(x):
    return round(x + 1e-9, 2)


# 1 借款
V(date(2026, 7, 1), "取得工商银行短期借款（期限 6 个月，月利率 4.2‰）",
  [("1002", 1000000.00, 0, "工商银行", "取得借款收到的现金"),
   ("2001", 0, 1000000.00, "工商银行", "")], "借款合同 GJ-2026-07")
# 2 房租
V(date(2026, 7, 1), "支付 7 月办公室房租",
  [("6602.03", 3000.00, 0, "管理部", ""),
   ("1002", 0, 3000.00, "", "支付其他与经营活动有关的现金")], "房租发票 20260701")
# 采购（按 日期+供应商 合并 30 张专票）
pidx = 0
PURCHASE_VOUCHER_INFO = []
for d, supplier, amounts in PURCHASES:
    net = money(sum(amounts))
    tax = money(net * 0.13)
    lines = [("1405", net, 0, "办公用品", ""),
             ("2221.01.01", tax, 0, "", ""),
             ("2202", 0, money(net + tax), supplier, "")]
    V(d, f"采购办公用品入库（{supplier}，{len(amounts)} 张增值税专用发票合并）", lines,
      f"进项发票 {len(amounts)} 张")
    PURCHASE_VOUCHER_INFO.append((d, supplier, net, tax, len(amounts)))
    pidx += len(amounts)

# 销售 + 回款 按日期插入
sales_by_date = {d: (c, g, a) for d, c, g, a in SALES}
coll_by_date = defaultdict(list)
for d, c, a in COLLECTIONS:
    coll_by_date[d].append((c, a))

SEQ = []          # (date, kind, payload) 用于按日期重排
SEQ.append((date(2026, 7, 1), "loan", None))
SEQ.append((date(2026, 7, 1), "rent", None))
for d, supplier, amounts in PURCHASES:
    SEQ.append((d, "purchase", (d, supplier, amounts)))
for d, c, g, a in SALES:
    SEQ.append((d, "sale", (d, c, g, a)))
for d, c, a in COLLECTIONS:
    SEQ.append((d, c and "collect", (d, c, a)))

# 清空后按日期顺序重建（保留手工凭证的插入位置）
VOUCHERS.clear()


def add_loan():
    V(date(2026, 7, 1), "取得工商银行短期借款（期限 6 个月，月利率 4.2‰）",
      [("1002", 1000000.00, 0, "工商银行", "取得借款收到的现金"),
       ("2001", 0, 1000000.00, "工商银行", "")], "借款合同 GJ-2026-07")


def add_rent():
    V(date(2026, 7, 1), "支付 7 月办公室房租",
      [("6602.03", 3000.00, 0, "管理部", ""),
       ("1002", 0, 3000.00, "", "支付其他与经营活动有关的现金")], "房租发票 20260701")


add_loan()
add_rent()

# 7/2 采购、7/3 采购
for d, supplier, amounts in PURCHASES[:2]:
    net = money(sum(amounts)); tax = money(net * 0.13)
    V(d, f"采购办公用品入库（{supplier}，{len(amounts)} 张增值税专用发票合并）",
      [("1405", net, 0, "办公用品", ""), ("2221.01.01", tax, 0, "", ""),
       ("2202", 0, money(net + tax), supplier, "")], f"进项发票 {len(amounts)} 张")

# 7/5 销售 华文
d, c, g, a = SALES[0]
V(d, f"销售办公用品（{c}），货款未收",
  [("1122", money(a * 1.13), 0, c, ""), ("6001", 0, a, c, ""),
   ("2221.01.02", 0, money(a * 0.13), "", "")], "销项发票 20260705")

# 7/6 回款 华文 + 采购 中信
V(date(2026, 7, 6), "收到华文公司货款（6 月应收）",
  [("1002", 120000.00, 0, "", "销售商品、提供劳务收到的现金"),
   ("1122", 0, 120000.00, "华文公司", "")], "银行回单 20260706")
d, supplier, amounts = PURCHASES[2]
net = money(sum(amounts)); tax = money(net * 0.13)
V(d, f"采购办公用品入库（{supplier}，{len(amounts)} 张增值税专用发票合并）",
  [("1405", net, 0, "办公用品", ""), ("2221.01.01", tax, 0, "", ""),
   ("2202", 0, money(net + tax), supplier, "")], f"进项发票 {len(amounts)} 张")

# 7/8 运费 + 付华信 + 销售中兴
V(date(2026, 7, 8), "支付销售商品运输费（取得 9% 增值税专用发票）",
  [("6601.03", 2000.00, 0, "销售部", ""), ("2221.01.01", 180.00, 0, "", ""),
   ("1002", 0, 2180.00, "", "购买商品、接受劳务支付的现金")], "运费发票 20260708")
V(date(2026, 7, 8), "支付华信纸品有限公司货款（6 月应付）",
  [("2202", 168000.00, 0, "华信纸品", ""),
   ("1002", 0, 168000.00, "", "购买商品、接受劳务支付的现金")], "银行回单 20260708")
d, c, g, a = SALES[1]
V(d, f"销售办公用品（{c}），货款未收",
  [("1122", money(a * 1.13), 0, c, ""), ("6001", 0, a, c, ""),
   ("2221.01.02", 0, money(a * 0.13), "", "")], "销项发票 20260708")

# 7/9 采购 华信
d, supplier, amounts = PURCHASES[3]
net = money(sum(amounts)); tax = money(net * 0.13)
V(d, f"采购办公用品入库（{supplier}，{len(amounts)} 张增值税专用发票合并）",
  [("1405", net, 0, "办公用品", ""), ("2221.01.01", tax, 0, "", ""),
   ("2202", 0, money(net + tax), supplier, "")], f"进项发票 {len(amounts)} 张")

# 7/10 发工资
V(date(2026, 7, 10), "发放 6 月工资（代扣个人社保公积金及个人所得税）",
  [("2211.01", 40000.00, 0, "管理部/销售部", ""),
   ("1002", 0, 31868.00, "", "支付给职工以及为职工支付的现金"),
   ("2241.01", 0, 8000.00, "", ""),
   ("2221.08", 0, 132.00, "", "")], "工资表 2026-06")

# 7/11 采购 恒达
d, supplier, amounts = PURCHASES[4]
net = money(sum(amounts)); tax = money(net * 0.13)
V(d, f"采购办公用品入库（{supplier}，{len(amounts)} 张增值税专用发票合并）",
  [("1405", net, 0, "办公用品", ""), ("2221.01.01", tax, 0, "", ""),
   ("2202", 0, money(net + tax), supplier, "")], f"进项发票 {len(amounts)} 张")

# 7/12 回款中兴 + 缴社保 + 销售海天
V(date(2026, 7, 12), "收到中兴办公货款（6 月应收）",
  [("1002", 75000.00, 0, "", "销售商品、提供劳务收到的现金"),
   ("1122", 0, 75000.00, "中兴办公", "")], "银行回单 20260712")
V(date(2026, 7, 12), "缴纳 6 月社保公积金（个人部分）",
  [("2241.01", 8000.00, 0, "", ""),
   ("1002", 0, 8000.00, "", "支付给职工以及为职工支付的现金")], "社保缴费凭证 20260712")
d, c, g, a = SALES[2]
V(d, f"销售办公用品（{c}），货款未收",
  [("1122", money(a * 1.13), 0, c, ""), ("6001", 0, a, c, ""),
   ("2221.01.02", 0, money(a * 0.13), "", "")], "销项发票 20260712")

# 7/14 采购 中信
d, supplier, amounts = PURCHASES[5]
net = money(sum(amounts)); tax = money(net * 0.13)
V(d, f"采购办公用品入库（{supplier}，{len(amounts)} 张增值税专用发票合并）",
  [("1405", net, 0, "办公用品", ""), ("2221.01.01", tax, 0, "", ""),
   ("2202", 0, money(net + tax), supplier, "")], f"进项发票 {len(amounts)} 张")

# 7/15 付恒达 + 缴税
V(date(2026, 7, 15), "支付恒达文具制造有限公司货款（6 月应付）",
  [("2202", 100000.00, 0, "恒达文具", ""),
   ("1002", 0, 100000.00, "", "购买商品、接受劳务支付的现金")], "银行回单 20260715")
V(date(2026, 7, 15), "缴纳 6 月增值税及附加税费、企业所得税",
  [("2221.02", 20000.00, 0, "", ""), ("2221.03", 1400.00, 0, "", ""),
   ("2221.04", 600.00, 0, "", ""), ("2221.05", 400.00, 0, "", ""),
   ("2221.07", 10100.00, 0, "", ""),
   ("1002", 0, 32500.00, "", "支付的各项税费")], "税收缴款书 20260715")

# 7/16 采购 华信
d, supplier, amounts = PURCHASES[6]
net = money(sum(amounts)); tax = money(net * 0.13)
V(d, f"采购办公用品入库（{supplier}，{len(amounts)} 张增值税专用发票合并）",
  [("1405", net, 0, "办公用品", ""), ("2221.01.01", tax, 0, "", ""),
   ("2202", 0, money(net + tax), supplier, "")], f"进项发票 {len(amounts)} 张")

# 7/18 回款海天 + 销售宏图
V(date(2026, 7, 18), "收到海天文具货款（6 月应收）",
  [("1002", 35000.00, 0, "", "销售商品、提供劳务收到的现金"),
   ("1122", 0, 35000.00, "海天文具", "")], "银行回单 20260718")
d, c, g, a = SALES[3]
V(d, f"销售办公用品（{c}），货款未收",
  [("1122", money(a * 1.13), 0, c, ""), ("6001", 0, a, c, ""),
   ("2221.01.02", 0, money(a * 0.13), "", "")], "销项发票 20260718")

# 7/19 采购 恒达
d, supplier, amounts = PURCHASES[7]
net = money(sum(amounts)); tax = money(net * 0.13)
V(d, f"采购办公用品入库（{supplier}，{len(amounts)} 张增值税专用发票合并）",
  [("1405", net, 0, "办公用品", ""), ("2221.01.01", tax, 0, "", ""),
   ("2202", 0, money(net + tax), supplier, "")], f"进项发票 {len(amounts)} 张")

# 7/20 办公费 + 宣传费
V(date(2026, 7, 20), "支付办公费（快递、水电、零星采购，普通发票）",
  [("6602.04", 1441.67, 0, "管理部", ""),
   ("1002", 0, 1441.67, "", "支付其他与经营活动有关的现金")], "费用报销单 BX-0720")
V(date(2026, 7, 20), "支付业务宣传费（展会物料，普通发票）",
  [("6601.05", 8275.00, 0, "销售部", ""),
   ("1002", 0, 8275.00, "", "支付其他与经营活动有关的现金")], "费用报销单 BX-0721")

# 7/21 采购 中信 + 利息收入
d, supplier, amounts = PURCHASES[8]
net = money(sum(amounts)); tax = money(net * 0.13)
V(d, f"采购办公用品入库（{supplier}，{len(amounts)} 张增值税专用发票合并）",
  [("1405", net, 0, "办公用品", ""), ("2221.01.01", tax, 0, "", ""),
   ("2202", 0, money(net + tax), supplier, "")], f"进项发票 {len(amounts)} 张")
V(date(2026, 7, 21), "收到银行存款利息",
  [("1002", 1200.00, 0, "", "收到其他与经营活动有关的现金"),
   ("6603.03", 0, 1200.00, "", "")], "银行利息单 20260721")

# 7/22 回款宏图
V(date(2026, 7, 22), "收到宏图商行货款（6 月应收）",
  [("1002", 40000.00, 0, "", "销售商品、提供劳务收到的现金"),
   ("1122", 0, 40000.00, "宏图商行", "")], "银行回单 20260722")

# 7/23 采购 华信
d, supplier, amounts = PURCHASES[9]
net = money(sum(amounts)); tax = money(net * 0.13)
V(d, f"采购办公用品入库（{supplier}，{len(amounts)} 张增值税专用发票合并）",
  [("1405", net, 0, "办公用品", ""), ("2221.01.01", tax, 0, "", ""),
   ("2202", 0, money(net + tax), supplier, "")], f"进项发票 {len(amounts)} 张")

# 7/24 销售 联合
d, c, g, a = SALES[4]
V(d, f"销售办公用品（{c}），货款未收",
  [("1122", money(a * 1.13), 0, c, ""), ("6001", 0, a, c, ""),
   ("2221.01.02", 0, money(a * 0.13), "", "")], "销项发票 20260724")

# 7/25 预付中信
V(date(2026, 7, 25), "预付中信包装材料有限公司下批货款（合同约定预付 30%）",
  [("2202", 200000.00, 0, "中信包装", ""),
   ("1002", 0, 200000.00, "", "购买商品、接受劳务支付的现金")], "付款申请单 FK-0725")

# 7/26 回款联合
V(date(2026, 7, 26), "收到联合采购货款（6 月应收）",
  [("1002", 65000.00, 0, "", "销售商品、提供劳务收到的现金"),
   ("1122", 0, 65000.00, "联合采购", "")], "银行回单 20260726")

# 7/28 销售 启明
d, c, g, a = SALES[5]
V(d, f"销售办公用品（{c}），货款未收",
  [("1122", money(a * 1.13), 0, c, ""), ("6001", 0, a, c, ""),
   ("2221.01.02", 0, money(a * 0.13), "", "")], "销项发票 20260728")

# 7/29 回款华文 + 7/30 回款中兴
V(date(2026, 7, 29), "收到华文公司货款（7 月 5 日销售）",
  [("1002", 226000.00, 0, "", "销售商品、提供劳务收到的现金"),
   ("1122", 0, 226000.00, "华文公司", "")], "银行回单 20260729")
V(date(2026, 7, 30), "收到中兴办公货款（7 月 8 日销售）",
  [("1002", 135600.00, 0, "", "销售商品、提供劳务收到的现金"),
   ("1122", 0, 135600.00, "中兴办公", "")], "银行回单 20260730")

# 7/31 支票付款 + 回款宏图 + 差旅 + 手续费 + 利息 + 设备
V(date(2026, 7, 31), "开出转账支票支付华信纸品货款（银行尚未划款）",
  [("2202", 56500.00, 0, "华信纸品", ""),
   ("1002", 0, 56500.00, "", "购买商品、接受劳务支付的现金")], "转账支票 0021875")
V(date(2026, 7, 31), "收到宏图商行货款（部分回款，银行尚未入账）",
  [("1002", 23000.00, 0, "", "销售商品、提供劳务收到的现金"),
   ("1122", 0, 23000.00, "宏图商行", "")], "银行进账单 20260731")
V(date(2026, 7, 31), "支付销售部门差旅费（普通发票）",
  [("6601.04", 5000.00, 0, "销售部", ""),
   ("1002", 0, 5000.00, "", "支付其他与经营活动有关的现金")], "差旅报销单 CL-0731")
V(date(2026, 7, 31), "支付银行手续费",
  [("6603.01", 2000.00, 0, "", ""),
   ("1002", 0, 2000.00, "", "支付其他与经营活动有关的现金")], "银行收费凭证 20260731")
V(date(2026, 7, 31), "支付短期借款 7 月利息（1,000,000×4.2‰）",
  [("6603.02", 4200.00, 0, "", ""),
   ("1002", 0, 4200.00, "", "分配股利、利润或偿付利息支付的现金")], "利息清单 20260731")
V(date(2026, 7, 31), "购入仓储货架一批（小规模纳税人，普通发票，进项不可抵扣）",
  [("1601", 45000.00, 0, "机器设备", ""),
   ("1002", 0, 45000.00, "", "购建固定资产、无形资产和其他长期资产支付的现金")], "设备发票 20260731")

# 月末结转
V(date(2026, 7, 31), "结转已售商品成本",
  [("6401", 390000.00, 0, "办公用品", ""), ("1405", 0, 390000.00, "办公用品", "")],
  "成本计算表", "期末结转")

DEP = []
for no, nm, cat, dept, d0, cost, rate, years, acc in ASSETS:
    m = money(cost * (1 - rate) / years / 12)
    DEP.append((no, nm, dept, cost, rate, years, m, acc))
dep_total = money(sum(x[6] for x in DEP))
dep_admin = money(sum(x[6] for x in DEP if x[2] == "管理部"))
dep_sales = money(dep_total - dep_admin)

V(date(2026, 7, 31), "计提 7 月固定资产折旧（直线法，当月新增资产下月起计提）",
  [("6602.02", dep_admin, 0, "管理部", ""), ("6601.02", dep_sales, 0, "销售部", ""),
   ("1602", 0, dep_total, "", "")], "折旧计提表", "期末结转")

# 工资计提
PAYROLL = []
for nm, dept, post, salary, spec in STAFF:
    si = money(salary * SI_RATE)
    hf = money(salary * HF_RATE)
    unit = money(salary * UNIT_RATE)
    taxable_m = max(salary - 5000 - si - hf - spec, 0.0)
    cum6 = monthly_tax(taxable_m, 6)
    cum7 = monthly_tax(taxable_m, 7)
    cur = money(cum7 - cum6)
    net_pay = money(salary - si - hf - cur)
    PAYROLL.append(dict(name=nm, dept=dept, post=post, salary=salary, si=si, hf=hf,
                        spec=spec, taxable_m=taxable_m, cum6=cum6, cum7=cum7,
                        tax=cur, net=net_pay, unit=unit))
pay_total = money(sum(p["salary"] for p in PAYROLL))
pay_unit = money(sum(p["unit"] for p in PAYROLL))
pay_admin = money(sum(p["salary"] + p["unit"] for p in PAYROLL if p["dept"] == "管理部"))
pay_sales = money(sum(p["salary"] + p["unit"] for p in PAYROLL if p["dept"] == "销售部"))
pay_tax = money(sum(p["tax"] for p in PAYROLL))
pay_net = money(sum(p["net"] for p in PAYROLL))

V(date(2026, 7, 31), "计提 7 月工资及单位承担的社会保险费、住房公积金",
  [("6602.01", pay_admin, 0, "管理部", ""), ("6601.01", pay_sales, 0, "销售部", ""),
   ("2211.01", 0, pay_total, "", ""), ("2211.02", 0, pay_unit, "", "")],
  "工资表 2026-07", "期末结转")

# 增值税与附加
out_vat = money(sum(a * 0.13 for _, _, _, a in SALES))
in_vat = money(sum(money(sum(x[2]) * 0.13) for x in PURCHASES)) + 180.00
vat_payable = money(out_vat - in_vat)
city = money(vat_payable * 0.07)
edu = money(vat_payable * 0.03)
edu2 = money(vat_payable * 0.02)
stamp_base = 300000.00 + 520000.00 + 2000.00
stamp = money(stamp_base * 0.0003)
surcharge = money(city + edu + edu2)
tax_surcharge = money(surcharge + stamp)

V(date(2026, 7, 31), "计提 7 月城市维护建设税、教育费附加、地方教育附加及印花税",
  [("6403", tax_surcharge, 0, "", ""), ("2221.03", 0, city, "", ""),
   ("2221.04", 0, edu, "", ""), ("2221.05", 0, edu2, "", ""),
   ("2221.06", 0, stamp, "", "")], "税金计算表", "期末结转")
V(date(2026, 7, 31), "转出本月未交增值税",
  [("2221.01.03", vat_payable, 0, "", ""), ("2221.02", 0, vat_payable, "", "")],
  "增值税计算表", "期末结转")

# 先算利润总额
def bal(code):
    d0, c0 = OPENING.get(code, (0, 0))
    d1 = sum(l[1] for v in VOUCHERS for l in v["lines"] if l[0] == code)
    c1 = sum(l[2] for v in VOUCHERS for l in v["lines"] if l[0] == code)
    return money(d0 - c0 + d1 - c1)


def sum_bal(prefix):
    return money(sum(bal(c[0]) for c in ACCOUNTS if c[0].startswith(prefix)))


revenue = money(-sum_bal("6001"))
cost = money(sum_bal("6401"))
tax_add = money(sum_bal("6403"))
sell_exp = money(sum_bal("6601"))
admin_exp = money(sum_bal("6602"))
fin_exp = money(sum_bal("6603"))
operating = money(revenue - cost - tax_add - sell_exp - admin_exp - fin_exp)
income_tax = money(operating * 0.25)
net_profit = money(operating - income_tax)

V(date(2026, 7, 31), "计提 7 月企业所得税（利润总额 ×25%）",
  [("6801", income_tax, 0, "", ""), ("2221.07", 0, income_tax, "", "")],
  "所得税计算表", "期末结转")
V(date(2026, 7, 31), "结转本月主营业务收入至本年利润",
  [("6001", revenue, 0, "", ""), ("4103", 0, revenue, "", "")], "", "期末结转")
V(date(2026, 7, 31), "结转本月成本、费用及所得税至本年利润",
  [("4103", money(cost + tax_add + sell_exp + admin_exp +
                  sum_bal("6603.01") + sum_bal("6603.02") + income_tax), 0, "", ""),
   ("6401", 0, cost, "", ""), ("6403", 0, tax_add, "", ""),
   ("6601.01", 0, money(sum_bal("6601.01")), "", ""),
   ("6601.02", 0, money(sum_bal("6601.02")), "", ""),
   ("6601.03", 0, money(sum_bal("6601.03")), "", ""),
   ("6601.04", 0, money(sum_bal("6601.04")), "", ""),
   ("6601.05", 0, money(sum_bal("6601.05")), "", ""),
   ("6602.01", 0, money(sum_bal("6602.01")), "", ""),
   ("6602.02", 0, money(sum_bal("6602.02")), "", ""),
   ("6602.03", 0, money(sum_bal("6602.03")), "", ""),
   ("6602.04", 0, money(sum_bal("6602.04")), "", ""),
   ("6603.01", 0, money(sum_bal("6603.01")), "", ""),
   ("6603.02", 0, money(sum_bal("6603.02")), "", ""),
   ("6801", 0, income_tax, "", "")], "", "期末结转")
V(date(2026, 7, 31), "结转财务费用—利息收入至本年利润",
  [("6603.03", 1200.00, 0, "", ""), ("4103", 0, 1200.00, "", "")], "", "期末结转")
V(date(2026, 7, 31), "结转本年利润至利润分配—未分配利润",
  [("4103", net_profit, 0, "", ""), ("4104.01", 0, net_profit, "", "")], "", "期末结转")

# 凭证编号
for i, v in enumerate(VOUCHERS, 1):
    v["no"] = f"记-{i:03d}"

# ---------------------------------------------------------------- 发生额
MOV = defaultdict(lambda: [0.0, 0.0])
for v in VOUCHERS:
    for code, d, c, aux, cf in v["lines"]:
        MOV[code][0] += d
        MOV[code][1] += c
for k in MOV:
    MOV[k] = [money(MOV[k][0]), money(MOV[k][1])]


def closing(code):
    d0, c0 = OPENING.get(code, (0, 0))
    return money(d0 - c0 + MOV[code][0] - MOV[code][1])


def closing_dc(code):
    n = closing(code)
    return (max(n, 0), max(-n, 0))


# 往来明细
AR_AUX = defaultdict(lambda: [0.0, 0.0])
AP_AUX = defaultdict(lambda: [0.0, 0.0])
for v in VOUCHERS:
    for code, d, c, aux, cf in v["lines"]:
        if code == "1122":
            AR_AUX[aux][0] += d; AR_AUX[aux][1] += c
        if code == "2202":
            AP_AUX[aux][0] += d; AP_AUX[aux][1] += c
AR_CLOSE = {}
for k, v in OPEN_AR.items():
    AR_CLOSE[k] = money(v + AR_AUX[k][0] - AR_AUX[k][1])
for k, v in AR_AUX.items():
    if k not in AR_CLOSE:
        AR_CLOSE[k] = money(AR_AUX[k][0] - AR_AUX[k][1])
AP_CLOSE = {}
for k, v in OPEN_AP.items():
    AP_CLOSE[k] = money(v + AP_AUX[k][1] - AP_AUX[k][0])
for k, v in AP_AUX.items():
    if k not in AP_CLOSE:
        AP_CLOSE[k] = money(AP_AUX[k][1] - AP_AUX[k][0])

# ---------------------------------------------------------------- 期末账龄
# FIFO：先冲最早的应收
OPEN_ITEMS = []
for cust, buckets in OPEN_AGING.items():
    for b in BUCKETS:
        if buckets.get(b):
            OPEN_ITEMS.append(dict(cust=cust, amount=buckets[b], bucket=b))
NEW_ITEMS = [dict(cust=c, amount=money(a * 1.13), bucket="1 个月以内") for _, c, _, a in SALES]
items = OPEN_ITEMS + NEW_ITEMS
for _, cust, amt in COLLECTIONS:
    left = amt
    for it in items:
        if it["cust"] == cust and it["amount"] > 0 and left > 0:
            use = min(it["amount"], left)
            it["amount"] = money(it["amount"] - use)
            left = money(left - use)
CLOSE_AGING = defaultdict(lambda: {b: 0.0 for b in BUCKETS})
for it in items:
    if it["amount"] > 0:
        CLOSE_AGING[it["cust"]][it["bucket"]] += it["amount"]
CLOSE_AGING = {k: {b: money(v) for b, v in d.items()} for k, d in CLOSE_AGING.items()}
ar_total = money(sum(sum(d.values()) for d in CLOSE_AGING.values()))

# ---------------------------------------------------------------- 银行
PARTY_RULES = [
    ("取得工商银行短期借款", "工商银行城东支行"),
    ("支付 7 月办公室房租", "城东物业管理公司"),
    ("支付销售商品运输费", "远洋物流有限公司"),
    ("支付华信纸品有限公司货款", "华信纸品有限公司"),
    ("开出转账支票支付华信纸品货款", "华信纸品有限公司"),
    ("支付恒达文具制造有限公司货款", "恒达文具制造有限公司"),
    ("预付中信包装材料有限公司", "中信包装材料有限公司"),
    ("发放 6 月工资", "员工（6 人）"),
    ("缴纳 6 月社保公积金", "市社会保险基金管理中心"),
    ("缴纳 6 月增值税", "国家税务总局城东区税务局"),
    ("收到银行存款利息", "工商银行城东支行"),
    ("支付销售部门差旅费", "员工报销"),
    ("支付银行手续费", "工商银行城东支行"),
    ("支付短期借款 7 月利息", "工商银行城东支行"),
    ("购入仓储货架", "恒信设备经营部"),
]


def party_of(summary):
    for key, name in PARTY_RULES:
        if summary.startswith(key):
            return name
    for cust in ("华文公司", "中兴办公", "海天文具", "宏图商行", "联合采购", "启明学校"):
        if summary.startswith("收到") and cust in summary:
            return cust + "（客户）"
    if summary.startswith("收到"):
        for cust in ("华文公司", "中兴办公", "海天文具", "宏图商行", "联合采购", "启明学校"):
            if cust in summary:
                return cust + "（客户）"
    return "—"


BANK_FLOW = []
run = 1250000.00
for v in VOUCHERS:
    for code, d, c, aux, cf in v["lines"]:
        if code == "1002":
            run = money(run + d - c)
            BANK_FLOW.append(dict(date=v["date"], no=v["no"], summary=v["summary"],
                                  party=party_of(v["summary"]), inflow=d, outflow=c,
                                  bal=run, cf=cf))
bank_close = run

# 对账单：剔除两笔未达账项
stmt = []
run2 = 1250000.00
for row in BANK_FLOW:
    if row["no"] in ("记-0" + str(0),):
        pass
    skip = (row["summary"].startswith("开出转账支票") or
            row["summary"].startswith("收到宏图商行货款（部分回款"))
    if skip:
        continue
    run2 = money(run2 + row["inflow"] - row["outflow"])
    stmt.append(dict(date=row["date"], summary=row["summary"], party=row["party"],
                     inflow=row["inflow"], outflow=row["outflow"], bal=run2))
stmt_close = run2

# ---------------------------------------------------------------- 报表
cash = money(sum_bal("1001") + sum_bal("1002"))
ar = money(sum_bal("1122"))
prepaid_open = money(sum_bal("1123"))
ap_credit = money(sum(v for v in AP_CLOSE.values() if v > 0))
ap_debit = money(-sum(v for v in AP_CLOSE.values() if v < 0))
prepaid = money(prepaid_open + ap_debit)
inventory = money(sum_bal("1405"))
fa_cost = money(sum_bal("1601"))
acc_dep = money(-sum_bal("1602"))
fa_net = money(fa_cost - acc_dep)
total_assets = money(cash + ar + prepaid + inventory + fa_net)

ap = ap_credit
payroll_pay = money(-sum_bal("2211"))
tax_pay = money(-sum_bal("2221"))
other_pay = money(-sum_bal("2241"))
short_loan = money(-sum_bal("2001"))
total_liab = money(ap + payroll_pay + tax_pay + other_pay + short_loan)
capital = money(-sum_bal("4001"))
retained = money(-sum_bal("4104") - sum_bal("4103"))
total_equity = money(capital + retained)
total_le = money(total_liab + total_equity)

# 现金流量（直接法，取凭证现金流量项目）
CF = defaultdict(lambda: [0.0, 0.0])
for v in VOUCHERS:
    for code, d, c, aux, cf in v["lines"]:
        if code == "1002" and cf:
            CF[cf][0] += d
            CF[cf][1] += c
CF = {k: [money(v[0]), money(v[1])] for k, v in CF.items()}
cf_sales = CF["销售商品、提供劳务收到的现金"][0]
cf_other_in = CF["收到其他与经营活动有关的现金"][0]
cf_buy = CF["购买商品、接受劳务支付的现金"][1]
cf_staff = CF["支付给职工以及为职工支付的现金"][1]
cf_tax = CF["支付的各项税费"][1]
cf_other_out = CF["支付其他与经营活动有关的现金"][1]
cf_oper_in = money(cf_sales + cf_other_in)
cf_oper_out = money(cf_buy + cf_staff + cf_tax + cf_other_out)
cf_oper = money(cf_oper_in - cf_oper_out)
cf_inv_out = CF["购建固定资产、无形资产和其他长期资产支付的现金"][1]
cf_inv = money(-cf_inv_out)
cf_fin_in = CF["取得借款收到的现金"][0]
cf_fin_out = CF["分配股利、利润或偿付利息支付的现金"][1]
cf_fin = money(cf_fin_in - cf_fin_out)
cf_net = money(cf_oper + cf_inv + cf_fin)
cash_begin = money(8500.00 + 1250000.00)
cash_end = money(cash_begin + cf_net)

# 间接法
d_ar = money(OPEN_AR_TOTAL := sum(OPEN_AR.values()) - ar)
d_inv = money(780000.00 - inventory)
d_prepaid = money(120000.00 - prepaid)
d_ap = money(ap - 268000.00)
d_pay = money(payroll_pay - 45000.00)
d_tax = money(tax_pay - 32500.00)
indirect = money(net_profit + dep_total + d_ar + d_inv + d_prepaid + d_ap
                 + d_pay + d_tax + 4200.00)

# ---------------------------------------------------------------- 断言
eps = 0.005
assert abs(sum(OPENING.get(c[0], (0, 0))[0] for c in ACCOUNTS) -
           sum(OPENING.get(c[0], (0, 0))[1] for c in ACCOUNTS)) < eps, "期初不平"
assert abs(sum(MOV[c[0]][0] for c in ACCOUNTS) - sum(MOV[c[0]][1] for c in ACCOUNTS)) < eps, "发生额不平"
for v in VOUCHERS:
    dd = sum(l[1] for l in v["lines"]); cc = sum(l[2] for l in v["lines"])
    assert abs(dd - cc) < eps, f"凭证 {v['no']} 借贷不平 {dd} {cc}"
assert abs(sum(max(closing(c[0]), 0) for c in ACCOUNTS) -
           sum(max(-closing(c[0]), 0) for c in ACCOUNTS)) < eps, "期末不平"
assert abs(total_assets - total_le) < eps, f"资产负债表不平 {total_assets} {total_le}"
assert abs(cash_end - cash) < eps, f"现金勾稽不平 {cash_end} {cash}"
assert abs(indirect - cf_oper) < eps, f"间接法不平 {indirect} {cf_oper}"
assert abs(ar - ar_total) < eps, f"账龄不平 {ar} {ar_total}"
assert abs(payroll_pay - money(53000.00)) < eps, payroll_pay

TOTAL_DEBIT = money(sum(l[1] for v in VOUCHERS for l in v['lines']))
TOTAL_CREDIT = money(sum(l[2] for v in VOUCHERS for l in v['lines']))
STMT = stmt
TB_D_CLOSE = money(sum(max(closing(c[0]), 0) for c in ACCOUNTS))
TB_C_CLOSE = money(sum(max(-closing(c[0]), 0) for c in ACCOUNTS))

MODEL = dict(
    MOV=MOV, closing=closing, closing_dc=closing_dc, sum_bal=sum_bal,
    revenue=revenue, cost=cost, tax_add=tax_add, sell_exp=sell_exp,
    admin_exp=admin_exp, fin_exp=fin_exp, operating=operating,
    income_tax=income_tax, net_profit=net_profit,
    cash=cash, ar=ar, prepaid=prepaid, inventory=inventory,
    fa_cost=fa_cost, acc_dep=acc_dep, fa_net=fa_net, total_assets=total_assets,
    ap=ap, ap_credit=ap_credit, ap_debit=ap_debit, payroll_pay=payroll_pay,
    tax_pay=tax_pay, other_pay=other_pay, short_loan=short_loan,
    total_liab=total_liab, capital=capital, retained=retained,
    total_equity=total_equity, total_le=total_le,
    cf_sales=cf_sales, cf_other_in=cf_other_in, cf_buy=cf_buy,
    cf_staff=cf_staff, cf_tax=cf_tax, cf_other_out=cf_other_out,
    cf_oper_in=cf_oper_in, cf_oper_out=cf_oper_out, cf_oper=cf_oper,
    cf_inv=cf_inv, cf_inv_out=cf_inv_out, cf_fin=cf_fin,
    cf_fin_in=cf_fin_in, cf_fin_out=cf_fin_out, cf_net=cf_net,
    cash_begin=cash_begin, cash_end=cash_end, indirect=indirect,
    dep_total=dep_total, dep_admin=dep_admin, dep_sales=dep_sales,
    pay_total=pay_total, pay_unit=pay_unit, pay_admin=pay_admin,
    pay_sales=pay_sales, pay_tax=pay_tax, pay_net=pay_net,
    out_vat=out_vat, in_vat=in_vat, vat_payable=vat_payable,
    city=city, edu=edu, edu2=edu2, stamp=stamp, surcharge=surcharge,
    tax_surcharge=tax_surcharge,
    ar_close=AR_CLOSE, ap_close=AP_CLOSE, close_aging=CLOSE_AGING,
    bank_flow=BANK_FLOW, bank_close=bank_close, stmt=stmt, stmt_close=stmt_close,
    payroll=PAYROLL, dep=DEP, vouchers=VOUCHERS,
    d_ar=d_ar, d_inv=d_inv, d_prepaid=d_prepaid, d_ap=d_ap,
    d_pay=d_pay, d_tax=d_tax,
)
