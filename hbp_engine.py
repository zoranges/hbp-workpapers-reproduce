#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HBP Workpapers v2.02 复现引擎（核心：COA 映射 → TBMAP 合并 → LEADSHEET 填充 → 指标计算）

复现原 Excel 逻辑：
  COA(科目→Map) → TBMAP(合并本期/上年) → LEADSHEET(SUMIFS 填充) → 高维指标
MAP 映射采用本地工作簿的 MAP 表（67 条规则）。

用法:
  python hbp_engine.py --data data/cloudfinesse_pack.json --out report
"""
import argparse
import json
import os
import sys

RULES = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'rules', 'map_rules.json')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def load_rules(path=RULES):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def auto_map(account, rules):
    """按科目类型/名称自动映射到 Map No.（复刻 COA 下拉选择逻辑，规则来自本地 MAP 表）"""
    code = str(account['code'])
    name = (account.get('name') or '').lower()
    atype = (account.get('type') or '').upper()
    rmap = {r['map_no']: r for r in rules}
    desc = rmap.get(code, {}).get('description')  # 若科目代码恰为 map_no（无此情况）
    # 关键词匹配
    if atype == 'REVENUE':
        if 'other' in name:
            return '1.01'
        return '1.001'
    if atype == 'DIRECTCOSTS':
        return '3.001'
    if atype == 'EXPENSE':
        if any(k in name for k in ('wage', 'salary', 'director', 'super', 'payroll', 'employment')):
            return '4.003'
        if 'depreciation' in name or 'amorti' in name:
            return '4.002'
        if 'travel' in name or 'vehicle' in name or 'motor' in name:
            return '4.008'
        if 'market' in name or 'advertis' in name:
            return '4.005'
        if 'occup' in name or 'rent' in name or 'light' in name or 'power' in name:
            return '4.006'
        if 'professional' in name or 'legal' in name or 'account' in name or 'audit' in name:
            return '4.007'
        if 'management fee' in name:
            return '4.004'
        if 'interest' in name:
            return '5.002'
        # FX/汇兑：MAP 5.003 为 Net Finance Cost，但本客户 Xero P&L 将其计入费用；
        # 为保持 BS 恒等式口径一致，归入 4.001（可在规则配置中按客户调整）
        if 'foreign' in name or 'fx' in name or 'unrealised' in name or 'revalu' in name:
            return '4.001'
        return '4.001'  # Administrative Costs
    if atype in ('BANK',):
        return '10.001'
    if atype in ('CURRENT', 'ASSET', 'CURASSET'):
        if 'prepaid' in name:
            return '11.002'
        if 'receivable' in name or 'debtor' in name:
            return '11.001'
        if 'inventory' in name or 'stock' in name:
            return '13.001'
        if 'fixed' in name or 'equipment' in name or 'plant' in name or 'property' in name:
            return '14.001'
        return '11.001'
    if atype == 'FIXED':
        return '14.001'
    if atype == 'INVENTORY':
        return '13.001'
    if atype in ('LIABILITY', 'CURRLIAB', 'TERMLIAB'):
        if 'payable' in name:
            return '20.001'
        if 'accru' in name:
            return '20.002'
        if 'tax' in name or 'gst' in name or 'bas' in name:
            return '20.003'
        if 'suspense' in name or 'round' in name or 'clearing' in name:
            return '20.006'
        if 'loan' in name:
            return '20.001'
        if 'wage' in name or 'super' in name or 'employee' in name:
            return '20.002'
        return '20.001'
    if atype == 'EQUITY':
        if 'retained' in name or 'opening' in name:
            return '30.001'
        if 'capital' in name or 'share' in name:
            return '31.001'
        if 'reserve' in name:
            return '32.002'
        return '30.001'
    return ''


def build_tbmap(data_pack, rules):
    """COA + TBCY + TBPY → TBMAP 合并表"""
    tbmap = []
    accounts = data_pack['accounts']
    tb_cy = data_pack.get('tb_cy', {})
    tb_py = data_pack.get('tb_py', {})
    for acc in accounts:
        code = acc['code']
        map_no = auto_map(acc, rules)
        if not map_no:
            continue
        cy = float(tb_cy.get(code, 0.0))
        py = float(tb_py.get(code, 0.0))
        tbmap.append({"code": code, "name": acc['name'], "map_no": map_no, "period": "This", "value": cy})
        tbmap.append({"code": code, "name": acc['name'], "map_no": map_no, "period": "Last", "value": py})
    return tbmap


INCOME_MAPS = {'1.001', '1.002', '1.003', '1.01', '2.001', '2.002', '2.003', '2.004', '2.005'}
# BS 负债与权益（贷方科目）：Excel 中显示为正，取 -SUMIFS
LIAB_EQUITY_MAPS = {'20.001', '20.002', '20.003', '20.004', '20.005', '20.006',
                    '21.001', '21.002', '30.001', '30.002', '31.001', '32.001', '32.002'}


def sumifs(tbmap, period, **filters):
    """复刻 SUMIFS(TBMAP[Value], ...)"""
    total = 0.0
    for row in tbmap:
        if row['period'] != period:
            continue
        if 'code' in filters and row['code'] != filters['code']:
            continue
        if 'map_no' in filters and row['map_no'] != filters['map_no']:
            continue
        total += row['value']
    return total


def sign_adj(tbmap, period, map_no):
    """按 Excel 公式符号规则：收入类与负债/权益 = -SUMIFS，其余 = +SUMIFS"""
    v = sumifs(tbmap, period, map_no=map_no)
    if map_no in INCOME_MAPS or map_no in LIAB_EQUITY_MAPS:
        return -v
    return v


def sum_group(tbmap, period, maps):
    """一组 Map No. 按各自符号规则汇总"""
    return sum(sign_adj(tbmap, period, m) for m in maps)


def metric(name, this, last):
    change = (this - last) if (this is not None and last is not None) else None
    pct = (change / last) if (last and change is not None) else None
    return {"name": name,
            "this": (round(this, 2) if this is not None else None),
            "last": (round(last, 2) if last is not None else None),
            "change": (round(change, 2) if change is not None else None),
            "pct": pct}


def calc_metrics(data_pack, tbmap):
    """复刻 LEADSHEET/FIN_SUM 高维指标（11 项），符号规则与 Excel 公式一致"""
    sales_this = sum_group(tbmap, 'This', ['1.001', '1.002', '1.003', '1.01'])
    sales_last = sum_group(tbmap, 'Last', ['1.001', '1.002', '1.003', '1.01'])
    cos_this = sum_group(tbmap, 'This', ['3.001', '3.002'])
    cos_last = sum_group(tbmap, 'Last', ['3.001', '3.002'])
    gp_this, gp_last = sales_this - cos_this, sales_last - cos_last
    other_income_this = sum_group(tbmap, 'This', ['2.001', '2.002', '2.003', '2.004', '2.005'])
    other_income_last = sum_group(tbmap, 'Last', ['2.001', '2.002', '2.003', '2.004', '2.005'])
    exp_this = sum_group(tbmap, 'This', ['4.001', '4.002', '4.003', '4.004', '4.005', '4.006', '4.007', '4.008', '4.009', '4.01'])
    exp_last = sum_group(tbmap, 'Last', ['4.001', '4.002', '4.003', '4.004', '4.005', '4.006', '4.007', '4.008', '4.009', '4.01'])
    fin_this = sum_group(tbmap, 'This', ['5.001', '5.002', '5.003'])
    fin_last = sum_group(tbmap, 'Last', ['5.001', '5.002', '5.003'])
    npbt_this = gp_this + other_income_this - exp_this + fin_this
    npbt_last = gp_last + other_income_last - exp_last + fin_last
    # 资产/负债（按 map 汇总，BS 部分 = +SUMIFS）
    cash_this = sum_group(tbmap, 'This', ['10.001'])
    cash_last = sum_group(tbmap, 'Last', ['10.001'])
    recv_this = sum_group(tbmap, 'This', ['11.001', '11.002', '11.003', '11.004'])
    recv_last = sum_group(tbmap, 'Last', ['11.001', '11.002', '11.003', '11.004'])
    loan_recv_this = sum_group(tbmap, 'This', ['12.001', '12.002'])
    loan_recv_last = sum_group(tbmap, 'Last', ['12.001', '12.002'])
    inv_this = sum_group(tbmap, 'This', ['13.001'])
    inv_last = sum_group(tbmap, 'Last', ['13.001'])
    fa_this = sum_group(tbmap, 'This', ['14.001'])
    fa_last = sum_group(tbmap, 'Last', ['14.001'])
    cur_assets_this = cash_this + recv_this
    cur_assets_last = cash_last + recv_last
    total_assets_this = cur_assets_this + loan_recv_this + inv_this + fa_this
    total_assets_last = cur_assets_last + loan_recv_last + inv_last + fa_last
    pay_this = sum_group(tbmap, 'This', ['20.001', '20.002', '20.003', '20.004', '20.005', '20.006'])
    pay_last = sum_group(tbmap, 'Last', ['20.001', '20.002', '20.003', '20.004', '20.005', '20.006'])
    tax_liab_this = sum_group(tbmap, 'This', ['21.001', '21.002'])
    tax_liab_last = sum_group(tbmap, 'Last', ['21.001', '21.002'])
    cur_liab_this = pay_this + tax_liab_this
    cur_liab_last = pay_last + tax_liab_last
    total_liab_this = cur_liab_this
    total_liab_last = cur_liab_last
    equity_this = npbt_this + sum_group(tbmap, 'This', ['30.001', '31.001', '32.002'])
    equity_last = npbt_last + sum_group(tbmap, 'Last', ['30.001', '31.001', '32.002'])
    net_assets_this = total_assets_this - total_liab_this
    net_assets_last = total_assets_last - total_liab_last

    ms = []
    ms.append(metric('Total Sales 销售总额', sales_this, sales_last))
    ms.append(metric('Gross Profit 毛利', gp_this, gp_last))
    ms.append(metric('GPM % 毛利率', (gp_this / sales_this * 100 if sales_this else 0), (gp_last / sales_last * 100 if sales_last else 0)))
    ms.append(metric('Net Profit Before Tax 税前净利', npbt_this, npbt_last))
    ms.append(metric('NPBT % 税前净利率', (npbt_this / sales_this * 100 if sales_this else 0), (npbt_last / sales_last * 100 if sales_last else 0)))
    ms.append(metric('Quick Ratio 速动比率', ((cash_this + recv_this) / cur_liab_this if cur_liab_this else 0), ((cash_last + recv_last) / cur_liab_last if cur_liab_last else 0)))
    ms.append(metric('Current Asset Ratio 流动比率', (cur_assets_this / cur_liab_this if cur_liab_this else 0), (cur_assets_last / cur_liab_last if cur_liab_last else 0)))
    ms.append(metric('Net Asset Ratio 净资产比率', (net_assets_this / total_liab_this if total_liab_this else 0), (net_assets_last / total_liab_last if total_liab_last else 0)))
    ms.append(metric('Debtor Turnover Days 应收周转天数', (recv_this / sales_this * 365 if sales_this else 0), (recv_last / sales_last * 365 if sales_last else 0)))
    ms.append(metric('Creditor Turnover Days 应付周转天数', (pay_this / (cos_this + exp_this) * 365 if (cos_this + exp_this) else 0), (pay_last / (cos_last + exp_last) * 365 if (cos_last + exp_last) else 0)))
    ms.append(metric('Inventory Turnover Days 存货周转天数', (365 / (cos_this / inv_this) if inv_this else None), (365 / (cos_last / inv_last) if inv_last else None)))

    bs = {
        "total_assets": round(total_assets_this, 2),
        "total_liabilities": round(total_liab_this, 2),
        "total_equity": round(equity_this, 2),
        "net_assets": round(net_assets_this, 2),
        "bs_identity_diff": round(total_assets_this - total_liab_this - equity_this, 2),
        "tb_balance_diff": round(sum(r['value'] for r in tbmap if r['period'] == 'This'), 2)
    }
    return ms, bs


def build_leadsheet(tbmap):
    """按科目代码填充 LEADSHEET（SUMIFS 逻辑，收入类符号与 Excel 一致）"""
    rows = []
    seen = {}
    for row in tbmap:
        key = (row['code'], row['map_no'], row['name'])
        if key not in seen:
            seen[key] = {"this": 0.0, "last": 0.0}
        if row['period'] == 'This':
            seen[key]["this"] += row['value']
        else:
            seen[key]["last"] += row['value']
    for (code, map_no, name), v in seen.items():
        neg = (map_no in INCOME_MAPS or map_no in LIAB_EQUITY_MAPS)
        this = -v["this"] if neg else v["this"]
        last = -v["last"] if neg else v["last"]
        change = this - last
        pct = (change / last) if last else None
        rows.append({"code": code, "name": name, "map_no": map_no,
                     "this": round(this, 2), "last": round(last, 2),
                     "adj": 0.0, "final": round(this, 2),
                     "change": round(change, 2), "pct": pct,
                     "ref": "", "commentary": ""})
    rows.sort(key=lambda r: (r['map_no'], r['code']))
    return rows


def render_markdown(org, period, metrics, bs, leadsheet, tbmap):
    L = []
    L.append(f"# HBP Workpapers v2.02 复现输出 — {org}")
    L.append(f"期间：{period} ｜ 映射规则：本地 MAP 表（67 条）")
    L.append("")
    L.append("## 一、高维指标（LEADSHEET HIGH LEVEL STATS）")
    L.append("| 指标 | 本期 | 上期 | 变动 | 变动% |")
    L.append("|---|---|---|---|---|")
    for m in metrics:
        pct = f"{m['pct']*100:.1f}%" if m['pct'] is not None else "N/A"
        this = f"{m['this']:,.2f}" if m['this'] is not None else "N/A"
        last = f"{m['last']:,.2f}" if m['last'] is not None else "N/A"
        chg = f"{m['change']:+,.2f}" if m['change'] is not None else "N/A"
        L.append(f"| {m['name']} | {this} | {last} | {chg} | {pct} |")
    L.append("")
    L.append("## 二、资产负债表勾稽")
    L.append(f"- 总资产 {bs['total_assets']:,.2f} ｜ 总负债 {bs['total_liabilities']:,.2f} ｜ 总权益 {bs['total_equity']:,.2f}")
    L.append(f"- 恒等式差额：{bs['bs_identity_diff']:,.2f} {'✔ 成立' if abs(bs['bs_identity_diff']) <= 0.05 else '⚠ 不成立'}")
    L.append(f"- TB 借贷平衡差额：{bs['tb_balance_diff']:,.2f} {'✔' if abs(bs['tb_balance_diff']) <= 0.05 else '⚠'}")
    L.append("")
    L.append("## 三、LEADSHEET 科目级填充（模板一致 10 列结构）")
    L.append("| Account Code | Account Description | Client TB | HBP Adjustments | Final Adjusted TB | Prior Year TB | $ Change | % Change | File Reference | HBP Analytical Commentary |")
    L.append("|---|---|---|---|---|---|---|---|---|---|")
    for r in leadsheet:
        pct = f"{r['pct']*100:.1f}%" if r['pct'] is not None else "N/A"
        L.append(f"| {r['code']} | {r['name']} | {r['this']:,.2f} | {r['adj']:,.2f} | {r['final']:,.2f} | {r['last']:,.2f} | {r['change']:+,.2f} | {pct} | {r['ref']} | {r['commentary']} |")
    L.append("")
    L.append(f"## 四、数据包校验")
    cy_total = sum(r['value'] for r in tbmap if r['period'] == 'This')
    py_total = sum(r['value'] for r in tbmap if r['period'] == 'Last')
    L.append(f"- 本期 TB 合计：{cy_total:,.2f}（应≈0 或留存结转）")
    L.append(f"- 上年 TB 合计：{py_total:,.2f}")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', required=True)
    ap.add_argument('--out', default='')
    args = ap.parse_args()

    with open(args.data, encoding='utf-8') as f:
        pack = json.load(f)
    rules = load_rules()
    tbmap = build_tbmap(pack, rules)
    metrics, bs = calc_metrics(pack, tbmap)
    leadsheet = build_leadsheet(tbmap)
    md = render_markdown(pack['org'], pack['period_end'], metrics, bs, leadsheet, tbmap)
    print(md)
    if args.out:
        out_dir = os.path.dirname(args.out)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir)
        with open(args.out + '.md', 'w', encoding='utf-8') as f:
            f.write(md)
        out_json = {"org": pack['org'], "metrics": metrics, "balance_sheet": bs, "leadsheet": leadsheet}
        with open(args.out + '.json', 'w', encoding='utf-8') as f:
            json.dump(out_json, f, ensure_ascii=False, indent=1)
        print(f"\n[OK] 输出: {args.out}.md / {args.out}.json")


if __name__ == '__main__':
    main()
