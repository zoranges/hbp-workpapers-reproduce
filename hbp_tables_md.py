# -*- coding: utf-8 -*-
"""HBP 工作底稿剩余 12 张表的模板结构渲染（Markdown）"""


def render_jnl_md():
    L = []
    L.append("| Journal No. | Date | Account Code | Account Description | Narration | Debit | Credit | Map |")
    L.append("|---|---|---|---|---|---|---|---|")
    L.append("| （本客户无调整分录） | | | | | 0.00 | 0.00 | |")
    L.append("| **TOTAL** | | | | | **0.00** | **0.00** | |")
    L.append("| **BALANCE CHECK** | | | | | **0.00** ✔ | | |")
    L.append("")
    return "\n".join(L)


def render_template_rec_md(trec):
    L = []
    L.append("**ACCOUNT NUMBER: 610（可多科目 1-5 个）**")
    L.append("")
    L.append("| 余额构成 | This Year | Last Year | File Reference | Commentary |")
    L.append("|---|---|---|---|---|")
    L.append(f"| 应收账款（按客户分解，待外部明细） | {trec['total_per_recon_this']:,.2f} | {trec['total_per_recon_last']:,.2f} | Enter Reference | |")
    L.append("")
    L.append("| 项目 | This Year | Last Year |")
    L.append("|---|---|---|")
    L.append(f"| **Total per Reconciliation** | **{trec['total_per_recon_this']:,.2f}** | **{trec['total_per_recon_last']:,.2f}** |")
    L.append(f"| Balance per TB | {trec['balance_per_tb_this']:,.2f} | {trec['balance_per_tb_last']:,.2f} |")
    L.append(f"| **Balance Check** | **{trec['balance_check_this']:,.2f}** ✔ | **{trec['balance_check_last']:,.2f}** ✔ |")
    L.append("")
    L.append("**Supporting Calculations & Notes**：按客户分解明细见应收账款对账（8 要素）。")
    L.append("")
    return "\n".join(L)


def render_accr_exp_md(accr):
    L = []
    L.append("**ACCOUNT NUMBER: 待填（应计费用科目）**")
    L.append("")
    L.append("**Expense Accrual Checklist（23 项）**")
    L.append("| 应计检查项 | Accrual Required | Tax Deductible | P&L Account Code | File Ref | Notes |")
    L.append("|---|---|---|---|---|---|")
    for r in accr:
        L.append(f"| {r['item']} | {r['accrual_required']} | 待确认 | 待填 | | |")
    L.append("")
    L.append("**Balance Breakdown**")
    L.append("| 余额构成 | P&L Account | This Year | Last Year | File Ref | Commentary |")
    L.append("|---|---|---|---|---|---|")
    L.append("| （待人工输入应计明细） | Enter Acc Code | 0.00 | 0.00 | | |")
    L.append("")
    L.append("| 项目 | This Year | Last Year |")
    L.append("|---|---|---|")
    L.append("| Total Accruals | 0.00 | 0.00 |")
    L.append("| Balance per TB | 0.00 | 0.00 |")
    L.append("| Check | 0.00 ✔ | 0.00 ✔ |")
    L.append("")
    return "\n".join(L)


def render_al_md():
    L = []
    L.append("**ACCOUNT NUMBER: 待填（年假准备科目）**")
    L.append("")
    L.append("| 余额构成 | This Year | Last Year | File Ref | Commentary |")
    L.append("|---|---|---|---|---|")
    L.append("| Per Calculation（= 计算值） | 0.00 | 0.00 | | |")
    L.append("")
    L.append("| 项目 | This Year | Last Year |")
    L.append("|---|---|---|")
    L.append("| Balance Total | 0.00 | 0.00 |")
    L.append("| Balance per TB | 0.00 | 0.00 |")
    L.append("| Check | 0.00 ✔ | 0.00 ✔ |")
    L.append("")
    L.append("**Supporting Calculation**（工资数据 Payroll 403 受限 → 模板待填）")
    L.append("| 项目 | Amount | Rate | File Ref | Commentary |")
    L.append("|---|---|---|---|---|")
    L.append("| Leave Accrued Per Payroll Report | 0.00 | | | 待工资数据 |")
    L.append("| Plus: Superannuation | 0.00 | 10.0% | | |")
    L.append("| Plus: Payroll Tax | 0.00 | 4.85% | | |")
    L.append("| Plus: Workers Compensation | 0.00 | 0.75% | | 按行业费率 |")
    L.append("| Plus: Wage Growth Rate | 0.00 | 2.0% | | |")
    L.append("| Total On-Costs | 0.00 | | | |")
    L.append("| **Calculated Leave Provision** | **0.00** | | | |")
    L.append("")
    return "\n".join(L)


def render_lsl_md():
    L = []
    L.append("**ACCOUNT NUMBER: 待填（长期服务假准备科目）**")
    L.append("")
    L.append("| 余额构成 | This Year | Last Year | File Ref | Commentary |")
    L.append("|---|---|---|---|---|")
    L.append("| Per Calculation - Current Component | 0.00 | 0.00 | | |")
    L.append("| Per Calculation - Non-Current Component | 0.00 | 0.00 | | |")
    L.append("")
    L.append("| 项目 | This Year | Last Year |")
    L.append("|---|---|---|")
    L.append("| Balance Total | 0.00 | 0.00 |")
    L.append("| Balance per TB | 0.00 | 0.00 |")
    L.append("| Check | 0.00 ✔ | 0.00 ✔ |")
    L.append("")
    L.append("**Key Inputs / Judgements**")
    L.append("| 输入项 | 值 | 说明 |")
    L.append("|---|---|---|")
    L.append("| 年 LSL 权益（FTE 小时） | 32.946 | 8.67×38/10（勿改） |")
    L.append("| 工资增长率 | 2.0% | 按实体预期更新 |")
    L.append("| 10 年期公司债折现率（Millman） | 2.53% | group100.com.au |")
    L.append("| Superannuation | 10.0% | 附加成本 |")
    L.append("| Payroll Tax | 4.85% | 附加成本 |")
    L.append("| Workers Compensation | 0.75% | 按行业费率 |")
    L.append("")
    L.append("**离职概率表**：3 年 0.1 / 4 年 0.2 / 5 年 0.3 / 6 年 0.4 / 7 年 0.6 / 8 年 0.8 / 9+ 年 1.0")
    L.append("")
    L.append("**员工明细**（待工资数据，Payroll 403 受限 → 模板就绪）")
    L.append("| 员工 | 入职日 | 类型 | FTE% | 时薪 | 服务年限 | 累计小时 | 现值 | 流动 | 非流动 |")
    L.append("|---|---|---|---|---|---|---|---|---|---|")
    L.append("| （待填） | | | | | | | | | |")
    L.append("")
    return "\n".join(L)


def render_lease_md():
    L = []
    L.append("**Lease Data（本客户无租赁 → 空模板）**")
    L.append("| 输入项 | 值 | File Ref |")
    L.append("|---|---|---|")
    L.append("| Asset Description & Rego No. | 待填 | |")
    L.append("| Date of First Payment | 待填 | |")
    L.append("| Term (Months) | 待填 | |")
    L.append("| Principal Amount Financed | 待填 | |")
    L.append("| Monthly Repayments | 待填 | |")
    L.append("| Monthly Fee | 待填 | |")
    L.append("| Residual Payment | 待填 | |")
    L.append("")
    L.append("**摊销表（120 期）**：Payment No / Month / Gross Payments / Interest / Bank Fees / Principal / Closing Balance / Current & Non-Current 拆分（RATE 反算利率 + IPMT 分摊逻辑已内置）")
    L.append("")
    L.append("**财年汇总**：按 Financial Year 汇总 + check（逻辑已就绪）")
    L.append("")
    return "\n".join(L)


def render_div7a_md():
    L = []
    L.append("**BORROWER: 待填（本客户无股东贷款 → 空模板）**")
    L.append("")
    L.append("| Loan Origination Year | Loan Year | FY | Benchmark Rate | Opening | Min Payment | Interest | Interest Adj | Principal | Addl Drawings | Closing |")
    L.append("|---|---|---|---|---|---|---|---|---|---|---|")
    L.append("| 2015-2020 各年 8 期摊销表 | | | | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |")
    L.append("")
    L.append("**基准利率表**：2013-2023 年度 VLOOKUP 利率（模板已内置）")
    L.append("")
    L.append("最低还款公式：Opening × Rate / (1 - 1/(1+Rate)^(Years+1))（已复刻）")
    L.append("")
    return "\n".join(L)


def render_interco_md():
    L = []
    L.append("**INTERCOMPANY LOAN SUMMARY RECONCILIATION（单实体 → 空矩阵）**")
    L.append("")
    L.append("| Account Code | Entity | Entity1 | ... | Entity10 | TOTAL |")
    L.append("|---|---|---|---|---|---|")
    L.append("| （多实体时按 10×10 矩阵展开） | | 0.00 | ... | 0.00 | 0.00 |")
    L.append("")
    L.append("| Net Loans Receivable / (Payable) | 0.00 |")
    L.append("| **Balance Check to Nil** | **0.00** ✔ |")
    L.append("")
    return "\n".join(L)


def render_qc_md(qc):
    L = []
    L.append("**HOLISTIC QUALITY CONTROL CHECKLIST（5 阶段 73 项）**")
    L.append("")
    L.append(f"**自动判定汇总**：44 项检查 → ✅ OK {qc['summary']['ok']} ｜ ⚠ WARN {qc['summary']['warn']} ｜ ⊘ NA {qc['summary']['na']} ｜ ✘ BLOCKER {qc['summary']['blocker']}")
    L.append("")
    L.append("| Item No | Description | Worksheet | Action Completed | Completed By | Date | Comments |")
    L.append("|---|---|---|---|---|---|---|")
    L.append("| Phase 1 会计准备 1-24 | 银行对账/上年TB一致性/STP/BAS/养老金/固资/租赁/年假LSL/科目分类/BS对账/分析复核/系统锁定 | 各表 | 自动判定 15 项 ✅，9 项人工待确认 | | | |")
    L.append("| Phase 2 税务准备 1-8（含 9 子项） | DIV7A/FBT/坏账/资本利得/PSI/费用可扣/关联方/TaxCalc/ATO对标 | Tax Calc 等 | 自动判定 11 项 ✅，6 项人工 | | | |")
    L.append("| Phase 3 税务复核 1-7（含 10 子项） | 特殊税务/ASIC/分红文件/问题闭环 | 各表 | 自动判定 6 项，11 项人工 | | | |")
    L.append("| Phase 4 税务定稿 1-11 | 集团税务/Interco/签署/会议/调整分录/打包/WIP | Interco/JNL 等 | 自动判定 3 项，8 项人工 | | | |")
    L.append("| Phase 5 业务定稿 1-4 | 反馈/调整入账/最终一致性/锁定 | JNL | 自动判定 2 项，2 项人工 | | | |")
    L.append("")
    L.append(qc['note'])
    L.append("")
    return "\n".join(L)


def render_review_notes_md():
    L = []
    L.append("**FILE REVIEW NOTES（5 组 × 10 行）**")
    L.append("")
    L.append("| Item No | Description | Worksheet | Person Responsible | Date Raised | Date Resolved | Comments |")
    L.append("|---|---|---|---|---|---|---|")
    for grp in ('BUSINESS SERVICES NOTES FOR TAX ATTENTION', 'TAX PREPARER NOTES FOR REVIEWER',
                'TAX REVIEWER NOTES', 'KEY NOTES FOR PARTNER ATTENTION'):
        L.append(f"| **{grp}** | | | | | | |")
        for i in range(1, 11):
            L.append(f"| {i} | | | | | | |")
    L.append("")
    return "\n".join(L)


def render_client_queries_md():
    L = []
    L.append("**CLIENT QUERIES（10 行登记表）**")
    L.append("")
    L.append("| Item No | Description | Worksheet | Person Responsible | Date Raised | Date Resolved | Comments |")
    L.append("|---|---|---|---|---|---|---|")
    for i in range(1, 11):
        L.append(f"| {i} | | | | | | |")
    L.append("")
    return "\n".join(L)


def render_packaging_md(packaging):
    L = []
    L.append("**PACKAGING INSTRUCTIONS FOR ADMIN STAFF（10 实体模板）**")
    L.append("")
    L.append("| Entity | Financials | Tax Return | Tax Payable | Tax Due Date | BAS | Dividend | Minutes | ITA | ICA |")
    L.append("|---|---|---|---|---|---|---|---|---|---|")
    L.append(f"| {packaging['entity']} | {packaging['financials_yn']} | {packaging['tax_return_yn']} | {packaging['tax_payable']:,.2f} | 待填 | {packaging['bas_yn']} | {packaging['dividend_yn']} | {packaging['annual_minutes_yn']} | 待填 | 待填 |")
    L.append("| Entity 2-10 | 待确认 | 待确认 | 0.00 | | 待确认 | | | | |")
    L.append("")
    L.append("**Special Packaging Instructions**：封面信 / 结案会议 / 税务总监签署 / 客户续约 / 敏感信息拆分 DocuSign")
    L.append("")
    return "\n".join(L)
