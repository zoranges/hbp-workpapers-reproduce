# HBP 工作底稿 v2.02 全表复现引擎 — 实施文档

> 版本：1.0 ｜ 日期：2026-08-11 ｜ 数据源：Xero MCP（mcp_xero2 · Cloudfinesse Pte Ltd）
> 目标：从 Xero 数据出发，使用 HBP 工作底稿（22 表 Excel 模板）中的全部公式与逻辑，**完整复现 19 张输出表**（11 数值 + 4 状态 + 4 中间/规则）。

---

## 1. 背景与目标

### 1.1 问题
HBP 年度工作底稿是 22 张工作表的 Excel 模板，含 **7,216 个公式**。会计师年结时手工执行：粘贴科目表与 TB → Power Query 合并 → LEADSHEET 自动填充 → 对账 → 税务调节 → 质量控制。

### 1.2 目标
- 从 Xero MCP 自动获取全部输入数据（科目表、本期/上期 TB）
- 使用工作簿内 **MAP 73 行映射规则**（本地化）自动归类科目
- 用 Python 复刻全部公式逻辑（SUMIFS/指标/税务调节/对账）
- 输出与原模板**结构一致**的 19 张表（Markdown + Excel）

### 1.3 已达成（验证结果）
- 恒等式差额 **0.0**（总资产 = 总负债 + 总权益）
- TB 借贷平衡 **0.0**
- 60 科目全部映射（MAP_ERR 无未映射）
- 高维指标与 Xero P&L 一致（Sales 39,300.74 / GP 34,863.74 / NPBT 16,355.98）

---

## 2. 架构总览

```
┌──────────────────────────────────────────────────────────────┐
│  输入层：数据包 JSON（data/cloudfinesse_pack.json）            │
│    org(组织) + accounts(科目表) + tb_cy(本期TB) + tb_py(上期TB) │
└────────────────────────────┬─────────────────────────────────┘
                             ▼
┌──────────────────────────────────────────────────────────────┐
│  规则层（rules/map_rules.json · 本地 MAP 表 67 条）            │
│    Map No. → 描述 / FIN Class / FIN Category / CF Class       │
└────────────────────────────┬─────────────────────────────────┘
                             ▼
┌──────────────────────────────────────────────────────────────┐
│  计算层（hbp_engine.py 核心 + hbp_reproduce_full.py 全表）     │
│    ① 科目自动映射（auto_map）                                  │
│    ② TBMAP 合并（科目×期间×金额）                              │
│    ③ LEADSHEET 填充（SUMIFS + 符号规则）                       │
│    ④ 11 项高维指标                                             │
│    ⑤ FIN_SUM（模板行序 + 双校验）                              │
│    ⑥ Tax Calc（16 区段完整结构）                               │
│    ⑦ 其余表（JNL/Template Rec/Accr/AL/LSL/Lease/DIV7A/        │
│       Interco/QC/Packaging/Review/Client）                     │
└────────────────────────────┬─────────────────────────────────┘
                             ▼
┌──────────────────────────────────────────────────────────────┐
│  输出层                                                     │
│    Markdown：output/hbp_full_report.md（19 张表）              │
│    Excel：   output/hbp_output.xlsx（15 个 sheet）             │
│    JSON：    output/cloudfinesse.json（结构化数据）             │
└──────────────────────────────────────────────────────────────┘
```

### 文件结构

```
hbp_workpapers_reproduce/
├── hbp_engine.py              # 核心引擎：映射/TBMAP/LEADSHEET/指标
├── hbp_reproduce_full.py      # 全表生成器：19 张表 + Excel 输出
├── hbp_tables_md.py           # 12 张表的模板渲染函数（Markdown）
├── data/
│   └── cloudfinesse_pack.json # 输入数据包（真实 Xero 数据）
├── rules/
│   └── map_rules.json         # 本地 MAP 表（67 条规则，从工作簿提取）
└── output/
    ├── hbp_full_report.md     # 19 张表完整报告
    ├── hbp_output.xlsx        # 15 个 sheet 的 Excel 工作簿
    └── cloudfinesse.json      # 结构化输出
```

---

## 3. 数据契约（输入数据包 Schema）

所有输入封装为一个 JSON 数据包，字段如下：

```json
{
  "org": "Cloudfinesse Pte Ltd",          // 组织名称
  "currency": "SGD",                       // 币种
  "period_end": "2026-08-31",              // 本期截止
  "prior_period_end": "2025-12-31",        // 上期截止
  "accounts": [                            // 科目表（list_accounts）
    {"code": "200", "name": "Sales", "type": "REVENUE"},
    {"code": "800", "name": "Accounts Payable", "type": "CURRLIAB"}
  ],
  "tb_cy": {"200": -39249.00, "100": 9159.08},  // 本期 TB YTD（贷方为负）
  "tb_py": {"200": -25436.00, "100": 14414.42}   // 上期 TB YTD（贷方为负）
}
```

### 符号约定（关键）

| 科目类别 | Xero TB YTD 位置 | 数据包符号 |
|---|---|---|
| 收入 | YTD Credit（正） | **负**（如 Sales -39249） |
| 费用 | YTD Debit（正） | 正（如 Bank Fees +67.59） |
| 资产 | YTD Debit（正） | 正（如 AR +21275.50） |
| 负债 | YTD Credit（正） | **负**（如 AP -3839.00） |
| 权益 | YTD Credit（正） | **负**（如 RE -9990.68） |

即：`value = YTD Debit − YTD Credit`（借方净额为正、贷方净额为负）。这是复刻 Excel 公式符号的基础。

---

## 4. 核心逻辑实现

### 4.1 科目自动映射（auto_map）

复刻 COA 表的"Map Description 下拉选择"逻辑。规则来源：本地 MAP 表 + 科目类型 + 名称关键词。

```
REVENUE      → 'other' in name ? 1.01 : 1.001
DIRECTCOSTS  → 3.001
EXPENSE      → 关键词匹配（按序）：
               wage/salary/director/super/payroll → 4.003 雇佣成本
               depreciation/amorti                → 4.002 折旧摊销
               travel/vehicle/motor               → 4.008 车辆差旅
               market/advertis                    → 4.005 营销
               occup/rent/light/power             → 4.006 占用
               professional/legal/account/audit   → 4.007 专业费
               management fee                     → 4.004 管理费
               interest                           → 5.002 利息费用
               foreign/fx/unrealised/revalu       → 4.001 行政费用（口径说明见 5.2）
               其余                               → 4.001 行政费用
BANK         → 10.001
CURRENT/ASSET→ prepaid ? 11.002 : (receivable ? 11.001 : (inventory ? 13.001 : ...))
FIXED        → 14.001
INVENTORY    → 13.001
CURRLIAB/TERMLIAB/LIABILITY → payable ? 20.001 : (accrued ? 20.002 :
               (tax/gst/bas ? 20.003 : (suspense/rounding/clearing ? 20.006 : 20.001)))
EQUITY       → retained/opening ? 30.001 : (capital/share ? 31.001 : (reserve ? 32.002 : 30.001))
```

**注意（经验教训）**：Xero 的真实科目类型是 `CURRLIAB`/`TERMLIAB`/`FIXED`/`INVENTORY` 等，不是 `LIABILITY`/`ASSET`。若不处理这些类型，负债类科目会静默丢失（曾导致恒等式差 3,839.02）。所有 Xero 类型必须显式覆盖。

### 4.2 符号规则（sign_adj）

复刻 Excel 各行公式的符号（LEADSHEET/FIN_SUM）：

```
INCOME_MAPS（1.001~1.01, 2.001~2.005）   → 取 -SUMIFS（收入显示为正）
LIAB_EQUITY_MAPS（20.x, 21.x, 30.x, 31.x, 32.x） → 取 -SUMIFS（负债/权益显示为正）
其余（费用/资产）                          → 取 +SUMIFS
```

### 4.3 TBMAP 合并（build_tbmap）

将 COA + TBCY + TBPY 展开为明细行（对应 Excel 的 Power Query 合并）：

```python
tbmap = [
  {"code": "200", "name": "Sales", "map_no": "1.001", "period": "This", "value": -39249.0},
  {"code": "200", "name": "Sales", "map_no": "1.001", "period": "Last", "value": -25436.0},
  ...
]
```

### 4.4 LEADSHEET 填充（build_leadsheet + sumifs）

按科目代码从 TBMAP 汇总（复刻 `SUMIFS(TBMAP[Value], TBMAP[Account Code], $A23, TBMAP[Period], C$19)`）：

```python
def sumifs(tbmap, period, **filters):   # 等价 SUMIFS
    # 按 code/map_no 过滤后求和

def build_leadsheet(tbmap):
    # 按 (code, map_no, name) 分组 → 本期/上期/变动额/变动%
```

输出行结构：`科目代码 | 科目名称 | Map No. | 本期 | 上期 | 变动 | 变动%`

### 4.5 高维指标（calc_metrics）

复刻 LEADSHEET/FIN_SUM 顶部 11 项指标公式：

| 指标 | 公式（引擎实现） |
|---|---|
| Total Sales | Σ(1.001+1.002+1.003+1.01) 符号调整 |
| Gross Profit | Sales − Direct Costs |
| GPM % | GP / Sales × 100 |
| Net Profit Before Tax | GP + Other Income − Expenses + Net Finance |
| NPBT % | NPBT / Sales × 100 |
| Quick Ratio | (Cash + Receivables) / Current Liabilities |
| Current Asset Ratio | Current Assets / Current Liabilities |
| Net Asset Ratio | Net Assets / Total Liabilities |
| Debtor Turnover Days | Receivables / Sales × 365 |
| Creditor Turnover Days | Payables / (COS + Expenses) × 365 |
| Inventory Turnover Days | 365 / (COS / Inventory)（无存货 → N/A） |

### 4.6 FIN_SUM（build_fin_sum）

模板行定义在 `FIN_SUM_SECTIONS`（11 个区段，行序与 SUMMARY FINANCIALS 完全一致）：

```
Operating Revenue(1.001/1.002/1.003/1.01) → Direct Costs(3.001/3.002)
→ Other Income(2.001~2.005) → Operating Expenses(4.001~4.01)
→ Net Finance Costs(5.001~5.003) → Income Tax & Appropriations(6.001/7.001/7.002)
→ Current Assets(10.001~13.001) → Non Current Assets(11.02/12.003/12.004/14.001/15.001/16.001)
→ Current Liabilities(20.001~21.001) → Non Current Liabilities(20.02/21.002) → Equity(30.002/30.001/31.001/32.002)
```

**关键逻辑**：`Current Year Earnings（30.002）= NPBT`（本年盈余不在 TB 中，由损益结转）。否则恒等式不成立。

计算行：GROSS PROFIT → NPBT → NPAT（− 税与分配）→ CHECK PROFIT TO DATA → TOTAL ASSETS → TOTAL LIABILITIES → NET ASSETS → TOTAL EQUITY → BALANCE SHEET CHECK。

### 4.7 Tax Calc（calc_tax_calc + render_tax_calc_md，完整 16 区段）

| 区段 | 逻辑 |
|---|---|
| ① NPBT | 从指标取 |
| ② Add Back（11 项） | 当前为 0 占位（**改进方向：从 TB 准备科目自动识别**，见 8.2） |
| ③ Deduct（9 项） | 0 占位 |
| ④ Taxable Income | NPBT + 加回 − 减除 |
| ⑤ Tax Losses | 转入 − 利用 |
| ⑥ Taxable/(Loss) | ④ + ⑤ |
| ⑦ Tax Rate | 可配（SG 17% / AU BRE 25% / AU 30%） |
| ⑧ Tax Payable | 税额 − 红利抵免 − R&D − 其他 − PAYGI Q1-Q4 |
| ⑨-⑫ 对账表 | 亏损/资本亏损/税务准备/红利抵免（含 check to above / check to TB） |
| ⑬-⑭ 信托 | 非信托实体 → N/A 模板 |

### 4.8 其余表

| 表 | 实现方式 |
|---|---|
| JNL | 空模板 + 借贷平衡检查 |
| Template Rec | 科目余额对账（对账合计 vs TB，勾稽为 0） |
| Accr Exp | 23 项应计清单按 P&L 科目自动判定 YES/NO |
| AL / LSL | 模板 + 公式就绪（×附加成本率 / 精算参数） |
| Lease HP / DIV7A / Interco | 空模板 + 逻辑就绪（RATE/IPMT/矩阵） |
| HOLISTIC QC | 44 项自动判定结果（复用 run_qc_checks） |
| Review Notes / Client Queries | 人工记录空模板 |
| Packaging | 税额自动 + Y/N 占位 |

### 4.9 Excel 输出（write_excel）

openpyxl 生成 15 个 sheet，样式规范：
- 表头：深蓝底（1E3A8A）+ 白字
- 数字：`#,##0.00` 格式
- 关键行：加粗
- 列宽按内容设置

---

## 5. 关键设计决策与口径

### 5.1 符号约定统一
Xero TB 为"借方/贷方分列"（均为正），数据包转换为"带符号净额"（贷方为负），引擎按 Excel 公式符号规则（收入/负债/权益取反）恢复显示。**三层符号逻辑必须一致**：数据包 → TBMAP → 指标/填充。

### 5.2 FX 科目归类口径
MAP 表定义 5.003 = FX Gains & Losses（Net Finance Cost），但 Cloudfinesse 的 Xero P&L 把 FX 计入费用。若按 5.003 归类，NPBT 与 Xero 净利不一致导致恒等式差 67。**决策**：按客户 P&L 口径归入 4.001（可配置）。经验：FX 归类是税务口径判断，应按客户实际报表口径。

### 5.3 本年盈余不在 TB
Current Year Earnings（30.002）= NPBT 计算值，不是 TB 科目。所有报表结构（FIN_SUM/LEADSHEET/BS）的权益区必须用 NPBT 结转。

### 5.4 数据时效
P&L（区间报表）与 TB（截止日）可能因拉取时间不同而存在微小差异（本案例 Sales 差 530）。**应使用同一时间点的 TB 作为唯一数据源**，报表由其推导。

---

## 6. 验证结果（真实 Xero 数据）

| 校验项 | 结果 |
|---|---|
| BS 恒等式 | 0.0 ✔ |
| TB 借贷平衡 | 0.0 ✔ |
| MAP_ERR | 无未映射科目（60/60）✔ |
| Sales | 39,300.74（与 TB 一致）|
| Gross Profit | 34,863.74 |
| NPBT | 16,355.98 |
| 税额（17%） | 2,780.52 |
| FIN_SUM 双校验 | CHECK PROFIT 0 / BS CHECK 0 ✔ |

---

## 7. 部署与使用（换客户流程）

```bash
# 1. 准备数据包（从 Xero MCP 拉取）
#    list_accounts → accounts
#    get_trial_balance(本期) → tb_cy
#    get_trial_balance(上期) → tb_py
#    组装为 data/<client>_pack.json

# 2. 运行核心引擎（LEADSHEET + 指标）
python hbp_engine.py --data data/<client>_pack.json --out output/<client>

# 3. 运行全表生成器（19 张表 + Excel）
python hbp_reproduce_full.py   # 默认读取 data/cloudfinesse_pack.json
```

**接入 MCP 的自动化方式**：
1. 调用 `list_accounts` + `get_trial_balance` 拉数据
2. 组装数据包 JSON
3. 调用引擎生成全部表
4. 输出报告（Markdown/Excel）

---

## 8. 限制与改进方向

### 8.1 已识别限制
| 限制 | 原因 | 影响 |
|---|---|---|
| AL/LSL 为空 | 需工资数据（Payroll 403 受限） | 有数据即算 |
| 加回项为 0 | 硬编码占位 | 有准备科目余额的客户会漏加回 |
| 上年盈余为 0 | 上期 P&L 结转未接入 | 上期对比列待补充 |
| 人工表空 | 设计如此 | Review Notes/Client Queries |

### 8.2 改进方向（优先级排序）
1. **Tax Calc 加回项自动识别**：从 TB 准备科目自动映射加回（826 Superannuation Payable → Super 加回；年假/LSL 科目 → 准备加回）
2. **AL/LSL 工资数据输入模板**：提供 Payroll 报告粘贴模板 → 自动计算
3. **上年盈余结转**：上期 P&L 净利自动带入 30.002 上年列
4. **Lease/DIV7A 计算函数**：输入合同参数即出摊销表（RATE/IPMT 等价实现）
5. **MCP 一键生成**：封装 `generate_workpapers(org_id)` MCP 工具，端到端自动

---

## 9. 测试与验收

| 用例 | 输入 | 期望 |
|---|---|---|
| 正常客户 | 完整数据包 | 19 表输出，恒等式 0 |
| TB 不平衡 | 借贷不等 | 恒等式/TB 检查阻断 |
| 无租赁/股东贷款 | 无对应科目 | 空模板（N/A） |
| 多实体 | 提供 entities | Interco 矩阵生效 |
| 数据不足 | 缺 tb_py | 上期列 N/A，不误报 |
