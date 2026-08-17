# HBP Workpapers Reproduction Engine

Automated year-end workpaper generation from Xero trial balance data, reproducing the HBP (Holistic Business Partners) financial statements template logic.

## Features

- **LEADSHEET** (Trial Balance): 11 high-level KPIs + account-level table grouped by MAP categories (Balance Sheet & Profit & Loss), matching the template's 10-column structure
- **FIN_SUM** (Summary Financials): Balance Sheet by MAP category + P&L by expense category, with TOTAL ASSETS / TOTAL LIABILITIES / NET ASSETS / TOTAL EQUITY / BALANCE SHEET CHECK / CHECK PROFIT TO DATA rows
- **Tax Calc**: Net profit → add-backs → deductions → taxable income (16 sections)
- **Template Rec** & other reconciliations: 8-element structure
- **44-item QC**: automated quality checks (39/44 automatable) with evidence + suggested actions
- **Outputs**: full Markdown report (19 tables) + Excel workbook (15 sheets)

## Verification

Engine output has been validated line-by-line against Xero official reports (Profit & Loss / Balance Sheet) with **zero difference** when using the same data source.

## Usage

```bash
# Install dependencies
pip install openpyxl

# Run the engine
python hbp_reproduce_full.py

# Outputs
#   output/hbp_full_report.md   - full 19-table report
#   output/hbp_output.xlsx      - Excel workbook (15 sheets)
```

## Data Pack Format

The engine reads a JSON data pack (`data/*.json`) with this structure:

```json
{
  "org": "Company Name",
  "currency": "SGD",
  "period_end": "2026-08-31",
  "prior_period_end": "2025-12-31",
  "accounts": [{"code": "200", "name": "Sales", "type": "REVENUE"}],
  "tb_cy": {"200": -39249.00},
  "tb_py": {"200": -25436.00}
}
```

- `tb_cy` / `tb_py`: trial balance YTD values, **credits are negative**
- `accounts`: chart of accounts with type (REVENUE/EXPENSE/BANK/CURRENT/FIXED/CURRLIAB/TERMLIAB/EQUITY/...)

Data packs are excluded from this repository (`data/`, `output/` in `.gitignore`) as they contain client data.

## Files

| File | Purpose |
|---|---|
| `hbp_engine.py` | Core engine: data loading, MAP grouping, financials, metrics, tax calc, QC |
| `hbp_reproduce_full.py` | Full pipeline: generates Markdown report + Excel workbook |
| `hbp_tables_md.py` | Markdown table rendering helpers |
| `rules/map_rules.json` | 67 local MAP rules (account → financial statement category) |

## License

Internal use. Not for redistribution without permission.
