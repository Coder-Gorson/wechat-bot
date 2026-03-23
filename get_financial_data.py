"""
get_financial_data.py
=====================
Fetches publicly available financial report data for 联合能源集团 (00467.HK)
from Eastmoney (东方财富) and writes key indicators to an Excel workbook.

Dependencies (install once):
    pip install requests openpyxl

Usage:
    python get_financial_data.py
"""

import sys
import time
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

STOCK_CODE = "00467"  # 联合能源集团 HK listing code

# Eastmoney DataCenter API – Hong Kong financial report endpoint
_BASE = "https://datacenter.eastmoney.com/securities/api/data/v1/get"

_COMMON_PARAMS = {
    "pageNumber": 1,
    "pageSize": 5,   # fetch up to 5 periods; we will keep the latest 3
    "sortTypes": -1,
    "sortColumns": "REPORT_DATE",
    "source": "F10",
    "client": "PC",
}

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": "https://emweb.securities.eastmoney.com/",
}

# ---------------------------------------------------------------------------
# Report definitions
# ---------------------------------------------------------------------------

# Each entry: (label, api_column_name)
BALANCE_SHEET_FIELDS = [
    ("总资产 (Total Assets)",              "TOTAL_ASSETS"),
    ("总负债 (Total Liabilities)",          "TOTAL_LIABILITIES"),
    ("净资产 (Total Equity)",               "TOTAL_EQUITY"),
    ("流动资产 (Current Assets)",            "TOTAL_CURRENT_ASSETS"),
    ("流动负债 (Current Liabilities)",       "TOTAL_CURRENT_LIABILITIES"),
    ("货币资金 (Cash & Equivalents)",        "MONETARYFUNDS"),
]

INCOME_FIELDS = [
    ("营业收入 (Revenue)",                   "TOTAL_OPERATE_INCOME"),
    ("营业成本 (Cost of Revenue)",           "TOTAL_OPERATE_COST"),
    ("营业利润 (Operating Profit)",          "OPERATE_PROFIT"),
    ("净利润 (Net Profit)",                  "NETPROFIT"),
    ("归母净利润 (Net Profit Attr. to Parent)", "PARENT_NETPROFIT"),
]

CASHFLOW_FIELDS = [
    ("经营活动现金流净额 (Operating CF)",     "NETCASH_OPERATE"),
    ("投资活动现金流净额 (Investing CF)",     "NETCASH_INVEST"),
    ("筹资活动现金流净额 (Financing CF)",     "NETCASH_FINANCE"),
    ("期末现金及等价物 (Ending Cash)",        "END_CASH"),
]

_REPORT_CONFIGS = {
    "balance": {
        "reportName": "RPT_HKFINANCE_BALANCE_PC",
        "columns": ["REPORT_DATE"] + [f for _, f in BALANCE_SHEET_FIELDS],
        "fields": BALANCE_SHEET_FIELDS,
        "sheet_title": "资产负债表 Balance Sheet",
    },
    "income": {
        "reportName": "RPT_HKFINANCE_INCOME_PC",
        "columns": ["REPORT_DATE"] + [f for _, f in INCOME_FIELDS],
        "fields": INCOME_FIELDS,
        "sheet_title": "利润表 Income Statement",
    },
    "cashflow": {
        "reportName": "RPT_HKFINANCE_CASHFLOW_PC",
        "columns": ["REPORT_DATE"] + [f for _, f in CASHFLOW_FIELDS],
        "fields": CASHFLOW_FIELDS,
        "sheet_title": "现金流量表 Cash Flow Statement",
    },
}

# ---------------------------------------------------------------------------
# Derived-ratio definitions (computed from raw data)
# ---------------------------------------------------------------------------

def _pct(numerator, denominator):
    """Return percentage string, or '—' when denominator is zero/None."""
    try:
        if denominator and float(denominator) != 0:
            return f"{float(numerator) / float(denominator) * 100:.2f}%"
    except (TypeError, ValueError):
        pass
    return "—"


def compute_derived(report_type, rows_by_period):
    """Return list of (label, {period: value}) derived rows for a report type."""
    derived = []
    if report_type == "balance":
        derived.append((
            "资产负债率 (Debt Ratio)",
            {p: _pct(d.get("TOTAL_LIABILITIES"), d.get("TOTAL_ASSETS"))
             for p, d in rows_by_period.items()},
        ))
        derived.append((
            "流动比率 (Current Ratio)",
            {p: (
                f"{float(d.get('TOTAL_CURRENT_ASSETS', 0)) / float(d.get('TOTAL_CURRENT_LIABILITIES', 1)):.2f}"
                if d.get("TOTAL_CURRENT_LIABILITIES") and float(d.get("TOTAL_CURRENT_LIABILITIES", 0)) != 0
                else "—"
            )
             for p, d in rows_by_period.items()},
        ))
    elif report_type == "income":
        derived.append((
            "毛利率 (Gross Margin)",
            {p: _pct(
                float(d.get("TOTAL_OPERATE_INCOME", 0) or 0) - float(d.get("TOTAL_OPERATE_COST", 0) or 0),
                d.get("TOTAL_OPERATE_INCOME"),
            )
             for p, d in rows_by_period.items()},
        ))
        derived.append((
            "净利率 (Net Margin)",
            {p: _pct(d.get("NETPROFIT"), d.get("TOTAL_OPERATE_INCOME"))
             for p, d in rows_by_period.items()},
        ))
    return derived


# ---------------------------------------------------------------------------
# Data fetching
# ---------------------------------------------------------------------------

def fetch_report(report_type: str, max_periods: int = 3) -> dict:
    """
    Fetch a financial report from Eastmoney and return a dict:
        {
          "periods": [str, ...],          # sorted ascending, e.g. ['2022', '2023', '2024']
          "rows_by_period": {period: {field: value, ...}, ...},
          "fields": [(label, col), ...],
          "sheet_title": str,
        }
    Returns an empty dict on failure.
    """
    cfg = _REPORT_CONFIGS[report_type]
    params = {
        **_COMMON_PARAMS,
        "reportName": cfg["reportName"],
        "columns": ",".join(cfg["columns"]),
        "filter": f'(SECURITY_CODE="{STOCK_CODE}")',
    }

    try:
        resp = requests.get(_BASE, params=params, headers=_HEADERS, timeout=15)
        resp.raise_for_status()
        data = resp.json()
    except Exception as exc:
        print(f"  ⚠  Failed to fetch {report_type}: {exc}")
        return {}

    records = (data.get("result") or {}).get("data") or []
    if not records:
        print(f"  ⚠  No data returned for {report_type}.")
        return {}

    # Build period-keyed dict (use year part of REPORT_DATE)
    rows_by_period: dict[str, dict] = {}
    for rec in records:
        raw_date = rec.get("REPORT_DATE", "")
        year = str(raw_date)[:4] if raw_date else "Unknown"
        rows_by_period[year] = rec

    # Keep only the latest `max_periods` periods, sorted ascending
    periods = sorted(rows_by_period.keys())[-max_periods:]
    rows_by_period = {p: rows_by_period[p] for p in periods}

    return {
        "periods": periods,
        "rows_by_period": rows_by_period,
        "fields": cfg["fields"],
        "sheet_title": cfg["sheet_title"],
    }


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

_HDR_FILL   = PatternFill("solid", fgColor="1F4E79")
_HDR_FONT   = Font(bold=True, color="FFFFFF", size=11)
_LABEL_FILL = PatternFill("solid", fgColor="D6E4F0")
_LABEL_FONT = Font(bold=True, size=10)
_TITLE_FONT = Font(bold=True, size=13)
_DERIVED_FILL = PatternFill("solid", fgColor="EBF3FB")
_THIN = Side(style="thin", color="AAAAAA")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

# Eastmoney HK report values are in thousands HKD; this converts to 亿 (100 million)
_THOUSANDS_HKD_TO_YI = 100_000_000


def _fmt_value(val) -> str:
    """Format a raw numeric value to a human-readable string (億 units)."""
    if val is None:
        return "—"
    try:
        num = float(val)
        # Values from Eastmoney HK reports are in thousands HKD
        yi = num / _THOUSANDS_HKD_TO_YI
        return f"{yi:,.4f} 亿"
    except (TypeError, ValueError):
        return str(val)


def write_sheet(ws, report_data: dict, report_type: str):
    """Write one financial report sheet."""
    periods = report_data["periods"]
    rows_by_period = report_data["rows_by_period"]
    fields = report_data["fields"]
    title = report_data["sheet_title"]

    # --- Title row ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + len(periods))
    title_cell = ws.cell(row=1, column=1, value=f"联合能源集团 (00467.HK) — {title}")
    title_cell.font = _TITLE_FONT
    title_cell.alignment = _CENTER

    # --- Header row ---
    ws.cell(row=2, column=1, value="指标 / Metric").font = _HDR_FONT
    ws.cell(row=2, column=1).fill = _HDR_FILL
    ws.cell(row=2, column=1).alignment = _CENTER
    for col_idx, period in enumerate(periods, start=2):
        cell = ws.cell(row=2, column=col_idx, value=period)
        cell.font = _HDR_FONT
        cell.fill = _HDR_FILL
        cell.alignment = _CENTER

    # --- Data rows (raw) ---
    for row_idx, (label, col) in enumerate(fields, start=3):
        label_cell = ws.cell(row=row_idx, column=1, value=label)
        label_cell.fill = _LABEL_FILL
        label_cell.font = _LABEL_FONT
        label_cell.alignment = _LEFT
        for col_idx, period in enumerate(periods, start=2):
            raw = rows_by_period[period].get(col)
            cell = ws.cell(row=row_idx, column=col_idx, value=_fmt_value(raw))
            cell.alignment = _CENTER

    # --- Derived ratios ---
    derived = compute_derived(report_type, rows_by_period)
    offset = 3 + len(fields)
    for d_idx, (label, period_values) in enumerate(derived):
        row_idx = offset + d_idx
        label_cell = ws.cell(row=row_idx, column=1, value=label)
        label_cell.fill = _DERIVED_FILL
        label_cell.font = _LABEL_FONT
        label_cell.alignment = _LEFT
        for col_idx, period in enumerate(periods, start=2):
            cell = ws.cell(row=row_idx, column=col_idx, value=period_values.get(period, "—"))
            cell.alignment = _CENTER

    # --- Apply borders to data area ---
    max_row = offset + len(derived) - 1
    max_col = 1 + len(periods)
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).border = _BORDER

    # --- Column widths ---
    ws.column_dimensions[get_column_letter(1)].width = 38
    for c in range(2, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 18
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 22

    # --- Note row ---
    note_row = max_row + 2
    note_cell = ws.cell(
        row=note_row, column=1,
        value="注：数值单位为亿港元（HKD 亿），数据来源：东方财富（eastmoney.com）",
    )
    note_cell.font = Font(italic=True, size=9, color="666666")
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=max_col)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    output_file = "联合能源集团_财务数据.xlsx"

    print("=" * 60)
    print("  联合能源集团 (00467.HK) 财务数据获取工具")
    print("  数据来源: 东方财富 (eastmoney.com)")
    print("=" * 60)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    report_order = [
        ("balance",  "资产负债表"),
        ("income",   "利润表"),
        ("cashflow", "现金流量表"),
    ]

    all_ok = True
    for report_type, label in report_order:
        print(f"\n正在获取 {label} …")
        time.sleep(0.5)  # be polite to the server
        data = fetch_report(report_type)
        if not data:
            all_ok = False
            ws = wb.create_sheet(title=label)
            ws.cell(row=1, column=1, value=f"⚠ 无法获取 {label} 数据，请稍后重试。")
            continue

        periods_str = "、".join(data["periods"])
        print(f"  ✓ 获取到 {len(data['periods'])} 期数据（{periods_str}）")
        ws = wb.create_sheet(title=label)
        write_sheet(ws, data, report_type)

    # --- Summary / cover sheet ---
    cover = wb.create_sheet(title="说明", index=0)
    cover["A1"] = "联合能源集团 (00467.HK) 财务指标简表"
    cover["A1"].font = Font(bold=True, size=16)
    cover["A3"] = f"数据来源: 东方财富 (eastmoney.com) — {_BASE}"
    cover["A4"] = "说明: 数值单位为亿港元（HKD），原始数据精度为千港元（×1000 HKD）"
    cover["A5"] = "包含报表: 资产负债表 / 利润表 / 现金流量表"
    cover["A6"] = "衍生指标: 资产负债率、流动比率、毛利率、净利率"
    cover.column_dimensions["A"].width = 70

    wb.save(output_file)
    print(f"\n{'✅' if all_ok else '⚠'} 已保存至: {output_file}")
    if not all_ok:
        print("  部分报表获取失败，请检查网络连接后重试。")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
