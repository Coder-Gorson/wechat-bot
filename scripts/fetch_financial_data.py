"""
联合能源集团 (00467.HK) 财务数据获取工具
从东方财富数据中心抓取资产负债表、利润表、现金流量表关键指标，并输出到 Excel。

依赖安装：
    pip install -r requirements.txt

运行方式：
    python fetch_financial_data.py
"""

import sys
import time
import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

STOCK_CODE = "00467"
STOCK_NAME = "联合能源集团"

BASE_URL = "https://datacenter.eastmoney.com/securities/api/data/v1/get"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": "https://emweb.securities.eastmoney.com/",
}

# ---------------------------------------------------------------------------
# 东方财富港股财务报表字段映射
# ---------------------------------------------------------------------------

BALANCE_COLUMNS = (
    "REPORT_DATE,TOTAL_ASSETS,TOTAL_LIABILITIES,TOTAL_EQUITY,"
    "MONETARYFUNDS,ACCOUNTS_RECE,INVENTORY,"
    "TOTAL_CURRENT_ASSETS,TOTAL_CURRENT_LIABILITIES"
)

INCOME_COLUMNS = (
    "REPORT_DATE,TOTAL_OPERATE_INCOME,OPERATE_PROFIT,TOTAL_PROFIT,"
    "NETPROFIT,GROSS_PROFIT,TOTAL_OPERATE_COST"
)

CASHFLOW_COLUMNS = (
    "REPORT_DATE,NETCASH_OPERATE,NETCASH_INVEST,NETCASH_FINANCE,"
    "FREE_CASHFLOW,END_CASH"
)

# 中文字段标签
BALANCE_LABELS = {
    "TOTAL_ASSETS": "总资产",
    "TOTAL_LIABILITIES": "总负债",
    "TOTAL_EQUITY": "净资产（股东权益）",
    "MONETARYFUNDS": "货币资金",
    "ACCOUNTS_RECE": "应收账款",
    "INVENTORY": "存货",
    "TOTAL_CURRENT_ASSETS": "流动资产合计",
    "TOTAL_CURRENT_LIABILITIES": "流动负债合计",
}

INCOME_LABELS = {
    "TOTAL_OPERATE_INCOME": "营业收入",
    "OPERATE_PROFIT": "经营利润",
    "TOTAL_PROFIT": "税前利润",
    "NETPROFIT": "净利润",
    "GROSS_PROFIT": "毛利润",
    "TOTAL_OPERATE_COST": "营业成本合计",
}

CASHFLOW_LABELS = {
    "NETCASH_OPERATE": "经营活动现金流净额",
    "NETCASH_INVEST": "投资活动现金流净额",
    "NETCASH_FINANCE": "筹资活动现金流净额",
    "FREE_CASHFLOW": "自由现金流",
    "END_CASH": "期末现金及现金等价物",
}


def fetch_report(report_name: str, columns: str, page_size: int = 5) -> list[dict]:
    """从东方财富数据中心获取港股财务报表数据。"""
    params = {
        "reportName": report_name,
        "columns": columns,
        "filter": f'(SECURITY_CODE="{STOCK_CODE}")',
        "pageNumber": 1,
        "pageSize": page_size,
        "sortTypes": -1,
        "sortColumns": "REPORT_DATE",
        "source": "HSF10",
        "client": "PC",
    }
    try:
        resp = requests.get(BASE_URL, params=params, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        rows = data.get("result", {}).get("data") or []
        return rows
    except requests.RequestException as exc:
        print(f"  [错误] 请求 {report_name} 失败: {exc}", file=sys.stderr)
        return []


def format_value(value) -> str:
    """将数值格式化为 '亿' 单位的字符串，保留两位小数。"""
    if value is None or value == "" or (isinstance(value, float) and value != value):
        return "—"
    try:
        num = float(value)
        return f"{num / 1e8:.2f} 亿"
    except (TypeError, ValueError):
        return str(value)


def extract_year(date_str: str) -> str:
    """从日期字符串（如 '2023-12-31'）提取年份标签。"""
    if not date_str:
        return "未知"
    try:
        dt = datetime.strptime(date_str[:10], "%Y-%m-%d")
        return f"{dt.year}年"
    except ValueError:
        return date_str[:4]


def build_sheet_data(rows: list[dict], labels: dict) -> tuple[list[str], dict]:
    """
    构建表头（年份列表）和数据字典 {field: {year: formatted_value}}。
    只保留最近 5 个年报（12-31）。
    """
    annual_rows = [r for r in rows if r.get("REPORT_DATE", "").endswith("12-31")]
    # 取最近 N 年（最多 5 年），按时间升序排列用于表格列
    annual_rows = sorted(annual_rows, key=lambda r: r["REPORT_DATE"])[-5:]

    years = [extract_year(r["REPORT_DATE"]) for r in annual_rows]
    data: dict[str, dict[str, str]] = {field: {} for field in labels}

    for row in annual_rows:
        year = extract_year(row["REPORT_DATE"])
        for field in labels:
            data[field][year] = format_value(row.get(field))

    return years, data


def build_derived_balance(rows: list[dict], years: list[str]) -> dict:
    """计算资产负债表衍生指标（资产负债率、流动比率）。"""
    derived: dict[str, dict[str, str]] = {
        "资产负债率": {},
        "流动比率": {},
    }
    annual_rows = {
        extract_year(r["REPORT_DATE"]): r
        for r in rows
        if r.get("REPORT_DATE", "").endswith("12-31")
    }
    for year in years:
        row = annual_rows.get(year, {})
        try:
            total_assets = float(row.get("TOTAL_ASSETS") or 0)
            total_liabilities = float(row.get("TOTAL_LIABILITIES") or 0)
            current_assets = float(row.get("TOTAL_CURRENT_ASSETS") or 0)
            current_liabilities = float(row.get("TOTAL_CURRENT_LIABILITIES") or 0)

            if total_assets:
                derived["资产负债率"][year] = f"{total_liabilities / total_assets * 100:.2f}%"
            else:
                derived["资产负债率"][year] = "—"

            if current_liabilities:
                derived["流动比率"][year] = f"{current_assets / current_liabilities:.2f}"
            else:
                derived["流动比率"][year] = "—"
        except (TypeError, ValueError):
            derived["资产负债率"][year] = "—"
            derived["流动比率"][year] = "—"
    return derived


def build_derived_income(rows: list[dict], years: list[str]) -> dict:
    """计算利润表衍生指标（毛利率、净利率）。"""
    derived: dict[str, dict[str, str]] = {
        "毛利率": {},
        "净利率": {},
    }
    annual_rows = {
        extract_year(r["REPORT_DATE"]): r
        for r in rows
        if r.get("REPORT_DATE", "").endswith("12-31")
    }
    for year in years:
        row = annual_rows.get(year, {})
        try:
            revenue = float(row.get("TOTAL_OPERATE_INCOME") or 0)
            gross_profit = float(row.get("GROSS_PROFIT") or 0)
            net_profit = float(row.get("NETPROFIT") or 0)

            if revenue:
                derived["毛利率"][year] = f"{gross_profit / revenue * 100:.2f}%"
                derived["净利率"][year] = f"{net_profit / revenue * 100:.2f}%"
            else:
                derived["毛利率"][year] = "—"
                derived["净利率"][year] = "—"
        except (TypeError, ValueError):
            derived["毛利率"][year] = "—"
            derived["净利率"][year] = "—"
    return derived


# ---------------------------------------------------------------------------
# Excel 样式辅助函数
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
SUBHEADER_FILL = PatternFill("solid", fgColor="2E75B6")
ROW_FILL_ODD = PatternFill("solid", fgColor="DEEAF1")
ROW_FILL_EVEN = PatternFill("solid", fgColor="FFFFFF")

THIN_BORDER = Border(
    left=Side(style="thin", color="BDD7EE"),
    right=Side(style="thin", color="BDD7EE"),
    top=Side(style="thin", color="BDD7EE"),
    bottom=Side(style="thin", color="BDD7EE"),
)


def style_cell(cell, bold=False, color="000000", fill=None, align="center"):
    cell.font = Font(bold=bold, color=color, name="微软雅黑", size=11)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border = THIN_BORDER
    if fill:
        cell.fill = fill


def write_section(
    ws,
    start_row: int,
    title: str,
    years: list[str],
    main_labels: dict,
    main_data: dict,
    derived: dict | None = None,
) -> int:
    """
    在工作表中写入一个报表区块，返回下一个可用行号。

    布局：
      标题行（合并）
      表头行：指标 | year1 | year2 | ...
      数据行（交替底色）
      衍生指标行（可选）
    """
    col_count = 1 + len(years)

    # 标题行
    ws.merge_cells(
        start_row=start_row,
        start_column=1,
        end_row=start_row,
        end_column=col_count,
    )
    title_cell = ws.cell(row=start_row, column=1, value=title)
    style_cell(title_cell, bold=True, color="FFFFFF", fill=HEADER_FILL)
    start_row += 1

    # 表头行
    ws.cell(row=start_row, column=1, value="指标")
    style_cell(ws.cell(row=start_row, column=1), bold=True, color="FFFFFF", fill=SUBHEADER_FILL)
    for col_idx, year in enumerate(years, start=2):
        cell = ws.cell(row=start_row, column=col_idx, value=year)
        style_cell(cell, bold=True, color="FFFFFF", fill=SUBHEADER_FILL)
    start_row += 1

    # 主要指标行
    for row_idx, (field, label) in enumerate(main_labels.items()):
        fill = ROW_FILL_ODD if row_idx % 2 == 0 else ROW_FILL_EVEN
        label_cell = ws.cell(row=start_row, column=1, value=label)
        style_cell(label_cell, align="left", fill=fill)
        for col_idx, year in enumerate(years, start=2):
            val = main_data.get(field, {}).get(year, "—")
            cell = ws.cell(row=start_row, column=col_idx, value=val)
            style_cell(cell, fill=fill)
        start_row += 1

    # 衍生指标行
    if derived:
        offset = len(main_labels)
        for row_idx, (label, year_vals) in enumerate(derived.items()):
            fill = ROW_FILL_ODD if (row_idx + offset) % 2 == 0 else ROW_FILL_EVEN
            label_cell = ws.cell(row=start_row, column=1, value=label)
            style_cell(label_cell, align="left", fill=fill)
            for col_idx, year in enumerate(years, start=2):
                val = year_vals.get(year, "—")
                cell = ws.cell(row=start_row, column=col_idx, value=val)
                style_cell(cell, fill=fill)
            start_row += 1

    return start_row + 1  # 空一行


def save_excel(
    balance_rows: list[dict],
    income_rows: list[dict],
    cashflow_rows: list[dict],
    output_path: str,
) -> None:
    """将三张报表数据写入 Excel 文件。"""

    # ---- 资产负债表 ----
    b_years, b_data = build_sheet_data(balance_rows, BALANCE_LABELS)
    b_derived = build_derived_balance(balance_rows, b_years)

    # ---- 利润表 ----
    i_years, i_data = build_sheet_data(income_rows, INCOME_LABELS)
    i_derived = build_derived_income(income_rows, i_years)

    # ---- 现金流量表 ----
    c_years, c_data = build_sheet_data(cashflow_rows, CASHFLOW_LABELS)

    # 取三张表的年份并集，统一列对齐
    all_years_set: set[str] = set(b_years) | set(i_years) | set(c_years)
    all_years = sorted(all_years_set)

    # 如果某张表没有数据，用空列补充
    def fill_missing_years(data: dict, years: list[str]) -> dict:
        for field in data:
            for year in years:
                data[field].setdefault(year, "—")
        return data

    b_data = fill_missing_years(b_data, all_years)
    i_data = fill_missing_years(i_data, all_years)
    c_data = fill_missing_years(c_data, all_years)
    for d in (b_derived, i_derived):
        for label in d:
            for year in all_years:
                d[label].setdefault(year, "—")

    wb = Workbook()
    ws = wb.active
    ws.title = f"{STOCK_NAME}财务指标简表"

    # 工作表标题
    col_count = 1 + len(all_years)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_count)
    title_cell = ws.cell(
        row=1,
        column=1,
        value=f"{STOCK_NAME}（{STOCK_CODE}.HK）财务指标简表（单位：亿港元，除另有注明）",
    )
    style_cell(title_cell, bold=True, color="FFFFFF", fill=PatternFill("solid", fgColor="1F3864"))
    ws.row_dimensions[1].height = 30

    current_row = 3
    current_row = write_section(
        ws,
        current_row,
        "一、资产负债表关键指标",
        all_years,
        BALANCE_LABELS,
        b_data,
        derived=b_derived,
    )
    current_row = write_section(
        ws,
        current_row,
        "二、利润表关键指标",
        all_years,
        INCOME_LABELS,
        i_data,
        derived=i_derived,
    )
    current_row = write_section(
        ws,
        current_row,
        "三、现金流量表关键指标",
        all_years,
        CASHFLOW_LABELS,
        c_data,
    )

    # 列宽自动调整
    ws.column_dimensions[get_column_letter(1)].width = 22
    for col in range(2, col_count + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16

    # 冻结首列和前两行
    ws.freeze_panes = "B3"

    # 数据来源注释
    note_row = current_row
    ws.merge_cells(
        start_row=note_row, start_column=1, end_row=note_row, end_column=col_count
    )
    note_cell = ws.cell(
        row=note_row,
        column=1,
        value=f"数据来源：东方财富数据中心  |  获取时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}",
    )
    note_cell.font = Font(italic=True, color="808080", size=9)
    note_cell.alignment = Alignment(horizontal="left")

    wb.save(output_path)
    print(f"\n✅ Excel 文件已保存至：{output_path}")


def main():
    print(f"📊 开始获取 {STOCK_NAME}（{STOCK_CODE}.HK）财务数据...\n")

    print("  正在获取资产负债表数据...")
    balance_rows = fetch_report("RPT_HKF10_FN_BALANCE", BALANCE_COLUMNS, page_size=10)
    time.sleep(0.5)

    print("  正在获取利润表数据...")
    income_rows = fetch_report("RPT_HKF10_FN_INCOME", INCOME_COLUMNS, page_size=10)
    time.sleep(0.5)

    print("  正在获取现金流量表数据...")
    cashflow_rows = fetch_report("RPT_HKF10_FN_CASH", CASHFLOW_COLUMNS, page_size=10)

    if not any([balance_rows, income_rows, cashflow_rows]):
        print(
            "\n❌ 未能获取到任何财务数据。\n"
            "   可能原因：\n"
            "   1. 网络连接问题或接口限流，请稍后重试。\n"
            "   2. 东方财富接口字段或报表名称已更新。\n",
            file=sys.stderr,
        )
        sys.exit(1)

    output_path = f"{STOCK_NAME}_{STOCK_CODE}_财务指标简表.xlsx"
    save_excel(balance_rows, income_rows, cashflow_rows, output_path)

    # 在控制台打印简要摘要
    print("\n--- 数据摘要预览 ---")
    for label, rows in [
        ("资产负债表", balance_rows),
        ("利润表", income_rows),
        ("现金流量表", cashflow_rows),
    ]:
        annual = [r for r in rows if r.get("REPORT_DATE", "").endswith("12-31")]
        dates = sorted(r["REPORT_DATE"][:7] for r in annual)
        print(f"  {label}：共 {len(annual)} 条年报数据，期间：{dates[0] if dates else '—'} ~ {dates[-1] if dates else '—'}")


if __name__ == "__main__":
    main()
