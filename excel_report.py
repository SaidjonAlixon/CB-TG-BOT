from __future__ import annotations

from collections import defaultdict
from datetime import date, timedelta
from io import BytesIO
from typing import Any

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import bot_db
from shifts import fmt_time, status_text


NAVY = "172B4D"
BLUE = "2F6BFF"
GREEN = "D9F2E3"
RED = "FADBD8"
ORANGE = "FCE4D6"
LIGHT = "F4F7FB"
WHITE = "FFFFFF"
THIN = Side(style="thin", color="D9E2F2")


def _style_sheet(ws, title: str, max_col: int) -> None:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    title_cell = ws.cell(1, 1, title)
    title_cell.font = Font(size=18, bold=True, color=WHITE)
    title_cell.fill = PatternFill("solid", fgColor=NAVY)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30


def _header(ws, row: int, values: list[str]) -> None:
    for col, value in enumerate(values, 1):
        cell = ws.cell(row, col, value)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN)
    ws.auto_filter.ref = f"A{row}:{get_column_letter(len(values))}{row}"
    ws.row_dimensions[row].height = 28


def _finish(ws, widths: dict[int, int]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = cell.alignment.copy(vertical="center")
            if cell.row > 2:
                cell.border = Border(bottom=THIN)


def _dates(start: date, end: date) -> list[date]:
    return [start + timedelta(days=i) for i in range((end - start).days + 1)]


def _attendance_map(records: list[dict[str, Any]]) -> dict[tuple[int, date], dict[str, Any]]:
    return {(int(row["employee_id"]), row["attendance_date"]): row for row in records}


def _row_for(employee: dict[str, Any], day: date, record: dict[str, Any] | None) -> dict[str, Any]:
    record = record or {}
    arrival_status = record.get("arrival_status") or "ABSENT"
    departure_status = record.get("departure_status") or ""
    return {
        "Sana": day,
        "Hafta": f"{((day - date(day.year, day.month, 1)).days // 7) + 1}-hafta",
        "Filial kodi": employee["branch_code"],
        "Filial nomi": employee["branch_name"],
        "F.I.Sh": employee["full_name"],
        "Lavozim": employee["position"],
        "Smena": f"{employee['shift']}-smena",
        "Keldi": fmt_time(record.get("arrival_at")),
        "Kelish holati": status_text(arrival_status),
        "Ketdi": fmt_time(record.get("departure_at")),
        "Ketish holati": status_text(departure_status) if departure_status else "—",
        "Ishlagan vaqt": (
            f"{(record.get('worked_minutes') or 0) // 60:02d}:{(record.get('worked_minutes') or 0) % 60:02d}"
            if record.get("worked_minutes") is not None
            else "—"
        ),
        "Kechikish (daq.)": record.get("late_minutes") or 0,
        "Erta ketish (daq.)": record.get("early_leave_minutes") or 0,
    }


def _write_rows(ws, rows: list[dict[str, Any]], headers: list[str], start_row: int = 3) -> None:
    for row_idx, row in enumerate(rows, start_row):
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row_idx, col_idx, row.get(header, "—"))
            if header == "Sana" and isinstance(cell.value, date):
                cell.number_format = "dd.mm.yyyy"
            if header in {"Kechikish (daq.)", "Erta ketish (daq.)"}:
                cell.number_format = "0"
            if row.get("Kelish holati", "").startswith("🟢"):
                cell.fill = PatternFill("solid", fgColor=GREEN)
            elif row.get("Kelish holati", "").startswith("🔴"):
                cell.fill = PatternFill("solid", fgColor=RED)
            if row.get("Ketish holati", "").startswith("🟠"):
                cell.fill = PatternFill("solid", fgColor=ORANGE)


def build_report(start: date, end: date, label: str) -> BytesIO:
    employees = bot_db.get_all_employees()
    records = bot_db.get_attendance_range(start, end)
    record_map = _attendance_map(records)
    days = _dates(start, end)
    daily_rows = [
        _row_for(employee, day, record_map.get((int(employee["id"]), day)))
        for day in days
        for employee in employees
    ]

    wb = Workbook()
    dashboard = wb.active
    dashboard.title = "Dashboard"
    _style_sheet(dashboard, f"FILIAL ATTENDANCE — {label}", 6)
    arrived = sum(1 for row in daily_rows if row["Keldi"] != "—")
    late = sum(1 for row in daily_rows if row["Kelish holati"].startswith("🔴"))
    absent = sum(1 for row in daily_rows if row["Kelish holati"].startswith("❌"))
    total = len(daily_rows)
    summary = [
        ("JAMI XODIM-KUN", total),
        ("KELGAN", arrived),
        ("KECHIKKAN", late),
        ("KELMAGAN", absent),
        ("DAVOMAT", f"{(arrived / total * 100) if total else 0:.1f}%"),
    ]
    for row_idx, (name, value) in enumerate(summary, 3):
        dashboard.cell(row_idx, 1, name).font = Font(bold=True, color=NAVY)
        dashboard.cell(row_idx, 2, value).font = Font(size=15, bold=True, color=BLUE)
        dashboard.cell(row_idx, 1).fill = PatternFill("solid", fgColor=LIGHT)
        dashboard.cell(row_idx, 2).fill = PatternFill("solid", fgColor=LIGHT)
    chart = BarChart()
    chart.title = "Davomat holati"
    chart.y_axis.title = "Soni"
    data = Reference(dashboard, min_col=2, min_row=4, max_row=6)
    cats = Reference(dashboard, min_col=1, min_row=4, max_row=6)
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)
    chart.height = 7
    chart.width = 13
    dashboard.add_chart(chart, "D3")
    _finish(dashboard, {1: 24, 2: 20, 3: 3, 4: 18, 5: 18, 6: 18})

    monthly = wb.create_sheet("Oylik hisobot")
    monthly_headers = [
        "№", "Filial kodi", "Filial nomi", "F.I.Sh", "Lavozim", "Smena",
        "Ish kunlari", "Kelgan", "Kelmagan", "Kechikkan", "Jami kechikish",
        "Jami ishlagan (daq.)", "Erta ketish",
    ]
    _style_sheet(monthly, f"Oylik hisobot: {label}", len(monthly_headers))
    _header(monthly, 2, monthly_headers)
    stats = bot_db.get_monthly_stats(start, end)
    for idx, row in enumerate(stats, 1):
        days_count = len(days)
        values = [
            idx, row["branch_code"], row["branch_name"], row["full_name"], row["position"],
            f"{row['shift']}-smena", days_count, row["arrived"],
            days_count - row["arrived"], row["late_days"], row["late_minutes"],
            row["worked_minutes"], row["early_minutes"],
        ]
        for col, value in enumerate(values, 1):
            monthly.cell(idx + 2, col, value)
    _finish(monthly, {1: 6, 2: 16, 3: 22, 4: 24, 5: 18, 6: 12, 7: 12, 8: 10, 9: 12, 10: 12, 11: 16, 12: 20, 13: 14})

    daily_headers = [
        "Sana", "Hafta", "Filial kodi", "Filial nomi", "F.I.Sh", "Lavozim", "Smena",
        "Keldi", "Kelish holati", "Ketdi", "Ketish holati", "Ishlagan vaqt",
        "Kechikish (daq.)", "Erta ketish (daq.)",
    ]
    for week_index, week_start in enumerate(range(0, len(days), 7), 1):
        week_days = days[week_start : week_start + 7]
        ws = wb.create_sheet(f"{week_index}-hafta")
        _style_sheet(ws, f"{week_index}-HAFTA", len(daily_headers))
        _header(ws, 2, daily_headers)
        rows = [row for row in daily_rows if row["Sana"] in week_days]
        _write_rows(ws, rows, daily_headers)
        _finish(ws, {1: 13, 2: 11, 3: 15, 4: 22, 5: 24, 6: 17, 7: 11, 8: 10, 9: 20, 10: 10, 11: 20, 12: 18, 13: 16, 14: 18})

    late_ws = wb.create_sheet("Kechikishlar")
    late_headers = ["Sana", "Filial", "F.I.Sh", "Lavozim", "Smena", "Belgilangan", "Keldi", "Kechikish"]
    _style_sheet(late_ws, "Kechikishlar", len(late_headers))
    _header(late_ws, 2, late_headers)
    late_rows = []
    for row in daily_rows:
        if row["Kelish holati"].startswith("🔴"):
            late_rows.append({
                "Sana": row["Sana"], "Filial": row["Filial nomi"], "F.I.Sh": row["F.I.Sh"],
                "Lavozim": row["Lavozim"], "Smena": row["Smena"],
                "Belgilangan": "08:15" if row["Smena"].startswith("1") else "17:15",
                "Keldi": row["Keldi"], "Kechikish": f"{row['Kechikish (daq.)']} daq.",
            })
    _write_rows(late_ws, late_rows, late_headers)
    _finish(late_ws, {1: 13, 2: 22, 3: 24, 4: 18, 5: 11, 6: 15, 7: 12, 8: 16})

    absent_ws = wb.create_sheet("Kelmaganlar")
    absent_headers = ["Sana", "Filial", "F.I.Sh", "Lavozim", "Smena", "Holat"]
    _style_sheet(absent_ws, "Kelmaganlar", len(absent_headers))
    _header(absent_ws, 2, absent_headers)
    absent_rows = [
        {
            "Sana": row["Sana"], "Filial": row["Filial nomi"], "F.I.Sh": row["F.I.Sh"],
            "Lavozim": row["Lavozim"], "Smena": row["Smena"], "Holat": "❌ Kelmagan",
        }
        for row in daily_rows
        if row["Kelish holati"].startswith("❌")
    ]
    _write_rows(absent_ws, absent_rows, absent_headers)
    _finish(absent_ws, {1: 13, 2: 22, 3: 24, 4: 18, 5: 11, 6: 18})

    branches_ws = wb.create_sheet("Filiallar")
    branch_headers = ["№", "Filial kodi", "Filial nomi", "Xodimlar", "Kelganlar", "Kechikkanlar", "Kelmaganlar", "Davomat"]
    _style_sheet(branches_ws, "Filiallar statistikasi", len(branch_headers))
    _header(branches_ws, 2, branch_headers)
    branch_groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in daily_rows:
        branch_groups[row["Filial kodi"]].append(row)
    for idx, (code, rows) in enumerate(branch_groups.items(), 1):
        present_count = sum(1 for row in rows if row["Keldi"] != "—")
        late_count = sum(1 for row in rows if row["Kelish holati"].startswith("🔴"))
        absent_count = sum(1 for row in rows if row["Kelish holati"].startswith("❌"))
        values = [idx, code, rows[0]["Filial nomi"], len({r["F.I.Sh"] for r in rows}), present_count, late_count, absent_count, f"{present_count / len(rows) * 100:.1f}%" if rows else "0.0%"]
        for col, value in enumerate(values, 1):
            branches_ws.cell(idx + 2, col, value)
    _finish(branches_ws, {1: 6, 2: 16, 3: 24, 4: 12, 5: 12, 6: 15, 7: 14, 8: 12})

    employees_ws = wb.create_sheet("Xodimlar")
    employee_headers = ["№", "Filial kodi", "Filial nomi", "F.I.Sh", "Telefon", "Lavozim", "Smena", "Holat"]
    _style_sheet(employees_ws, "Xodimlar", len(employee_headers))
    _header(employees_ws, 2, employee_headers)
    for idx, employee in enumerate(employees, 1):
        values = [idx, employee["branch_code"], employee["branch_name"], employee["full_name"], employee["phone"], employee["position"], f"{employee['shift']}-smena", "Faol"]
        for col, value in enumerate(values, 1):
            employees_ws.cell(idx + 2, col, value)
    _finish(employees_ws, {1: 6, 2: 16, 3: 24, 4: 24, 5: 18, 6: 18, 7: 11, 8: 12})

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value == "❌ Kelmagan":
                    cell.fill = PatternFill("solid", fgColor=RED)
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
