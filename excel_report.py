from __future__ import annotations

from collections import defaultdict
from datetime import date
from io import BytesIO
from typing import Any

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import bot_db
from shifts import fmt_hours_clock, fmt_time, now_local, status_text


NAVY = "0F2744"
BLUE = "1F4E79"
ACCENT = "2E75B6"
GREEN_BG = "C6EFCE"
GREEN_FG = "006100"
RED_BG = "FFC7CE"
RED_FG = "9C0006"
ORANGE_BG = "FCE4D6"
ORANGE_FG = "C65911"
YELLOW_BG = "FFF2CC"
LIGHT = "F2F5F9"
WHITE = "FFFFFF"
GRAY = "667085"
LINE = "D0D7E2"
THIN = Side(style="thin", color=LINE)
MED = Side(style="medium", color=BLUE)


def _fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def _font(size: int = 11, bold: bool = False, color: str = NAVY, name: str = "Calibri") -> Font:
    return Font(name=name, size=size, bold=bold, color=color)


def _align(h: str = "left", v: str = "center", wrap: bool = False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _border() -> Border:
    return Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _style_sheet(ws, title: str, max_col: int, subtitle: str = "") -> None:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5" if subtitle else "A4"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    title_cell = ws.cell(1, 1, title)
    title_cell.font = _font(18, True, WHITE)
    title_cell.fill = _fill(NAVY)
    title_cell.alignment = _align("center")
    ws.row_dimensions[1].height = 34
    if subtitle:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
        sub = ws.cell(2, 1, subtitle)
        sub.font = _font(11, False, WHITE)
        sub.fill = _fill(BLUE)
        sub.alignment = _align("center")
        ws.row_dimensions[2].height = 22


def _header(ws, row: int, values: list[str]) -> None:
    for col, value in enumerate(values, 1):
        cell = ws.cell(row, col, value)
        cell.font = _font(11, True, WHITE)
        cell.fill = _fill(ACCENT)
        cell.alignment = _align("center", wrap=True)
        cell.border = _border()
    last = get_column_letter(len(values))
    ws.auto_filter.ref = f"A{row}:{last}{row}"
    ws.row_dimensions[row].height = 28


def _finish(ws, widths: dict[int, int]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    for row in ws.iter_rows():
        for cell in row:
            if cell.alignment.vertical != "center":
                cell.alignment = cell.alignment.copy(vertical="center")


def _kpi_card(ws, row: int, col: int, title: str, value: Any, bg: str, fg: str = NAVY) -> None:
    label = ws.cell(row, col, title)
    label.font = _font(9, True, GRAY)
    label.fill = _fill(bg)
    label.alignment = _align("center")
    label.border = _border()
    val = ws.cell(row + 1, col, value)
    val.font = _font(18, True, fg)
    val.fill = _fill(bg)
    val.alignment = _align("center")
    val.border = _border()
    ws.row_dimensions[row].height = 18
    ws.row_dimensions[row + 1].height = 32


def _section_title(ws, row: int, text: str, max_col: int, bg: str = BLUE) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
    cell = ws.cell(row, 1, text)
    cell.font = _font(12, True, WHITE)
    cell.fill = _fill(bg)
    cell.alignment = _align("left")
    ws.row_dimensions[row].height = 24


def _dates(start: date, end: date) -> list[date]:
    """Hisobot uchun faqat admin belgilagan ish kunlari."""
    return bot_db.list_work_days(start, end)


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
            fmt_hours_clock(record.get("worked_minutes"))
            if record.get("worked_minutes") is not None
            else "—"
        ),
        "Kechikish (soat)": fmt_hours_clock(record.get("late_minutes") or 0),
        "Erta ketish (soat)": fmt_hours_clock(record.get("early_leave_minutes") or 0),
    }


def _write_rows(ws, rows: list[dict[str, Any]], headers: list[str], start_row: int = 3) -> None:
    for row_idx, row in enumerate(rows, start_row):
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row_idx, col_idx, row.get(header, "—"))
            cell.border = _border()
            cell.alignment = _align("center" if header != "F.I.Sh" else "left")
            if header == "Sana" and isinstance(cell.value, date):
                cell.number_format = "dd.mm.yyyy"
            holat = str(row.get("Kelish holati", row.get("Holat", "")))
            if holat.startswith("🟢") or holat.startswith("✅"):
                cell.fill = _fill(GREEN_BG)
            elif holat.startswith("🔴") or holat.startswith("❌"):
                cell.fill = _fill(RED_BG)
            elif holat.startswith("🟠"):
                cell.fill = _fill(ORANGE_BG)


def _build_dashboard(
    wb: Workbook,
    label: str,
    focus_day: date,
    daily_rows: list[dict[str, Any]],
    employees: list[dict[str, Any]],
    day_rows: list[dict[str, Any]],
) -> None:
    ws = wb.active
    ws.title = "Dashboard"
    _style_sheet(
        ws,
        "FILIAL ATTENDANCE — DASHBOARD",
        8,
        subtitle=f"Davr: {label}   |   Kunlik sana: {focus_day.strftime('%d.%m.%Y')}   |   Yangilangan: {now_local().strftime('%d.%m.%Y %H:%M')}",
    )

    arrived = sum(1 for row in day_rows if row["Keldi"] != "—")
    late = sum(1 for row in day_rows if row["Kelish holati"].startswith("🔴"))
    absent = sum(1 for row in day_rows if row["Kelish holati"].startswith("❌"))
    on_time = sum(1 for row in day_rows if row["Kelish holati"].startswith("🟢"))
    total_emp = len(employees)
    rate = (arrived / total_emp * 100) if total_emp else 0.0

    period_arrived = sum(1 for row in daily_rows if row["Keldi"] != "—")
    period_late = sum(1 for row in daily_rows if row["Kelish holati"].startswith("🔴"))
    period_absent = sum(1 for row in daily_rows if row["Kelish holati"].startswith("❌"))
    period_total = len(daily_rows)
    period_rate = (period_arrived / period_total * 100) if period_total else 0.0

    _section_title(ws, 4, f"📌 BUGUNGI KO‘RSATKICHLAR — {focus_day.strftime('%d.%m.%Y')}", 8, NAVY)
    cards = [
        ("JAMI XODIM", total_emp, LIGHT, NAVY),
        ("KELGAN", arrived, GREEN_BG, GREEN_FG),
        ("VAQTIDA", on_time, GREEN_BG, GREEN_FG),
        ("KECHIKKAN", late, ORANGE_BG, ORANGE_FG),
        ("KELMAGAN", absent, RED_BG, RED_FG),
        ("DAVOMAT %", f"{rate:.1f}%", YELLOW_BG, NAVY),
    ]
    for idx, (title, value, bg, fg) in enumerate(cards):
        _kpi_card(ws, 5, idx + 1, title, value, bg, fg)

    _section_title(ws, 8, f"📊 DAVR BO‘YICHA YAKUN — {label}", 8, BLUE)
    period_cards = [
        ("XODIM-KUN", period_total, LIGHT, NAVY),
        ("KELGAN", period_arrived, GREEN_BG, GREEN_FG),
        ("KECHIKKAN", period_late, ORANGE_BG, ORANGE_FG),
        ("KELMAGAN", period_absent, RED_BG, RED_FG),
        ("DAVOMAT %", f"{period_rate:.1f}%", YELLOW_BG, NAVY),
    ]
    for idx, (title, value, bg, fg) in enumerate(period_cards):
        _kpi_card(ws, 9, idx + 1, title, value, bg, fg)

    # Chart source (hidden-ish area)
    ws.cell(9, 7, "Kelgan").font = _font(9, True, GRAY)
    ws.cell(9, 8, period_arrived)
    ws.cell(10, 7, "Kechikkan").font = _font(9, True, GRAY)
    ws.cell(10, 8, period_late)
    ws.cell(11, 7, "Kelmagan").font = _font(9, True, GRAY)
    ws.cell(11, 8, period_absent)

    chart = BarChart()
    chart.type = "col"
    chart.title = "Davomat holati (davr)"
    chart.y_axis.title = "Soni"
    chart.style = 10
    data = Reference(ws, min_col=8, min_row=9, max_row=11)
    cats = Reference(ws, min_col=7, min_row=9, max_row=11)
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)
    chart.shape = 4
    chart.height = 8
    chart.width = 12
    ws.add_chart(chart, "A12")

    # Branch table
    table_start = 22
    _section_title(ws, table_start, "🏢 FILIALLAR KESIMIDA (BUGUN)", 8, ACCENT)
    headers = ["№", "Filial", "Xodimlar", "Kelgan", "Kechikkan", "Kelmagan", "Davomat %", "Holat"]
    _header(ws, table_start + 1, headers)

    branch_groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in day_rows:
        branch_groups[row["Filial nomi"]].append(row)

    r = table_start + 2
    for idx, (name, rows) in enumerate(sorted(branch_groups.items()), 1):
        emp_count = len(rows)
        a = sum(1 for row in rows if row["Keldi"] != "—")
        l = sum(1 for row in rows if row["Kelish holati"].startswith("🔴"))
        ab = sum(1 for row in rows if row["Kelish holati"].startswith("❌"))
        pct = (a / emp_count * 100) if emp_count else 0.0
        status = "Yaxshi" if pct >= 90 else ("O‘rtacha" if pct >= 70 else "Past")
        values = [idx, name, emp_count, a, l, ab, f"{pct:.1f}%", status]
        for col, value in enumerate(values, 1):
            cell = ws.cell(r, col, value)
            cell.border = _border()
            cell.alignment = _align("center" if col != 2 else "left")
            if col == 8:
                if status == "Yaxshi":
                    cell.fill = _fill(GREEN_BG)
                    cell.font = _font(11, True, GREEN_FG)
                elif status == "O‘rtacha":
                    cell.fill = _fill(YELLOW_BG)
                    cell.font = _font(11, True, ORANGE_FG)
                else:
                    cell.fill = _fill(RED_BG)
                    cell.font = _font(11, True, RED_FG)
        r += 1

    guide_row = r + 1
    _section_title(ws, guide_row, "ℹ️ QANDAY O‘QILADI", 8, NAVY)
    tips = [
        "• Kunlik list — tanlangan sanadagi kelganlar va kelmaganlar.",
        "• Kelganlar list — faqat bugun (hisobot sanasi) kelgan xodimlar, vaqt bilan.",
        "• Dashboard foizi = Kelgan / Jami xodim. Dam olish kunida kelmagan = 0.",
        "• Excel hisobotlari faqat admin belgilagan ish kunlari asosida hisoblanadi.",
    ]
    for i, tip in enumerate(tips):
        cell = ws.cell(guide_row + 1 + i, 1, tip)
        ws.merge_cells(start_row=guide_row + 1 + i, start_column=1, end_row=guide_row + 1 + i, end_column=8)
        cell.font = _font(10, False, GRAY)
        cell.fill = _fill(LIGHT)

    _finish(ws, {1: 8, 2: 28, 3: 12, 4: 12, 5: 12, 6: 12, 7: 12, 8: 12})


def _build_kunlik(
    wb: Workbook,
    focus_day: date,
    day_rows: list[dict[str, Any]],
) -> None:
    ws = wb.create_sheet("Kunlik", 1)
    date_label = focus_day.strftime("%d.%m.%Y")
    _style_sheet(
        ws,
        f"KUNLIK HISOBOT — {date_label}",
        8,
        subtitle=f"Sana: {date_label}   |   Ish kuni: {'Ha' if bot_db.is_work_day(focus_day) else 'Yo‘q'}",
    )

    arrived_rows = [row for row in day_rows if row["Keldi"] != "—"]
    absent_rows = [row for row in day_rows if row["Kelish holati"].startswith("❌")]
    late_rows = [row for row in day_rows if row["Kelish holati"].startswith("🔴")]

    _kpi_card(ws, 4, 1, "SANA", date_label, LIGHT, NAVY)
    _kpi_card(ws, 4, 2, "JAMI", len(day_rows), LIGHT, NAVY)
    _kpi_card(ws, 4, 3, "KELGAN", len(arrived_rows), GREEN_BG, GREEN_FG)
    _kpi_card(ws, 4, 4, "KECHIKKAN", len(late_rows), ORANGE_BG, ORANGE_FG)
    _kpi_card(ws, 4, 5, "KELMAGAN", len(absent_rows), RED_BG, RED_FG)

    # Kelganlar section
    start = 7
    _section_title(ws, start, f"✅ BUGUN KELGANLAR — {date_label}", 8, GREEN_FG)
    arrived_headers = ["№", "Filial", "F.I.Sh", "Lavozim", "Smena", "Keldi", "Holat", "Kechikish"]
    _header(ws, start + 1, arrived_headers)
    if arrived_rows:
        for idx, row in enumerate(arrived_rows, 1):
            values = [
                idx,
                row["Filial nomi"],
                row["F.I.Sh"],
                row["Lavozim"],
                row["Smena"],
                row["Keldi"],
                row["Kelish holati"],
                row["Kechikish (soat)"],
            ]
            r = start + 1 + idx
            for col, value in enumerate(values, 1):
                cell = ws.cell(r, col, value)
                cell.border = _border()
                cell.alignment = _align("center" if col != 3 else "left")
                cell.fill = _fill(GREEN_BG if not str(row["Kelish holati"]).startswith("🔴") else ORANGE_BG)
        next_row = start + 2 + len(arrived_rows) + 1
    else:
        ws.cell(start + 2, 1, "Bugun kelganlar yo‘q.").font = _font(11, False, GRAY)
        next_row = start + 4

    _section_title(ws, next_row, f"❌ BUGUN KELMAGANLAR — {date_label}", 8, RED_FG)
    absent_headers = ["№", "Filial", "F.I.Sh", "Lavozim", "Smena", "Holat"]
    _header(ws, next_row + 1, absent_headers)
    if absent_rows:
        for idx, row in enumerate(absent_rows, 1):
            values = [idx, row["Filial nomi"], row["F.I.Sh"], row["Lavozim"], row["Smena"], "❌ Kelmagan"]
            r = next_row + 1 + idx
            for col, value in enumerate(values, 1):
                cell = ws.cell(r, col, value)
                cell.border = _border()
                cell.alignment = _align("center" if col != 3 else "left")
                cell.fill = _fill(RED_BG)
    else:
        msg = "Bugun barcha xodimlar kelgan." if bot_db.is_work_day(focus_day) else "Bugun ish kuni emas — kelmaganlar hisoblanmaydi."
        ws.cell(next_row + 2, 1, msg).font = _font(11, False, GRAY)

    _finish(ws, {1: 6, 2: 24, 3: 26, 4: 16, 5: 12, 6: 12, 7: 16, 8: 12})


def _build_kelganlar(wb: Workbook, focus_day: date, day_rows: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet("Kelganlar", 2)
    date_label = focus_day.strftime("%d.%m.%Y")
    arrived_rows = [row for row in day_rows if row["Keldi"] != "—"]
    _style_sheet(
        ws,
        f"BUGUN KELGANLAR — {date_label}",
        9,
        subtitle=f"Sana: {date_label}   |   Jami kelgan: {len(arrived_rows)} ta",
    )

    _kpi_card(ws, 4, 1, "SANA", date_label, LIGHT, NAVY)
    _kpi_card(ws, 4, 2, "KELGANLAR", len(arrived_rows), GREEN_BG, GREEN_FG)
    on_time = sum(1 for row in arrived_rows if row["Kelish holati"].startswith("🟢"))
    late = sum(1 for row in arrived_rows if row["Kelish holati"].startswith("🔴"))
    _kpi_card(ws, 4, 3, "VAQTIDA", on_time, GREEN_BG, GREEN_FG)
    _kpi_card(ws, 4, 4, "KECHIKKAN", late, ORANGE_BG, ORANGE_FG)

    headers = ["№", "Sana", "Filial", "F.I.Sh", "Lavozim", "Smena", "Keldi", "Ketdi", "Holat"]
    _header(ws, 7, headers)
    for idx, row in enumerate(arrived_rows, 1):
        values = [
            idx,
            focus_day,
            row["Filial nomi"],
            row["F.I.Sh"],
            row["Lavozim"],
            row["Smena"],
            row["Keldi"],
            row["Ketdi"],
            row["Kelish holati"],
        ]
        for col, value in enumerate(values, 1):
            cell = ws.cell(7 + idx, col, value)
            cell.border = _border()
            cell.alignment = _align("center" if col != 4 else "left")
            if col == 2:
                cell.number_format = "dd.mm.yyyy"
            if str(row["Kelish holati"]).startswith("🔴"):
                cell.fill = _fill(ORANGE_BG)
            else:
                cell.fill = _fill(GREEN_BG)
    if not arrived_rows:
        ws.cell(8, 1, f"{date_label} sanasida kelganlar yo‘q.").font = _font(11, False, GRAY)

    _finish(ws, {1: 6, 2: 12, 3: 22, 4: 26, 5: 16, 6: 12, 7: 10, 8: 10, 9: 16})


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

    focus_day = end
    # Kunlik fokusat focus_day; agar ish kuni bo‘lmasa ham xodimlar holatini ko‘rsatamiz
    day_rows = [
        _row_for(employee, focus_day, record_map.get((int(employee["id"]), focus_day)))
        for employee in employees
    ]

    wb = Workbook()
    _build_dashboard(wb, label, focus_day, daily_rows, employees, day_rows)
    _build_kunlik(wb, focus_day, day_rows)
    _build_kelganlar(wb, focus_day, day_rows)

    monthly = wb.create_sheet("Oylik hisobot")
    monthly_headers = [
        "№", "Filial kodi", "Filial nomi", "F.I.Sh", "Lavozim", "Smena",
        "Ish kunlari", "Kelgan", "Kelmagan", "Kechikkan", "Jami kechikish (soat)",
        "Jami ishlagan (soat)", "Erta ketish (soat)",
    ]
    _style_sheet(monthly, f"Oylik hisobot: {label}", len(monthly_headers), subtitle=f"Davr: {label}")
    _header(monthly, 3, monthly_headers)
    stats = bot_db.get_monthly_stats(start, end)
    for idx, row in enumerate(stats, 1):
        days_count = len(days)
        values = [
            idx, row["branch_code"], row["branch_name"], row["full_name"], row["position"],
            f"{row['shift']}-smena", days_count, row["arrived"],
            max(0, days_count - row["arrived"]), row["late_days"],
            fmt_hours_clock(row["late_minutes"]),
            fmt_hours_clock(row["worked_minutes"]),
            fmt_hours_clock(row["early_minutes"]),
        ]
        for col, value in enumerate(values, 1):
            cell = monthly.cell(idx + 3, col, value)
            cell.border = _border()
    _finish(monthly, {1: 6, 2: 16, 3: 22, 4: 24, 5: 18, 6: 12, 7: 12, 8: 10, 9: 12, 10: 12, 11: 20, 12: 20, 13: 18})

    daily_headers = [
        "Sana", "Hafta", "Filial kodi", "Filial nomi", "F.I.Sh", "Lavozim", "Smena",
        "Keldi", "Kelish holati", "Ketdi", "Ketish holati", "Ishlagan vaqt",
        "Kechikish (soat)", "Erta ketish (soat)",
    ]
    for week_index, week_start in enumerate(range(0, len(days), 7), 1):
        week_days = days[week_start : week_start + 7]
        ws = wb.create_sheet(f"{week_index}-hafta")
        _style_sheet(ws, f"{week_index}-HAFTA", len(daily_headers))
        _header(ws, 3, daily_headers)
        rows = [row for row in daily_rows if row["Sana"] in week_days]
        _write_rows(ws, rows, daily_headers, start_row=4)
        _finish(ws, {1: 13, 2: 11, 3: 15, 4: 22, 5: 24, 6: 17, 7: 11, 8: 10, 9: 20, 10: 10, 11: 20, 12: 18, 13: 16, 14: 18})

    late_ws = wb.create_sheet("Kechikishlar")
    late_headers = ["Sana", "Filial", "F.I.Sh", "Lavozim", "Smena", "Belgilangan", "Keldi", "Kechikish"]
    _style_sheet(late_ws, "Kechikishlar", len(late_headers))
    _header(late_ws, 3, late_headers)
    late_rows = []
    for row in daily_rows:
        if row["Kelish holati"].startswith("🔴"):
            late_rows.append({
                "Sana": row["Sana"], "Filial": row["Filial nomi"], "F.I.Sh": row["F.I.Sh"],
                "Lavozim": row["Lavozim"], "Smena": row["Smena"],
                "Belgilangan": "08:15" if row["Smena"].startswith("1") else "17:15",
                "Keldi": row["Keldi"], "Kechikish": row["Kechikish (soat)"],
            })
    _write_rows(late_ws, late_rows, late_headers, start_row=4)
    _finish(late_ws, {1: 13, 2: 22, 3: 24, 4: 18, 5: 11, 6: 15, 7: 12, 8: 16})

    absent_ws = wb.create_sheet("Kelmaganlar")
    absent_headers = ["Sana", "Filial", "F.I.Sh", "Lavozim", "Smena", "Holat"]
    _style_sheet(absent_ws, "Kelmaganlar", len(absent_headers))
    _header(absent_ws, 3, absent_headers)
    absent_rows = [
        {
            "Sana": row["Sana"], "Filial": row["Filial nomi"], "F.I.Sh": row["F.I.Sh"],
            "Lavozim": row["Lavozim"], "Smena": row["Smena"], "Holat": "❌ Kelmagan",
        }
        for row in daily_rows
        if row["Kelish holati"].startswith("❌")
    ]
    _write_rows(absent_ws, absent_rows, absent_headers, start_row=4)
    _finish(absent_ws, {1: 13, 2: 22, 3: 24, 4: 18, 5: 11, 6: 18})

    branches_ws = wb.create_sheet("Filiallar")
    branch_headers = ["№", "Filial kodi", "Filial nomi", "Xodimlar", "Kelganlar", "Kechikkanlar", "Kelmaganlar", "Davomat"]
    _style_sheet(branches_ws, "Filiallar statistikasi", len(branch_headers))
    _header(branches_ws, 3, branch_headers)
    branch_groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in daily_rows:
        branch_groups[row["Filial kodi"]].append(row)
    for idx, (code, rows) in enumerate(branch_groups.items(), 1):
        present_count = sum(1 for row in rows if row["Keldi"] != "—")
        late_count = sum(1 for row in rows if row["Kelish holati"].startswith("🔴"))
        absent_count = sum(1 for row in rows if row["Kelish holati"].startswith("❌"))
        values = [
            idx, code, rows[0]["Filial nomi"], len({r["F.I.Sh"] for r in rows}),
            present_count, late_count, absent_count,
            f"{present_count / len(rows) * 100:.1f}%" if rows else "0.0%",
        ]
        for col, value in enumerate(values, 1):
            cell = branches_ws.cell(idx + 3, col, value)
            cell.border = _border()
    _finish(branches_ws, {1: 6, 2: 16, 3: 24, 4: 12, 5: 12, 6: 15, 7: 14, 8: 12})

    employees_ws = wb.create_sheet("Xodimlar")
    employee_headers = ["№", "Filial kodi", "Filial nomi", "F.I.Sh", "Telefon", "Lavozim", "Smena", "Holat"]
    _style_sheet(employees_ws, "Xodimlar", len(employee_headers))
    _header(employees_ws, 3, employee_headers)
    for idx, employee in enumerate(employees, 1):
        values = [
            idx, employee["branch_code"], employee["branch_name"], employee["full_name"],
            employee["phone"], employee["position"], f"{employee['shift']}-smena", "Faol",
        ]
        for col, value in enumerate(values, 1):
            cell = employees_ws.cell(idx + 3, col, value)
            cell.border = _border()
    _finish(employees_ws, {1: 6, 2: 16, 3: 24, 4: 24, 5: 18, 6: 18, 7: 11, 8: 12})

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value == "❌ Kelmagan":
                    cell.fill = _fill(RED_BG)
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
