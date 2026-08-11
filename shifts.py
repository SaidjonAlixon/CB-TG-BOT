from __future__ import annotations

from datetime import datetime, time, timedelta
from zoneinfo import ZoneInfo


LOCAL_TZ = ZoneInfo("Asia/Tashkent")

SHIFT_RULES = {
    "1": {
        "name": "1-SMENA",
        "arrival_deadline": time(8, 15),
        "departure_deadline": time(17, 0),
    },
    "2": {
        "name": "2-SMENA",
        "arrival_deadline": time(17, 15),
        "departure_deadline": time(23, 45),
    },
}


def now_local() -> datetime:
    return datetime.now(LOCAL_TZ)


def minute_value(value: time) -> int:
    return value.hour * 60 + value.minute


def normalized_minutes(moment: datetime) -> int:
    local = moment.astimezone(LOCAL_TZ)
    return local.hour * 60 + local.minute


def arrival_result(moment: datetime, shift: str) -> tuple[str, int]:
    deadline = minute_value(SHIFT_RULES[shift]["arrival_deadline"])
    actual = normalized_minutes(moment)
    late = max(0, actual - deadline)
    return ("LATE" if late else "ON_TIME", late)


def departure_result(moment: datetime, shift: str) -> tuple[str, int]:
    deadline = minute_value(SHIFT_RULES[shift]["departure_deadline"])
    local = moment.astimezone(LOCAL_TZ)
    actual = local.hour * 60 + local.minute
    if shift == "2" and local.hour < 6:
        actual += 24 * 60
    early = max(0, deadline - actual)
    return ("EARLY_LEAVE" if early else "ON_TIME", early)


def worked_minutes(arrival: datetime, departure: datetime) -> int:
    start = arrival.astimezone(LOCAL_TZ)
    end = departure.astimezone(LOCAL_TZ)
    if end < start:
        end += timedelta(days=1)
    return max(0, int((end - start).total_seconds() // 60))


def fmt_time(value: datetime | None) -> str:
    if not value:
        return "—"
    return value.astimezone(LOCAL_TZ).strftime("%H:%M")


def fmt_duration(minutes: int | None) -> str:
    if minutes is None:
        return "—"
    total = max(0, int(minutes))
    hours, mins = divmod(total, 60)
    if hours and mins:
        return f"{hours} soat {mins:02d} daqiqa"
    if hours:
        return f"{hours} soat"
    return f"{mins} daqiqa"


def fmt_hours_clock(minutes: int | None) -> str:
    """Excel / qisqa ko‘rinish: 05:46"""
    if minutes is None:
        return "—"
    total = max(0, int(minutes))
    return f"{total // 60:02d}:{total % 60:02d}"


def status_text(status: str | None) -> str:
    return {
        "ON_TIME": "🟢 Vaqtida",
        "LATE": "🔴 Kechikdi",
        "EARLY_LEAVE": "🟠 Erta ketdi",
        "COMPLETED": "✅ Ish kuni yakunlandi",
        "MISSING_CHECKOUT": "⚠️ Ketish qayd qilinmagan",
        "ABSENT": "❌ Kelmagan",
    }.get(status or "", "—")
