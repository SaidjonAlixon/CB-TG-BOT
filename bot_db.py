from __future__ import annotations

import os
from datetime import date, datetime, timedelta
from typing import Any

import psycopg
from psycopg.rows import dict_row


SCHEMA_SQL = """
CREATE SCHEMA IF NOT EXISTS filial;

CREATE TABLE IF NOT EXISTS filial.branches (
    id SERIAL PRIMARY KEY,
    branch_code TEXT NOT NULL UNIQUE,
    branch_name TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'active',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS filial.employees (
    id SERIAL PRIMARY KEY,
    telegram_id BIGINT NOT NULL UNIQUE,
    phone TEXT NOT NULL,
    full_name TEXT NOT NULL,
    branch_id INTEGER NOT NULL REFERENCES filial.branches(id),
    position TEXT NOT NULL,
    shift TEXT NOT NULL CHECK (shift IN ('1', '2')),
    status TEXT NOT NULL DEFAULT 'active',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS filial.attendance (
    id SERIAL PRIMARY KEY,
    employee_id INTEGER NOT NULL REFERENCES filial.employees(id) ON DELETE CASCADE,
    attendance_date DATE NOT NULL,
    arrival_at TIMESTAMPTZ,
    departure_at TIMESTAMPTZ,
    worked_minutes INTEGER,
    arrival_status TEXT,
    departure_status TEXT,
    late_minutes INTEGER NOT NULL DEFAULT 0,
    early_leave_minutes INTEGER NOT NULL DEFAULT 0,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE (employee_id, attendance_date)
);

CREATE TABLE IF NOT EXISTS filial.admins (
    telegram_id BIGINT PRIMARY KEY,
    full_name TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'admin',
    status TEXT NOT NULL DEFAULT 'active',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS filial.change_requests (
    id SERIAL PRIMARY KEY,
    employee_id INTEGER NOT NULL REFERENCES filial.employees(id) ON DELETE CASCADE,
    request_type TEXT NOT NULL,
    requested_value TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'pending',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    reviewed_at TIMESTAMPTZ
);

CREATE INDEX IF NOT EXISTS attendance_date_idx ON filial.attendance (attendance_date);
CREATE INDEX IF NOT EXISTS attendance_employee_idx ON filial.attendance (employee_id);

CREATE TABLE IF NOT EXISTS filial.schema_meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS filial.work_weekdays (
    weekday SMALLINT PRIMARY KEY CHECK (weekday BETWEEN 0 AND 6),
    is_work BOOLEAN NOT NULL DEFAULT TRUE
);

CREATE TABLE IF NOT EXISTS filial.work_day_overrides (
    work_date DATE PRIMARY KEY,
    is_work BOOLEAN NOT NULL,
    note TEXT NOT NULL DEFAULT '',
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
"""


def _database_url() -> str:
    url = os.environ.get("DATABASE_URL", "").strip()
    if not url:
        raise RuntimeError("DATABASE_URL sozlanmagan.")
    # Railway sometimes injects postgres:// — psycopg3 expects postgresql://
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://") :]
    return url


def _connect() -> psycopg.Connection:
    conn = psycopg.connect(_database_url(), row_factory=dict_row)
    conn.execute("SET search_path TO filial, public")
    return conn


# Bitta Telegram polling instance uchun (Railway + lokal to‘qnashuvini oldini oladi)
BOT_INSTANCE_LOCK_KEY = 874_512_309


def acquire_bot_instance_lock() -> psycopg.Connection:
    """Faqat bitta process polling qilsin. Ulanish ochiq turganda lock saqlanadi."""
    conn = psycopg.connect(_database_url(), autocommit=True)
    conn.execute("SELECT pg_advisory_lock(%s)", (BOT_INSTANCE_LOCK_KEY,))
    return conn


def init_db() -> None:
    """Jadvallarni yaratadi yoki mavjudlarini qoldiradi.

    CREATE IF NOT EXISTS ishlatiladi — restart / redeploy / qayta o‘rnatishda
    filiallar, xodimlar va davomat HECH QACHON o‘chirilmaydi.
    """
    with _connect() as conn:
        conn.execute(SCHEMA_SQL)
        conn.execute(
            """
            INSERT INTO schema_meta (key, value)
            VALUES ('schema_version', '1')
            ON CONFLICT (key) DO UPDATE
            SET value = EXCLUDED.value, updated_at = NOW()
            """
        )
        admin_ids = os.environ.get("ADMIN_TELEGRAM_IDS", "")
        for raw_id in admin_ids.split(","):
            raw_id = raw_id.strip()
            if raw_id.isdigit():
                conn.execute(
                    """
                    INSERT INTO admins (telegram_id, full_name, role)
                    VALUES (%s, 'Administrator', 'admin')
                    ON CONFLICT (telegram_id) DO UPDATE
                    SET status = 'active', role = 'admin'
                    """,
                    (int(raw_id),),
                )
        # Default: Dushanba–Shanba ish kuni, Yakshanba dam olish
        for weekday in range(7):
            conn.execute(
                """
                INSERT INTO work_weekdays (weekday, is_work)
                VALUES (%s, %s)
                ON CONFLICT (weekday) DO NOTHING
                """,
                (weekday, weekday != 6),
            )


def persistence_stats() -> dict[str, int]:
    """Saqlangan asosiy yozuvlar soni (restart diagnostikasi uchun)."""
    with _connect() as conn:
        row = conn.execute(
            """
            SELECT
                (SELECT COUNT(*)::int FROM branches) AS branches,
                (SELECT COUNT(*)::int FROM employees WHERE status = 'active') AS employees,
                (SELECT COUNT(*)::int FROM attendance) AS attendance_rows,
                (SELECT COUNT(*)::int FROM admins WHERE status = 'active') AS admins
            """
        ).fetchone()
    return {
        "branches": int(row["branches"]),
        "employees": int(row["employees"]),
        "attendance_rows": int(row["attendance_rows"]),
        "admins": int(row["admins"]),
    }


def is_admin(telegram_id: int) -> bool:
    """To‘liq yoki yordamchi admin."""
    return get_admin_role(telegram_id) is not None


def get_admin_role(telegram_id: int) -> str | None:
    with _connect() as conn:
        row = conn.execute(
            """
            SELECT role FROM admins
            WHERE telegram_id = %s AND status = 'active'
            """,
            (telegram_id,),
        ).fetchone()
    if not row:
        return None
    role = str(row["role"] or "admin").strip().lower()
    return role if role in {"admin", "helper"} else "admin"


def is_full_admin(telegram_id: int) -> bool:
    return get_admin_role(telegram_id) == "admin"


def is_helper_admin(telegram_id: int) -> bool:
    return get_admin_role(telegram_id) == "helper"


MAX_MANUAL_FULL_ADMINS = 2


def env_admin_ids() -> set[int]:
    ids: set[int] = set()
    for raw_id in os.environ.get("ADMIN_TELEGRAM_IDS", "").split(","):
        raw_id = raw_id.strip()
        if raw_id.isdigit():
            ids.add(int(raw_id))
    return ids


def list_manual_full_admins() -> list[dict[str, Any]]:
    """ENV dan tashqari, ID orqali qo‘shilgan to‘liq adminlar."""
    env_ids = env_admin_ids()
    with _connect() as conn:
        rows = conn.execute(
            """
            SELECT telegram_id, full_name, status, created_at
            FROM admins
            WHERE role = 'admin' AND status = 'active'
            ORDER BY created_at
            """
        ).fetchall()
    return [row for row in rows if int(row["telegram_id"]) not in env_ids]


def add_full_admin(telegram_id: int, full_name: str = "Administrator") -> tuple[bool, str]:
    """ID orqali to‘liq admin qo‘shadi (maksimal 2 ta)."""
    if telegram_id in env_admin_ids():
        return False, "Bu ID allaqachon asosiy (ENV) admin."
    manuals = list_manual_full_admins()
    if telegram_id not in {int(row["telegram_id"]) for row in manuals} and len(manuals) >= MAX_MANUAL_FULL_ADMINS:
        return False, f"Ko‘pi bilan {MAX_MANUAL_FULL_ADMINS} ta admin qo‘shish mumkin."
    with _connect() as conn:
        existing = conn.execute(
            "SELECT role, status FROM admins WHERE telegram_id = %s",
            (telegram_id,),
        ).fetchone()
        if existing and existing["role"] == "helper" and existing["status"] == "active":
            # Yordamchini to‘liq adminga ko‘tarish
            pass
        conn.execute(
            """
            INSERT INTO admins (telegram_id, full_name, role, status)
            VALUES (%s, %s, 'admin', 'active')
            ON CONFLICT (telegram_id) DO UPDATE SET
                full_name = EXCLUDED.full_name,
                role = 'admin',
                status = 'active'
            """,
            (telegram_id, full_name.strip() or "Administrator"),
        )
    return True, "Admin qo‘shildi. U /start bosib to‘liq admin panelini ochadi."


def remove_full_admin(telegram_id: int) -> tuple[bool, str]:
    if telegram_id in env_admin_ids():
        return False, "Asosiy (ENV) adminni o‘chirib bo‘lmaydi."
    with _connect() as conn:
        existing = conn.execute(
            "SELECT role FROM admins WHERE telegram_id = %s AND status = 'active'",
            (telegram_id,),
        ).fetchone()
        if not existing:
            return False, "Bunday faol admin topilmadi."
        if existing["role"] != "admin":
            return False, "Bu ID to‘liq admin emas."
        conn.execute(
            "UPDATE admins SET status = 'inactive' WHERE telegram_id = %s AND role = 'admin'",
            (telegram_id,),
        )
    return True, "Admin o‘chirildi."


def list_helper_admins() -> list[dict[str, Any]]:
    with _connect() as conn:
        return conn.execute(
            """
            SELECT telegram_id, full_name, status, created_at
            FROM admins
            WHERE role = 'helper' AND status = 'active'
            ORDER BY created_at
            """
        ).fetchall()


def add_helper_admin(telegram_id: int, full_name: str = "Yordamchi admin") -> tuple[bool, str]:
    """Yordamchi admin qo‘shadi. To‘liq adminni pastga tushirmaydi."""
    with _connect() as conn:
        existing = conn.execute(
            "SELECT role, status FROM admins WHERE telegram_id = %s",
            (telegram_id,),
        ).fetchone()
        if existing and existing["role"] == "admin" and existing["status"] == "active":
            return False, "Bu ID allaqachon to‘liq admin."
        conn.execute(
            """
            INSERT INTO admins (telegram_id, full_name, role, status)
            VALUES (%s, %s, 'helper', 'active')
            ON CONFLICT (telegram_id) DO UPDATE SET
                full_name = EXCLUDED.full_name,
                role = 'helper',
                status = 'active'
            """,
            (telegram_id, full_name.strip() or "Yordamchi admin"),
        )
    return True, "Yordamchi admin qo‘shildi."


def remove_helper_admin(telegram_id: int) -> tuple[bool, str]:
    with _connect() as conn:
        existing = conn.execute(
            "SELECT role FROM admins WHERE telegram_id = %s AND status = 'active'",
            (telegram_id,),
        ).fetchone()
        if not existing:
            return False, "Bunday faol yordamchi admin topilmadi."
        if existing["role"] != "helper":
            return False, "To‘liq adminni shu yerdan o‘chirib bo‘lmaydi."
        conn.execute(
            "UPDATE admins SET status = 'inactive' WHERE telegram_id = %s AND role = 'helper'",
            (telegram_id,),
        )
    return True, "Yordamchi admin o‘chirildi."


def get_employee(telegram_id: int) -> dict[str, Any] | None:
    with _connect() as conn:
        return conn.execute(
            """
            SELECT e.*, b.branch_code, b.branch_name
            FROM employees e
            JOIN branches b ON b.id = e.branch_id
            WHERE e.telegram_id = %s AND e.status = 'active'
            """,
            (telegram_id,),
        ).fetchone()


def get_employee_by_id(employee_id: int) -> dict[str, Any] | None:
    with _connect() as conn:
        return conn.execute(
            """
            SELECT e.*, b.branch_code, b.branch_name
            FROM employees e
            JOIN branches b ON b.id = e.branch_id
            WHERE e.id = %s
            """,
            (employee_id,),
        ).fetchone()


def get_branches(search: str = "") -> list[dict[str, Any]]:
    with _connect() as conn:
        if search:
            return conn.execute(
                """
                SELECT * FROM branches
                WHERE status = 'active'
                  AND (branch_code ILIKE %s OR branch_name ILIKE %s)
                ORDER BY branch_name
                """,
                (f"%{search}%", f"%{search}%"),
            ).fetchall()
        return conn.execute(
            "SELECT * FROM branches WHERE status = 'active' ORDER BY branch_name"
        ).fetchall()


def get_branch(branch_id: int) -> dict[str, Any] | None:
    with _connect() as conn:
        return conn.execute("SELECT * FROM branches WHERE id = %s", (branch_id,)).fetchone()


def import_branches(rows: list[tuple[str, str]]) -> tuple[int, int]:
    added = 0
    updated = 0
    with _connect() as conn:
        for code, name in rows:
            existing = conn.execute(
                "SELECT id, branch_name, status FROM branches WHERE branch_code = %s",
                (code,),
            ).fetchone()
            if existing:
                conn.execute(
                    """
                    UPDATE branches SET branch_name = %s, status = 'active'
                    WHERE branch_code = %s
                    """,
                    (name, code),
                )
                updated += 1
            else:
                conn.execute(
                    "INSERT INTO branches (branch_code, branch_name) VALUES (%s, %s)",
                    (code, name),
                )
                added += 1
    return added, updated


def create_employee(
    telegram_id: int,
    phone: str,
    full_name: str,
    branch_id: int,
    position: str,
    shift: str,
) -> dict[str, Any]:
    with _connect() as conn:
        row = conn.execute(
            """
            INSERT INTO employees
                (telegram_id, phone, full_name, branch_id, position, shift)
            VALUES (%s, %s, %s, %s, %s, %s)
            ON CONFLICT (telegram_id) DO UPDATE SET
                phone = EXCLUDED.phone,
                full_name = EXCLUDED.full_name,
                branch_id = EXCLUDED.branch_id,
                position = EXCLUDED.position,
                shift = EXCLUDED.shift,
                status = 'active'
            RETURNING id
            """,
            (telegram_id, phone, full_name, branch_id, position, shift),
        ).fetchone()
    return get_employee_by_id(row["id"])  # type: ignore[index]


def update_employee_branch_shift(
    employee_id: int,
    branch_id: int,
    shift: str,
) -> dict[str, Any] | None:
    """Xodim filialini (va smenasini) yangilaydi — keyingi davomat shu filialga hisoblanadi."""
    if shift not in {"1", "2"}:
        raise ValueError("shift 1 yoki 2 bo‘lishi kerak")
    with _connect() as conn:
        row = conn.execute(
            """
            UPDATE employees
            SET branch_id = %s, shift = %s
            WHERE id = %s AND status = 'active'
            RETURNING id
            """,
            (branch_id, shift, employee_id),
        ).fetchone()
    if not row:
        return None
    return get_employee_by_id(int(row["id"]))


def get_attendance_for_date(employee_id: int, attendance_date: date) -> dict[str, Any] | None:
    with _connect() as conn:
        return conn.execute(
            """
            SELECT * FROM attendance
            WHERE employee_id = %s AND attendance_date = %s
            """,
            (employee_id, attendance_date),
        ).fetchone()


def record_arrival(
    employee_id: int,
    attendance_date: date,
    arrival_at: datetime,
    arrival_status: str,
    late_minutes: int,
) -> dict[str, Any] | None:
    with _connect() as conn:
        existing = conn.execute(
            """
            SELECT * FROM attendance
            WHERE employee_id = %s AND attendance_date = %s
            """,
            (employee_id, attendance_date),
        ).fetchone()
        if existing and existing["arrival_at"] is not None:
            return existing
        if existing:
            return conn.execute(
                """
                UPDATE attendance
                SET arrival_at = %s, arrival_status = %s, late_minutes = %s,
                    updated_at = NOW()
                WHERE id = %s
                RETURNING *
                """,
                (arrival_at, arrival_status, late_minutes, existing["id"]),
            ).fetchone()
        return conn.execute(
            """
            INSERT INTO attendance
                (employee_id, attendance_date, arrival_at, arrival_status, late_minutes)
            VALUES (%s, %s, %s, %s, %s)
            RETURNING *
            """,
            (employee_id, attendance_date, arrival_at, arrival_status, late_minutes),
        ).fetchone()


def find_open_attendance(employee_id: int, today: date) -> dict[str, Any] | None:
    with _connect() as conn:
        return conn.execute(
            """
            SELECT * FROM attendance
            WHERE employee_id = %s
              AND attendance_date BETWEEN %s AND %s
              AND arrival_at IS NOT NULL
              AND departure_at IS NULL
            ORDER BY attendance_date DESC
            LIMIT 1
            """,
            (employee_id, today - timedelta(days=1), today),
        ).fetchone()


def record_departure(
    attendance_id: int,
    departure_at: datetime,
    worked_minutes: int,
    departure_status: str,
    early_leave_minutes: int,
) -> dict[str, Any] | None:
    with _connect() as conn:
        return conn.execute(
            """
            UPDATE attendance
            SET departure_at = %s, worked_minutes = %s, departure_status = %s,
                early_leave_minutes = %s, updated_at = NOW()
            WHERE id = %s AND departure_at IS NULL
            RETURNING *
            """,
            (
                departure_at,
                worked_minutes,
                departure_status,
                early_leave_minutes,
                attendance_id,
            ),
        ).fetchone()


def get_today_rows(today: date) -> list[dict[str, Any]]:
    with _connect() as conn:
        return conn.execute(
            """
            SELECT e.*, b.branch_code, b.branch_name, a.*
            FROM employees e
            JOIN branches b ON b.id = e.branch_id
            LEFT JOIN attendance a
                ON a.employee_id = e.id AND a.attendance_date = %s
            WHERE e.status = 'active'
            ORDER BY b.branch_name, e.full_name
            """,
            (today,),
        ).fetchall()


def get_all_employees() -> list[dict[str, Any]]:
    with _connect() as conn:
        return conn.execute(
            """
            SELECT e.*, b.branch_code, b.branch_name
            FROM employees e
            JOIN branches b ON b.id = e.branch_id
            WHERE e.status = 'active'
            ORDER BY b.branch_name, e.full_name
            """
        ).fetchall()


def get_attendance_range(start: date, end: date) -> list[dict[str, Any]]:
    with _connect() as conn:
        return conn.execute(
            """
            SELECT a.*, e.telegram_id, e.full_name, e.position, e.shift,
                   e.branch_id, b.branch_code, b.branch_name
            FROM attendance a
            JOIN employees e ON e.id = a.employee_id
            JOIN branches b ON b.id = e.branch_id
            WHERE a.attendance_date BETWEEN %s AND %s
            ORDER BY a.attendance_date, b.branch_name, e.full_name
            """,
            (start, end),
        ).fetchall()


def get_monthly_stats(start: date, end: date) -> list[dict[str, Any]]:
    with _connect() as conn:
        return conn.execute(
            """
            SELECT e.id AS employee_id, e.full_name, e.position, e.shift,
                   b.branch_code, b.branch_name,
                   COUNT(a.id) FILTER (WHERE a.arrival_at IS NOT NULL)::int AS arrived,
                   COUNT(a.id) FILTER (WHERE a.arrival_status = 'LATE')::int AS late_days,
                   COALESCE(SUM(a.late_minutes), 0)::int AS late_minutes,
                   COALESCE(SUM(a.early_leave_minutes), 0)::int AS early_minutes,
                   COALESCE(SUM(a.worked_minutes), 0)::int AS worked_minutes
            FROM employees e
            JOIN branches b ON b.id = e.branch_id
            LEFT JOIN attendance a
                ON a.employee_id = e.id
               AND a.attendance_date BETWEEN %s AND %s
            WHERE e.status = 'active'
            GROUP BY e.id, e.full_name, e.position, e.shift,
                     b.branch_code, b.branch_name
            ORDER BY b.branch_name, e.full_name
            """,
            (start, end),
        ).fetchall()


def list_pending_requests() -> list[dict[str, Any]]:
    with _connect() as conn:
        return conn.execute(
            """
            SELECT r.*, e.telegram_id, e.full_name, e.position,
                   b.branch_name
            FROM change_requests r
            JOIN employees e ON e.id = r.employee_id
            JOIN branches b ON b.id = e.branch_id
            WHERE r.status = 'pending'
            ORDER BY r.created_at
            """
        ).fetchall()


def create_change_request(employee_id: int, request_type: str, requested_value: str) -> None:
    with _connect() as conn:
        conn.execute(
            """
            INSERT INTO change_requests (employee_id, request_type, requested_value)
            VALUES (%s, %s, %s)
            """,
            (employee_id, request_type, requested_value),
        )


def branch_employee_count(branch_id: int) -> int:
    with _connect() as conn:
        row = conn.execute(
            "SELECT COUNT(*)::int AS count FROM employees WHERE branch_id = %s AND status = 'active'",
            (branch_id,),
        ).fetchone()
        return int(row["count"])  # type: ignore[index]


WEEKDAY_NAMES_UZ = {
    0: "Dushanba",
    1: "Seshanba",
    2: "Chorshanba",
    3: "Payshanba",
    4: "Juma",
    5: "Shanba",
    6: "Yakshanba",
}


def get_work_weekdays() -> list[dict[str, Any]]:
    with _connect() as conn:
        rows = conn.execute(
            "SELECT weekday, is_work FROM work_weekdays ORDER BY weekday"
        ).fetchall()
    by_day = {int(row["weekday"]): bool(row["is_work"]) for row in rows}
    return [
        {
            "weekday": day,
            "name": WEEKDAY_NAMES_UZ[day],
            "is_work": by_day.get(day, day != 6),
        }
        for day in range(7)
    ]


def set_work_weekday(weekday: int, is_work: bool) -> None:
    if weekday < 0 or weekday > 6:
        raise ValueError("weekday 0..6 oralig‘ida bo‘lishi kerak")
    with _connect() as conn:
        conn.execute(
            """
            INSERT INTO work_weekdays (weekday, is_work)
            VALUES (%s, %s)
            ON CONFLICT (weekday) DO UPDATE SET is_work = EXCLUDED.is_work
            """,
            (weekday, is_work),
        )


def list_work_day_overrides(limit: int = 40) -> list[dict[str, Any]]:
    with _connect() as conn:
        return conn.execute(
            """
            SELECT work_date, is_work, note, updated_at
            FROM work_day_overrides
            ORDER BY work_date DESC
            LIMIT %s
            """,
            (limit,),
        ).fetchall()


def set_work_day_override(work_date: date, is_work: bool, note: str = "") -> None:
    with _connect() as conn:
        conn.execute(
            """
            INSERT INTO work_day_overrides (work_date, is_work, note, updated_at)
            VALUES (%s, %s, %s, NOW())
            ON CONFLICT (work_date) DO UPDATE SET
                is_work = EXCLUDED.is_work,
                note = EXCLUDED.note,
                updated_at = NOW()
            """,
            (work_date, is_work, note.strip()),
        )


def clear_work_day_override(work_date: date) -> bool:
    with _connect() as conn:
        row = conn.execute(
            "DELETE FROM work_day_overrides WHERE work_date = %s RETURNING work_date",
            (work_date,),
        ).fetchone()
    return row is not None


def is_work_day(day: date) -> bool:
    """Admin belgilagan ish kunimi yoki yo‘q."""
    with _connect() as conn:
        override = conn.execute(
            "SELECT is_work FROM work_day_overrides WHERE work_date = %s",
            (day,),
        ).fetchone()
        if override is not None:
            return bool(override["is_work"])
        weekday = day.weekday()  # Mon=0 .. Sun=6
        row = conn.execute(
            "SELECT is_work FROM work_weekdays WHERE weekday = %s",
            (weekday,),
        ).fetchone()
    if row is None:
        return weekday != 6
    return bool(row["is_work"])


def list_work_days(start: date, end: date) -> list[date]:
    with _connect() as conn:
        weekday_rows = conn.execute(
            "SELECT weekday, is_work FROM work_weekdays"
        ).fetchall()
        override_rows = conn.execute(
            """
            SELECT work_date, is_work
            FROM work_day_overrides
            WHERE work_date BETWEEN %s AND %s
            """,
            (start, end),
        ).fetchall()
    by_weekday = {int(row["weekday"]): bool(row["is_work"]) for row in weekday_rows}
    overrides = {row["work_date"]: bool(row["is_work"]) for row in override_rows}
    days: list[date] = []
    current = start
    while current <= end:
        if current in overrides:
            if overrides[current]:
                days.append(current)
        elif by_weekday.get(current.weekday(), current.weekday() != 6):
            days.append(current)
        current += timedelta(days=1)
    return days
