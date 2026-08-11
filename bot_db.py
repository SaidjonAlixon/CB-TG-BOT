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


def init_db() -> None:
    with _connect() as conn:
        conn.execute(SCHEMA_SQL)
        admin_ids = os.environ.get("ADMIN_TELEGRAM_IDS", "")
        for raw_id in admin_ids.split(","):
            raw_id = raw_id.strip()
            if raw_id.isdigit():
                conn.execute(
                    """
                    INSERT INTO admins (telegram_id, full_name)
                    VALUES (%s, 'Administrator')
                    ON CONFLICT (telegram_id) DO UPDATE SET status = 'active'
                    """,
                    (int(raw_id),),
                )


def is_admin(telegram_id: int) -> bool:
    with _connect() as conn:
        row = conn.execute(
            "SELECT 1 FROM admins WHERE telegram_id = %s AND status = 'active'",
            (telegram_id,),
        ).fetchone()
        return row is not None


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
