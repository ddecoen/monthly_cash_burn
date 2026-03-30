"""Monthly Cash Burn Tracker — Flask + SQLite backend.

Tracks monthly revenue and categorised expenses, calculates burn rate,
and estimates remaining runway given a cash balance.
"""

from __future__ import annotations

import os
import sqlite3
from contextlib import contextmanager

from flask import Flask, g, jsonify, render_template, request

# ---------------------------------------------------------------------------
# App / config
# ---------------------------------------------------------------------------

app = Flask(__name__)

DATABASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cash_burn.db")

# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------

EXPENSE_CATEGORIES = [
    "payroll",
    "rent",
    "marketing",
    "software",
    "professional_services",
    "other_expenses",
]

SCHEMA = """\
CREATE TABLE IF NOT EXISTS entries (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    month                 INTEGER NOT NULL CHECK (month BETWEEN 1 AND 12),
    year                  INTEGER NOT NULL,
    revenue               REAL    DEFAULT 0,
    payroll               REAL    DEFAULT 0,
    rent                  REAL    DEFAULT 0,
    marketing             REAL    DEFAULT 0,
    software              REAL    DEFAULT 0,
    professional_services REAL    DEFAULT 0,
    other_expenses        REAL    DEFAULT 0,
    created_at            TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
"""


def get_db() -> sqlite3.Connection:
    """Return a per-request database connection stored on Flask's *g*."""
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA journal_mode=WAL;")
    return g.db


@app.teardown_appcontext
def close_db(_exc: BaseException | None = None) -> None:
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db() -> None:
    """Create tables if they don't already exist."""
    conn = sqlite3.connect(DATABASE)
    try:
        conn.executescript(SCHEMA)
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Conversion helpers
# ---------------------------------------------------------------------------


def _row_to_dict(row: sqlite3.Row) -> dict:
    """Convert a Row into a plain dict with computed totals."""
    d = dict(row)
    total_expenses = sum(d.get(cat, 0) or 0 for cat in EXPENSE_CATEGORIES)
    revenue = d.get("revenue", 0) or 0
    d["total_expenses"] = round(total_expenses, 2)
    d["cash_burn"] = round(total_expenses - revenue, 2)
    return d


# ---------------------------------------------------------------------------
# Routes — pages
# ---------------------------------------------------------------------------


@app.route("/")
def index():
    """Serve the main single-page HTML interface."""
    return render_template("index.html")


# ---------------------------------------------------------------------------
# Routes — REST API
# ---------------------------------------------------------------------------


@app.route("/api/entries", methods=["GET"])
def list_entries():
    """Return every entry, sorted by year then month ascending."""
    db = get_db()
    rows = db.execute(
        "SELECT * FROM entries ORDER BY year ASC, month ASC"
    ).fetchall()
    return jsonify([_row_to_dict(r) for r in rows])


@app.route("/api/entries", methods=["POST"])
def create_entry():
    """Create a new monthly entry from a JSON body.

    Required fields: *month*, *year*.
    Optional numeric fields: *revenue* and each expense category.
    """
    data = request.get_json(silent=True)
    if data is None:
        return jsonify({"error": "Request body must be valid JSON."}), 400

    # -- Validate required fields ------------------------------------------------
    month = data.get("month")
    year = data.get("year")

    errors: list[str] = []
    if month is None:
        errors.append("'month' is required.")
    else:
        try:
            month = int(month)
            if not 1 <= month <= 12:
                errors.append("'month' must be between 1 and 12.")
        except (TypeError, ValueError):
            errors.append("'month' must be an integer.")

    if year is None:
        errors.append("'year' is required.")
    else:
        try:
            year = int(year)
            if year < 1900 or year > 2100:
                errors.append("'year' must be between 1900 and 2100.")
        except (TypeError, ValueError):
            errors.append("'year' must be an integer.")

    if errors:
        return jsonify({"error": "Validation failed.", "details": errors}), 400

    # -- Coerce optional numeric fields ------------------------------------------
    def _float_field(name: str) -> float:
        raw = data.get(name, 0)
        try:
            return float(raw) if raw is not None else 0.0
        except (TypeError, ValueError):
            return 0.0

    revenue = _float_field("revenue")
    payroll = _float_field("payroll")
    rent = _float_field("rent")
    marketing = _float_field("marketing")
    software = _float_field("software")
    professional_services = _float_field("professional_services")
    other_expenses = _float_field("other_expenses")

    # -- Insert ------------------------------------------------------------------
    db = get_db()
    cursor = db.execute(
        """
        INSERT INTO entries
            (month, year, revenue, payroll, rent, marketing,
             software, professional_services, other_expenses)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            month,
            year,
            revenue,
            payroll,
            rent,
            marketing,
            software,
            professional_services,
            other_expenses,
        ),
    )
    db.commit()

    row = db.execute(
        "SELECT * FROM entries WHERE id = ?", (cursor.lastrowid,)
    ).fetchone()
    return jsonify(_row_to_dict(row)), 201


@app.route("/api/entries/<int:entry_id>", methods=["DELETE"])
def delete_entry(entry_id: int):
    """Delete an entry by its primary key."""
    db = get_db()
    row = db.execute("SELECT id FROM entries WHERE id = ?", (entry_id,)).fetchone()
    if row is None:
        return jsonify({"error": f"Entry {entry_id} not found."}), 404
    db.execute("DELETE FROM entries WHERE id = ?", (entry_id,))
    db.commit()
    return jsonify({"message": f"Entry {entry_id} deleted."}), 200


@app.route("/api/summary", methods=["GET"])
def summary():
    """Return aggregate statistics across all stored entries.

    Accepts an optional *cash_balance* query parameter to compute an
    estimated runway (months until cash runs out at the average burn
    rate).
    """
    db = get_db()
    rows = db.execute(
        "SELECT * FROM entries ORDER BY year ASC, month ASC"
    ).fetchall()

    if not rows:
        return jsonify(
            {
                "months_of_data": 0,
                "total_revenue": 0,
                "total_expenses": 0,
                "total_burn": 0,
                "avg_monthly_burn": 0,
                "avg_monthly_revenue": 0,
                "avg_monthly_expenses": 0,
                "runway_months": None,
                "cash_balance": None,
            }
        )

    entries = [_row_to_dict(r) for r in rows]
    months_of_data = len(entries)
    total_revenue = round(sum(e["revenue"] for e in entries), 2)
    total_expenses = round(sum(e["total_expenses"] for e in entries), 2)
    total_burn = round(sum(e["cash_burn"] for e in entries), 2)
    avg_burn = round(total_burn / months_of_data, 2)
    avg_revenue = round(total_revenue / months_of_data, 2)
    avg_expenses = round(total_expenses / months_of_data, 2)

    # -- Optional runway estimate -----------------------------------------------
    cash_balance_raw = request.args.get("cash_balance") or request.args.get("cash_on_hand")
    cash_balance: float | None = None
    runway_months: float | None = None
    if cash_balance_raw is not None:
        try:
            cash_balance = float(cash_balance_raw)
            if avg_burn > 0:
                runway_months = round(cash_balance / avg_burn, 1)
            else:
                # Not burning cash — runway is effectively infinite.
                runway_months = None
        except (TypeError, ValueError):
            pass

    return jsonify(
        {
            "months_of_data": months_of_data,
            "total_revenue": total_revenue,
            "total_expenses": total_expenses,
            "total_burn": total_burn,
            "avg_monthly_burn": avg_burn,
            "avg_monthly_revenue": avg_revenue,
            "avg_monthly_expenses": avg_expenses,
            "runway_months": runway_months,
            "cash_balance": cash_balance,
        }
    )


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=True)
