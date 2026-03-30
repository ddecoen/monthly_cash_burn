"""Monthly Cash Burn Tracker — Flask + SQLite backend.

Imports NetSuite Statement of Cash Flows (SCF) Excel files, stores
parsed line items per period, and exposes a JSON API for burn analysis,
quarterly views, and runway estimation.
"""

from __future__ import annotations

import os
import re
import sqlite3
from datetime import datetime

from io import BytesIO

from flask import Flask, g, jsonify, render_template, request, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# App / config
# ---------------------------------------------------------------------------

app = Flask(__name__)

DATABASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cash_burn.db")
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")

# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------

SCHEMA = """\
CREATE TABLE IF NOT EXISTS periods (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    month INTEGER NOT NULL,
    year INTEGER NOT NULL,
    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(month, year)
);

CREATE TABLE IF NOT EXISTS line_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    period_id INTEGER NOT NULL REFERENCES periods(id) ON DELETE CASCADE,
    section TEXT NOT NULL,
    description TEXT NOT NULL,
    amount REAL DEFAULT 0,
    is_subtotal INTEGER DEFAULT 0,
    row_order INTEGER NOT NULL
);
"""


def get_db() -> sqlite3.Connection:
    """Return a per-request database connection stored on Flask's *g*."""
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA journal_mode=WAL;")
        g.db.execute("PRAGMA foreign_keys=ON;")
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
        conn.execute("PRAGMA foreign_keys=ON;")
        conn.executescript(SCHEMA)
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Excel parsing helpers
# ---------------------------------------------------------------------------

# Month name → number lookup.
MONTH_MAP: dict[str, int] = {
    "january": 1,
    "february": 2,
    "march": 3,
    "april": 4,
    "may": 5,
    "june": 6,
    "july": 7,
    "august": 8,
    "september": 9,
    "october": 10,
    "november": 11,
    "december": 12,
}

# Regex to pull "Month DD, YYYY" from the period string (row 5).
_PERIOD_RE = re.compile(
    r"(January|February|March|April|May|June|July|August|September"
    r"|October|November|December)\s+\d{1,2},\s*(\d{4})",
    re.IGNORECASE,
)

# Descriptions that mark section headers (no amount on these rows).
_SECTION_HEADERS: dict[str, str] = {
    "cash flows from operating activities": "operating",
    "cash flows from investing activities": "investing",
    "cash flows from financing activities": "financing",
}

# Descriptions that are subtotals — stored with is_subtotal=1.
_SUBTOTAL_DESCRIPTIONS: set[str] = {
    "net cash used in operating activities",
    "net cash provided by operating activities",
    "net cash used in investing activities",
    "net cash provided by investing activities",
    "net cash used in financing activities",
    "net cash provided by financing activities",
    "net decrease in cash and cash equivalents",
    "net increase in cash and cash equivalents",
    "cash and cash equivalents at beginning of period",
    "cash and cash equivalents at end of period",
}


def _parse_amount(raw: object) -> float:
    """Convert a cell value to a float.

    Handles:
    - ``None`` / empty → 0
    - Plain numbers → float
    - Strings with commas and/or parentheses for negatives
    """
    if raw is None:
        return 0.0
    if isinstance(raw, (int, float)):
        return float(raw)
    text = str(raw).strip()
    if not text or text == "-":
        return 0.0

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]

    text = text.replace(",", "").replace("$", "").strip()
    try:
        value = float(text)
    except ValueError:
        return 0.0
    return -value if negative else value


def _parse_period_string(text: str) -> tuple[int, int]:
    """Extract (month, year) from a NetSuite period string.

    Examples:
        "One Month Ended December 31, 2025" → (12, 2025)
        "One Month Ended January 31, 2026"  → (1, 2026)

    Raises ``ValueError`` when the string cannot be parsed.
    """
    match = _PERIOD_RE.search(text)
    if not match:
        raise ValueError(f"Cannot parse period from: {text!r}")
    month_name = match.group(1).lower()
    year = int(match.group(2))
    month = MONTH_MAP[month_name]
    return month, year


def _is_section_header(desc: str) -> str | None:
    """Return the section key if *desc* is a section header, else ``None``."""
    normalised = desc.strip().rstrip(":").lower()
    return _SECTION_HEADERS.get(normalised)


def _is_subsection(desc: str) -> bool:
    """Return True if the description is a subsection label (no amount)."""
    lower = desc.strip().rstrip(":").lower()
    return lower in (
        "adjustments to reconcile net loss to cash from operating activities",
        "adjustments to reconcile net income to cash from operating activities",
        "changes in operating assets and liabilities",
    )


def parse_scf_excel(filepath: str) -> dict:
    """Parse a NetSuite Statement of Cash Flows .xlsx file.

    Returns a dict with keys ``month``, ``year``, and ``line_items``
    (a list of dicts with ``section``, ``description``, ``amount``,
    ``is_subtotal``, ``row_order``).
    """
    wb = load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active

    # Row 5 contains the period string.
    period_cell = ws.cell(row=5, column=1).value
    if period_cell is None:
        raise ValueError("Row 5 (period string) is empty.")
    month, year = _parse_period_string(str(period_cell))

    # Walk data rows starting at row 8.
    current_section = "operating"
    line_items: list[dict] = []
    row_order = 0

    for row in ws.iter_rows(min_row=8, max_col=2, values_only=True):
        desc_raw = row[0]
        if desc_raw is None:
            continue
        desc = str(desc_raw).strip()
        if not desc:
            continue

        # Check for section header transitions.
        section_key = _is_section_header(desc)
        if section_key is not None:
            current_section = section_key
            continue

        # Skip subsection labels.
        if _is_subsection(desc):
            continue

        amount = _parse_amount(row[1])

        # Rows after the three main sections are "summary" items.
        desc_lower = desc.lower()
        section = current_section
        if desc_lower in (
            "net decrease in cash and cash equivalents",
            "net increase in cash and cash equivalents",
            "cash and cash equivalents at beginning of period",
            "cash and cash equivalents at end of period",
        ):
            section = "summary"

        is_subtotal = 1 if desc_lower in _SUBTOTAL_DESCRIPTIONS else 0

        row_order += 1
        line_items.append(
            {
                "section": section,
                "description": desc,
                "amount": amount,
                "is_subtotal": is_subtotal,
                "row_order": row_order,
            }
        )

    wb.close()
    return {"month": month, "year": year, "line_items": line_items}


# ---------------------------------------------------------------------------
# Database storage helpers
# ---------------------------------------------------------------------------


def _store_period(
    db: sqlite3.Connection,
    month: int,
    year: int,
    line_items: list[dict],
) -> int:
    """Insert (or replace) a period and its line items.

    If the month/year already exists the old data is deleted first.
    Returns the ``period_id``.
    """
    # Delete existing period for this month/year (cascade removes items).
    existing = db.execute(
        "SELECT id FROM periods WHERE month = ? AND year = ?",
        (month, year),
    ).fetchone()
    if existing is not None:
        db.execute("DELETE FROM periods WHERE id = ?", (existing["id"],))

    cursor = db.execute(
        "INSERT INTO periods (month, year) VALUES (?, ?)",
        (month, year),
    )
    period_id = cursor.lastrowid

    for item in line_items:
        db.execute(
            """
            INSERT INTO line_items
                (period_id, section, description, amount, is_subtotal, row_order)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                period_id,
                item["section"],
                item["description"],
                item["amount"],
                item["is_subtotal"],
                item["row_order"],
            ),
        )

    db.commit()
    return period_id


# ---------------------------------------------------------------------------
# Serialisation helpers
# ---------------------------------------------------------------------------

_MONTH_NAMES = [
    "",
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]

_MONTH_ABBREVS = [
    "", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]

_QUARTER_MONTH_RANGES: dict[int, list[int]] = {
    1: [1, 2, 3],
    2: [4, 5, 6],
    3: [7, 8, 9],
    4: [10, 11, 12],
}


def _period_to_dict(row: sqlite3.Row) -> dict:
    d = dict(row)
    d["month_name"] = _MONTH_NAMES[d["month"]]
    return d


def _line_item_to_dict(row: sqlite3.Row) -> dict:
    return dict(row)


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


@app.route("/api/upload", methods=["POST"])
def upload_file():
    """Accept an Excel (.xlsx) upload, parse it, and store the data."""
    if "file" not in request.files:
        return jsonify({"error": "No file part in request."}), 400

    file = request.files["file"]
    if file.filename == "" or file.filename is None:
        return jsonify({"error": "No file selected."}), 400

    if not file.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Only .xlsx files are accepted."}), 400

    # Ensure upload directory exists.
    os.makedirs(UPLOAD_DIR, exist_ok=True)

    # Save with a timestamped name to avoid collisions.
    safe_name = re.sub(r"[^\w.\-]", "_", file.filename)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    save_path = os.path.join(UPLOAD_DIR, f"{ts}_{safe_name}")
    file.save(save_path)

    try:
        parsed = parse_scf_excel(save_path)
    except Exception as exc:
        return jsonify({"error": f"Failed to parse file: {exc}"}), 400

    db = get_db()
    period_id = _store_period(
        db,
        parsed["month"],
        parsed["year"],
        parsed["line_items"],
    )

    return jsonify(
        {
            "period_id": period_id,
            "month": parsed["month"],
            "year": parsed["year"],
            "month_name": _MONTH_NAMES[parsed["month"]],
            "line_items": parsed["line_items"],
        }
    ), 201


@app.route("/api/periods", methods=["GET"])
def list_periods():
    """Return all periods sorted by year then month."""
    db = get_db()
    rows = db.execute(
        "SELECT * FROM periods ORDER BY year ASC, month ASC"
    ).fetchall()
    return jsonify([_period_to_dict(r) for r in rows])


@app.route("/api/period/<int:period_id>", methods=["GET"])
def get_period(period_id: int):
    """Return a single period with all its line items."""
    db = get_db()
    period = db.execute(
        "SELECT * FROM periods WHERE id = ?", (period_id,)
    ).fetchone()
    if period is None:
        return jsonify({"error": f"Period {period_id} not found."}), 404

    items = db.execute(
        "SELECT * FROM line_items WHERE period_id = ? ORDER BY row_order ASC",
        (period_id,),
    ).fetchall()

    result = _period_to_dict(period)
    result["line_items"] = [_line_item_to_dict(i) for i in items]
    return jsonify(result)


@app.route("/api/period/<int:period_id>", methods=["DELETE"])
def delete_period(period_id: int):
    """Delete a period and its line items."""
    db = get_db()
    period = db.execute(
        "SELECT id FROM periods WHERE id = ?", (period_id,)
    ).fetchone()
    if period is None:
        return jsonify({"error": f"Period {period_id} not found."}), 404

    db.execute("DELETE FROM periods WHERE id = ?", (period_id,))
    db.commit()
    return jsonify({"message": f"Period {period_id} deleted."}), 200


def _build_quarterly_data(year: int, quarter: int) -> dict:
    """Build quarterly SCF data. Returns the same structure as the JSON endpoint."""
    months = _QUARTER_MONTH_RANGES[quarter]

    db = get_db()

    # Fetch periods for the requested quarter.
    placeholders = ",".join("?" for _ in months)
    periods = db.execute(
        f"SELECT * FROM periods WHERE year = ? AND month IN ({placeholders}) "
        "ORDER BY month ASC",
        [year] + months,
    ).fetchall()

    period_map: dict[int, sqlite3.Row] = {p["month"]: p for p in periods}

    # Build a stable ordered list of descriptions from whichever periods
    # exist, preserving row_order.
    description_order: list[str] = []
    seen_descriptions: set[str] = set()
    section_map: dict[str, str] = {}
    subtotal_map: dict[str, int] = {}

    for m in months:
        if m not in period_map:
            continue
        items = db.execute(
            "SELECT * FROM line_items WHERE period_id = ? ORDER BY row_order ASC",
            (period_map[m]["id"],),
        ).fetchall()
        for item in items:
            desc = item["description"]
            if desc not in seen_descriptions:
                description_order.append(desc)
                seen_descriptions.add(desc)
                section_map[desc] = item["section"]
                subtotal_map[desc] = item["is_subtotal"]

    # Build month metadata the frontend expects.
    months_meta = [
        {
            "year": year,
            "month": m,
            "available": m in period_map,
            "label": f"{_MONTH_NAMES[m]} {year}",
        }
        for m in months
    ]

    # Map (section, is_subtotal) to the row type the frontend uses.
    section_header_descs = {
        "cash flows from operating activities",
        "cash flows from investing activities",
        "cash flows from financing activities",
    }
    subsection_descs = {
        "adjustments to reconcile net loss to cash from operating activities",
        "adjustments to reconcile net income to cash from operating activities",
        "changes in operating assets and liabilities",
    }

    def _row_type(desc: str, section: str, is_sub: int) -> str:
        dl = desc.strip().rstrip(":").lower()
        if dl in section_header_descs:
            return "section_header"
        if dl in subsection_descs:
            return "subsection"
        if section == "summary":
            return "grand_total"
        if is_sub:
            return "subtotal"
        return "line"

    # Insert section headers and subsections into the row list so the
    # frontend can render them.  Walk the ordered descriptions and
    # emit synthetic header rows when the section changes.
    section_titles = {
        "operating": "Cash flows from operating activities",
        "investing": "Cash flows from investing activities",
        "financing": "Cash flows from financing activities",
    }

    rows: list[dict] = []
    prev_section: str | None = None
    emitted_adjustments = False
    emitted_changes = False

    for desc in description_order:
        section = section_map[desc]
        is_sub = subtotal_map[desc]

        # Emit section header on section transition.
        if section != prev_section and section in section_titles:
            rows.append(
                {
                    "description": section_titles[section],
                    "type": "section_header",
                    "values": [None] * len(months),
                }
            )
            prev_section = section
            emitted_adjustments = False
            emitted_changes = False

        # Emit subsection headers within operating.
        if section == "operating" and not is_sub:
            dl = desc.lower()
            if (
                not emitted_adjustments
                and dl == "depreciation and amortization expense"
            ):
                rows.append(
                    {
                        "description": "Adjustments to reconcile net loss to cash from operating activities:",
                        "type": "subsection",
                        "values": [None] * len(months),
                    }
                )
                emitted_adjustments = True
            if not emitted_changes and dl == "accounts receivable":
                rows.append(
                    {
                        "description": "Changes in operating assets and liabilities:",
                        "type": "subsection",
                        "values": [None] * len(months),
                    }
                )
                emitted_changes = True

        # Build per-month values (None when month not uploaded).
        values: list[float | None] = []
        for m in months:
            if m not in period_map:
                values.append(None)
            else:
                row = db.execute(
                    "SELECT amount FROM line_items "
                    "WHERE period_id = ? AND description = ?",
                    (period_map[m]["id"], desc),
                ).fetchone()
                values.append(row["amount"] if row else 0.0)

        available_vals = [v for v in values if v is not None]
        total = round(sum(available_vals), 2) if available_vals else None

        rows.append(
            {
                "description": desc,
                "type": _row_type(desc, section, is_sub),
                "values": values,
                "total": total,
            }
        )

    return {
        "year": year,
        "quarter": quarter,
        "months": months_meta,
        "rows": rows,
    }


@app.route("/api/quarterly", methods=["GET"])
def quarterly_view():
    """Return a quarterly view of line items for column display.

    Query params:
        year   — required, e.g. 2025
        quarter — required, 1-4 (Q1=Jan-Mar … Q4=Oct-Dec)

    Returns each line item description mapped to an array of monthly
    values plus a "Total" sum column.
    """
    year_raw = request.args.get("year")
    quarter_raw = request.args.get("quarter")
    if year_raw is None or quarter_raw is None:
        return jsonify({"error": "'year' and 'quarter' are required."}), 400

    try:
        year = int(year_raw)
        quarter = int(quarter_raw)
    except (TypeError, ValueError):
        return jsonify({"error": "'year' and 'quarter' must be integers."}), 400

    if quarter not in (1, 2, 3, 4):
        return jsonify({"error": "'quarter' must be 1, 2, 3, or 4."}), 400

    return jsonify(_build_quarterly_data(year, quarter))


@app.route("/api/export", methods=["GET"])
def export_excel():
    """Export quarterly SCF data as a professionally formatted Excel workbook.

    Query params:
        year    — required, e.g. 2025
        quarter — required, 1-4 (Q1=Jan-Mar … Q4=Oct-Dec)

    Returns an .xlsx file as a downloadable attachment.
    """
    year_raw = request.args.get("year")
    quarter_raw = request.args.get("quarter")
    if year_raw is None or quarter_raw is None:
        return jsonify({"error": "'year' and 'quarter' are required."}), 400

    try:
        year = int(year_raw)
        quarter = int(quarter_raw)
    except (TypeError, ValueError):
        return jsonify({"error": "'year' and 'quarter' must be integers."}), 400

    if quarter not in (1, 2, 3, 4):
        return jsonify({"error": "'quarter' must be 1, 2, 3, or 4."}), 400

    data = _build_quarterly_data(year, quarter)
    months_meta = data["months"]
    data_rows = data["rows"]

    # Total columns: Description + one per month + Quarter Total.
    num_cols = 1 + len(months_meta) + 1

    wb = Workbook()
    ws = wb.active
    ws.title = f"SCF Q{quarter} {year}"

    # -- Row 1: Company title ------------------------------------------------
    ws.merge_cells(
        start_row=1, start_column=1, end_row=1, end_column=num_cols,
    )
    title_cell = ws.cell(row=1, column=1, value="Statement of Cash Flows")
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center")

    # -- Row 2: Period label -------------------------------------------------
    first_m = months_meta[0]["month"]
    last_m = months_meta[-1]["month"]
    period_label = (
        f"Q{quarter} {year} "
        f"({_MONTH_ABBREVS[first_m]} - {_MONTH_ABBREVS[last_m]})"
    )
    ws.merge_cells(
        start_row=2, start_column=1, end_row=2, end_column=num_cols,
    )
    period_cell = ws.cell(row=2, column=1, value=period_label)
    period_cell.font = Font(italic=True, size=11)
    period_cell.alignment = Alignment(horizontal="center")

    # -- Row 3: Amounts note -------------------------------------------------
    ws.merge_cells(
        start_row=3, start_column=1, end_row=3, end_column=num_cols,
    )
    note_cell = ws.cell(row=3, column=1, value="(amounts in thousands)")
    note_cell.font = Font(italic=True, size=9)
    note_cell.alignment = Alignment(horizontal="center")

    # -- Row 4: blank --------------------------------------------------------

    # -- Row 5: Column headers -----------------------------------------------
    header_font = Font(bold=True)
    header_fill = PatternFill(
        start_color="D6E4F0", end_color="D6E4F0", fill_type="solid",
    )
    header_border = Border(bottom=Side(style="thin"))

    headers = ["Description"]
    for m in months_meta:
        headers.append(f"{_MONTH_ABBREVS[m['month']]} {m['year']}")
    headers.append("Quarter Total")

    for col_idx, hdr in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=col_idx, value=hdr)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = header_border
        if col_idx > 1:
            cell.alignment = Alignment(horizontal="center")

    # -- Column widths -------------------------------------------------------
    ws.column_dimensions["A"].width = 45
    for col_idx in range(2, num_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 18

    # -- Reusable styles for data rows ---------------------------------------
    num_fmt = '#,##0;(#,##0);"-"'
    section_fill = PatternFill(
        start_color="E8EDF2", end_color="E8EDF2", fill_type="solid",
    )
    subtotal_fill = PatternFill(
        start_color="F0F4F8", end_color="F0F4F8", fill_type="solid",
    )
    subtotal_border = Border(top=Side(style="thin"))
    grand_border = Border(
        top=Side(style="thin"), bottom=Side(style="double"),
    )
    right_align = Alignment(horizontal="right")
    center_align = Alignment(horizontal="center")

    def _write_value(cell, val, fmt=num_fmt):
        """Write a numeric value or an em-dash for None."""
        if val is None:
            cell.value = "\u2014"
            cell.alignment = center_align
        else:
            cell.value = val
            cell.number_format = fmt
            cell.alignment = right_align

    # -- Data rows (starting at row 6) --------------------------------------
    current_row = 6

    for row_data in data_rows:
        desc = row_data["description"]
        row_type = row_data["type"]
        values = row_data.get("values", [])
        total = row_data.get("total")

        desc_cell = ws.cell(row=current_row, column=1)

        if row_type == "section_header":
            desc_cell.value = desc
            desc_cell.font = Font(bold=True)
            desc_cell.fill = section_fill
            for ci in range(2, num_cols + 1):
                ws.cell(row=current_row, column=ci).fill = section_fill

        elif row_type == "subsection":
            desc_cell.value = f"  {desc}"
            desc_cell.font = Font(italic=True)

        elif row_type == "line":
            desc_cell.value = f"  {desc}"
            for i, val in enumerate(values):
                _write_value(ws.cell(row=current_row, column=2 + i), val)
            _write_value(
                ws.cell(row=current_row, column=num_cols), total,
            )

        elif row_type == "subtotal":
            desc_cell.value = desc
            desc_cell.font = Font(bold=True)
            desc_cell.fill = subtotal_fill
            for i, val in enumerate(values):
                cell = ws.cell(row=current_row, column=2 + i)
                cell.fill = subtotal_fill
                cell.border = subtotal_border
                _write_value(cell, val)
            total_cell = ws.cell(row=current_row, column=num_cols)
            total_cell.fill = subtotal_fill
            total_cell.border = subtotal_border
            _write_value(total_cell, total)

        elif row_type == "grand_total":
            desc_cell.value = desc
            desc_cell.font = Font(bold=True)
            desc_lower = desc.lower()
            is_beginning = (
                desc_lower
                == "cash and cash equivalents at beginning of period"
            )
            is_ending = (
                desc_lower
                == "cash and cash equivalents at end of period"
            )

            for i, val in enumerate(values):
                cell = ws.cell(row=current_row, column=2 + i)
                cell.border = grand_border
                _write_value(cell, val)

            total_cell = ws.cell(row=current_row, column=num_cols)
            total_cell.border = grand_border
            if is_beginning:
                total_cell.value = "\u2014"
                total_cell.alignment = center_align
            elif is_ending:
                # Show last available month's value instead of the sum.
                last_val = None
                for v in reversed(values):
                    if v is not None:
                        last_val = v
                        break
                _write_value(total_cell, last_val)
            else:
                _write_value(total_cell, total)

        current_row += 1

    # -- Stream the workbook back as a download ------------------------------
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"SCF_Q{quarter}_{year}.xlsx"
    return send_file(
        output,
        mimetype=(
            "application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet"
        ),
        as_attachment=True,
        download_name=filename,
    )


@app.route("/api/burn-summary", methods=["GET"])
def burn_summary():
    """Return burn analysis across all stored periods.

    Optional query param ``cash_on_hand`` for runway estimation.
    """
    db = get_db()
    periods = db.execute(
        "SELECT * FROM periods ORDER BY year ASC, month ASC"
    ).fetchall()

    if not periods:
        return jsonify(
            {
                "months_of_data": 0,
                "monthly_burn_trend": [],
                "avg_monthly_burn": 0,
                "avg_net_cash_change": 0,
                "avg_operating_burn": 0,
                "cash_position": None,
                "cash_position_over_time": [],
                "runway_months": None,
            }
        )

    monthly_burn_trend: list[dict] = []
    cash_positions: list[dict] = []
    operating_burns: list[float] = []

    for period in periods:
        pid = period["id"]
        month_name = _MONTH_NAMES[period["month"]]
        label = f"{month_name} {period['year']}"

        # Net cash change = "Net decrease in cash and cash equivalents"
        # or "Net increase …". Fall back to summing the three section
        # subtotals if the summary row is missing.
        net_row = db.execute(
            "SELECT amount FROM line_items "
            "WHERE period_id = ? AND section = 'summary' "
            "AND LOWER(description) LIKE 'net % in cash and cash equivalents'",
            (pid,),
        ).fetchone()

        if net_row is not None:
            net_change = net_row["amount"]
        else:
            # Fallback: sum the three section subtotals.
            subtotals = db.execute(
                "SELECT amount FROM line_items "
                "WHERE period_id = ? AND is_subtotal = 1 "
                "AND section IN ('operating', 'investing', 'financing')",
                (pid,),
            ).fetchall()
            net_change = sum(r["amount"] for r in subtotals)

        # Operating burn = net cash used in operating activities.
        op_row = db.execute(
            "SELECT amount FROM line_items "
            "WHERE period_id = ? AND is_subtotal = 1 "
            "AND section = 'operating'",
            (pid,),
        ).fetchone()
        if op_row is not None:
            operating_burns.append(op_row["amount"])

        monthly_burn_trend.append(
            {
                "period": label,
                "month": period["month"],
                "year": period["year"],
                "net_cash_change": round(net_change, 2),
            }
        )

        # Cash position = ending cash for the period.
        cash_end_row = db.execute(
            "SELECT amount FROM line_items "
            "WHERE period_id = ? "
            "AND LOWER(description) = 'cash and cash equivalents at end of period'",
            (pid,),
        ).fetchone()

        cash_positions.append(
            {
                "period": label,
                "month": period["month"],
                "year": period["year"],
                "cash_position": round(
                    cash_end_row["amount"] if cash_end_row else 0.0, 2
                ),
            }
        )

    # Average monthly burn (negative net_cash_change means burning cash,
    # so average burn is the negated mean).
    total_net = sum(item["net_cash_change"] for item in monthly_burn_trend)
    months_of_data = len(monthly_burn_trend)
    avg_monthly_burn = round(-total_net / months_of_data, 2)
    avg_net_cash_change = round(total_net / months_of_data, 2)
    avg_operating_burn = (
        round(sum(operating_burns) / len(operating_burns), 2)
        if operating_burns
        else 0
    )
    latest_cash = (
        cash_positions[-1]["cash_position"] if cash_positions else None
    )

    # Runway estimate.
    cash_on_hand_raw = request.args.get("cash_on_hand")
    runway_months: float | None = None
    if cash_on_hand_raw is not None:
        try:
            cash_on_hand = float(cash_on_hand_raw)
            if avg_monthly_burn > 0:
                runway_months = round(cash_on_hand / avg_monthly_burn, 1)
        except (TypeError, ValueError):
            pass

    return jsonify(
        {
            "months_of_data": months_of_data,
            "monthly_burn_trend": monthly_burn_trend,
            "avg_monthly_burn": avg_monthly_burn,
            "avg_net_cash_change": avg_net_cash_change,
            "avg_operating_burn": avg_operating_burn,
            "cash_position": latest_cash,
            "cash_position_over_time": cash_positions,
            "runway_months": runway_months,
        }
    )


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=True)
