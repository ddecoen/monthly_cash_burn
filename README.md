# Monthly Cash Burn Dashboard

A Flask + SQLite web app that imports NetSuite **Statement of Cash Flows**
(SCF) Excel exports and displays monthly burn analysis with quarterly views.

## Features

- **Excel upload** — drag-and-drop your NetSuite monthly SCF `.xlsx` export.
  The parser auto-detects the period (month/year) and extracts all line items.
- **Quarterly view** — side-by-side monthly columns for any quarter, formatted
  like a real financial statement (section headers, subtotals, accounting-style
  negatives in parentheses).
- **Burn analysis** — average monthly burn, operating burn, cash position, and
  estimated runway based on your current cash on hand.
- **Trend charts** — Chart.js visualizations of burn trend and cash position
  over time.
- **Re-upload safe** — uploading the same month/year replaces the old data.
- **SQLite storage** — zero-config, created automatically on first run.

## Quickstart

```bash
git clone https://github.com/ddecoen/monthly_cash_burn.git
cd monthly_cash_burn
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

Open **http://localhost:5000** in your browser.

## Updating to the Latest Version

```bash
cd monthly_cash_burn
source .venv/bin/activate
git pull origin main
pip install -r requirements.txt   # pick up any new dependencies
python3 app.py
```

> **Note:** Your data is stored in `cash_burn.db` locally and is not
> affected by pulling updates. The `.gitignore` keeps it out of version
> control.

## Usage

1. In NetSuite, export your monthly SCF as `.xlsx` (one month per file).
2. Drag the file into the upload zone (or click to browse).
3. Select a year and quarter to view the quarterly cash flow table.
4. Upload additional months — they appear as columns in the quarterly view.
5. Enter your current cash on hand to see estimated runway.

## Expected Excel Format

The parser expects the standard NetSuite SCF layout:

| Row | Column A                              | Column B          |
|-----|---------------------------------------|-------------------|
| 1   | Company name                          |                   |
| 2   | Report title                          |                   |
| 3   | (amounts in thousands)                |                   |
| 4   | (unaudited)                           |                   |
| 5   | One Month Ended December 31, 2025     |                   |
| 6   | *(blank)*                             |                   |
| 7   | Description                           | Amount (thousands)|
| 8+  | Line items with amounts in column B   |                   |

Negative amounts can be plain negatives or in parentheses: `(3,518,233)`.

## API Endpoints

| Method | Path                      | Description                                    |
|--------|---------------------------|------------------------------------------------|
| GET    | `/`                       | Serve the dashboard UI.                        |
| POST   | `/api/upload`             | Upload a `.xlsx` SCF file (multipart form).    |
| GET    | `/api/periods`            | List all uploaded periods.                     |
| GET    | `/api/period/<id>`        | Get a period with all line items.              |
| DELETE | `/api/period/<id>`        | Delete a period.                               |
| GET    | `/api/quarterly`          | Quarterly view. Params: `year`, `quarter`.     |
| GET    | `/api/burn-summary`       | Burn metrics. Optional: `?cash_on_hand=N`.     |

## Project Structure

```
monthly_cash_burn/
├── app.py               # Flask application (backend + routes + parser)
├── cash_burn.db         # SQLite database (created at runtime)
├── requirements.txt     # Python dependencies (flask, openpyxl)
├── README.md            # This file
├── uploads/             # Stored Excel files (gitignored)
└── templates/
    └── index.html       # Dashboard UI (self-contained)
```
