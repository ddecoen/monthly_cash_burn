# Monthly Cash Burn Tracker

A lightweight Flask + SQLite web application for tracking monthly revenue,
categorised expenses, and cash burn rate. Designed for startups and small
teams that need a quick, self-hosted view of their financial runway.

## Features

- **Monthly data entry** — record revenue and expenses broken into six
  categories: payroll, rent, marketing, software/tools, professional
  services, and other.
- **Automatic burn calculation** — each entry computes
  `cash_burn = total_expenses − revenue`. Positive means you're burning
  cash; negative means the month was profitable.
- **Summary dashboard** — view aggregate stats including average monthly
  burn, total burn, and months of data on file.
- **Runway estimator** — enter your current cash balance to see how many
  months of runway remain at the average burn rate.
- **REST API** — all data is available as JSON so you can integrate with
  spreadsheets, dashboards, or scripts.
- **SQLite storage** — zero-config database created automatically on first
  run. No external database server required.

## Quickstart

```bash
cd /home/coder/monthly_cash_burn

# Install dependencies (ideally in a virtualenv).
pip install -r requirements.txt

# Run the app.
python app.py
```

The server starts on **http://0.0.0.0:5000**. Open that URL in a browser
to use the web interface.

## API Endpoints

| Method | Path                   | Description                                  |
|--------|------------------------|----------------------------------------------|
| GET    | `/`                    | Serve the main HTML page.                    |
| GET    | `/api/entries`         | Return all entries as JSON (sorted by date). |
| POST   | `/api/entries`         | Create a new entry (JSON body).              |
| DELETE | `/api/entries/<id>`    | Delete an entry by ID.                       |
| GET    | `/api/summary`         | Aggregate stats; optional `?cash_balance=N`. |

### POST /api/entries — example body

```json
{
  "month": 3,
  "year": 2026,
  "revenue": 42000,
  "payroll": 60000,
  "rent": 5000,
  "marketing": 8000,
  "software": 3200,
  "professional_services": 2500,
  "other_expenses": 1000
}
```

### GET /api/summary?cash_balance=500000 — example response

```json
{
  "months_of_data": 6,
  "total_revenue": 240000,
  "total_expenses": 470000,
  "total_burn": 230000,
  "avg_monthly_burn": 38333.33,
  "avg_monthly_revenue": 40000,
  "avg_monthly_expenses": 78333.33,
  "runway_months": 13.0,
  "cash_balance": 500000
}
```

## Database Schema

Stored in `cash_burn.db` (created automatically).

| Column                | Type      | Notes                          |
|-----------------------|-----------|--------------------------------|
| id                    | INTEGER   | Primary key, autoincrement.    |
| month                 | INTEGER   | 1–12.                          |
| year                  | INTEGER   | Four-digit year.               |
| revenue               | REAL      | Monthly revenue, default 0.    |
| payroll               | REAL      | Payroll expense, default 0.    |
| rent                  | REAL      | Rent expense, default 0.       |
| marketing             | REAL      | Marketing expense, default 0.  |
| software              | REAL      | Software/tools, default 0.     |
| professional_services | REAL      | Professional services, def. 0. |
| other_expenses        | REAL      | Other expenses, default 0.     |
| created_at            | TIMESTAMP | Auto-set on insert.            |

## Project Structure

```
monthly_cash_burn/
├── app.py               # Flask application (backend + routes)
├── cash_burn.db         # SQLite database (created at runtime)
├── requirements.txt     # Python dependencies
├── README.md            # This file
├── templates/
│   └── index.html       # Main page template
└── static/
    ├── style.css        # Styles
    └── app.js           # Client-side JavaScript
```
