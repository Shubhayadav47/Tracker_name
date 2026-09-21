# Team Ticket Tracking Platform

A web-based platform for uploading daily Excel ticket exports, tracking status changes, and managing team workload.

## Setup & Run

### 1. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 2. Start the server

```bash
python app.py
```

Open **http://localhost:5000** in your browser.

---

## How It Works

### First-Time Upload (Column Mapping)

1. Go to **Upload Excel** in the sidebar.
2. Select your `.xlsx` or `.xls` file.
3. Click **Preview & Map Columns** — a mapping dialog appears.
4. Map each Excel column header to the correct system field.
   - `ticket_id` is **required** (used as the unique identifier).
   - All other fields are optional.
   - Set unmapped columns to *— Skip —*.
5. Click **Save Mapping & Upload** — your mapping is saved for future uploads.

### Daily Uploads (After First Time)

1. Select your updated Excel file.
2. Click **Upload with Saved Mapping** — no remapping needed.
3. A summary dialog shows:
   - ✅ New tickets added
   - 🔄 Existing tickets updated
   - ⏭ Unchanged tickets skipped
   - 📦 Resolved/archived tickets
   - ❌ Invalid rows/errors

### Smart Import Logic

| Scenario | Action |
|---|---|
| Ticket ID not in DB | **Add** as new active ticket |
| Ticket ID exists, same status | **Skip** (no duplicate) |
| Ticket ID exists, status changed | **Update** existing ticket |
| Ticket status is Resolved/Closed/Done | **Archive** (removed from active dashboard) |
| Already-archived ticket seen again | Stays archived |

### Resolved Statuses

The following status values trigger archiving (case-insensitive):
`Resolved`, `Closed`, `Done`, `Completed`, `Fixed`, `Cancelled`, `Canceled`

You can add more by editing `RESOLVED_STATUSES` in `app.py`.

---

## Excel File Format

Any `.xlsx` or `.xls` file with a header row is supported. Example columns:

| Ticket ID | Title | Status | Priority | Assigned To | Category | Created Date |
|---|---|---|---|---|---|---|
| INC1001 | Login issue | Open | High | Alice | Auth | 2024-01-15 |
| INC1002 | Slow response | In Progress | Medium | Bob | Performance | 2024-01-16 |

Column names are flexible — you define the mapping on first use.

---

## Project Structure

```
ticket-tracker/
├── app.py              # Flask backend, routes, import logic
├── requirements.txt    # Python dependencies
├── instance/
│   └── tickets.db      # SQLite database (auto-created)
├── uploads/            # Temporary Excel file storage
└── templates/
    └── index.html      # Single-page frontend
```
