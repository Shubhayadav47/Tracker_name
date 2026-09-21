"""
Team Ticket Tracking Platform — Flask backend
"""
import json
import os
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, request, jsonify, render_template, g

# ── App setup ──────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = Path(__file__).parent / "uploads"
app.config["DATABASE"] = Path(__file__).parent / "instance" / "tickets.db"
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB

ALLOWED_EXTENSIONS = {"xlsx", "xls"}

# Status values that mean the ticket is resolved/closed
RESOLVED_STATUSES = {"resolved", "closed", "done", "completed", "fixed", "cancelled", "canceled"}

# ── DB helpers ─────────────────────────────────────────────────────────────
def get_db():
    db = getattr(g, "_database", None)
    if db is None:
        db = g._database = sqlite3.connect(
            app.config["DATABASE"], detect_types=sqlite3.PARSE_DECLTYPES
        )
        db.row_factory = sqlite3.Row
    return db


@app.teardown_appcontext
def close_db(exc):
    db = getattr(g, "_database", None)
    if db is not None:
        db.close()


def init_db():
    db = get_db()
    db.executescript("""
        CREATE TABLE IF NOT EXISTS tickets (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            ticket_id       TEXT    NOT NULL UNIQUE,
            title           TEXT,
            description     TEXT,
            created_date    TEXT,
            priority        TEXT,
            assigned_to     TEXT,
            category        TEXT,
            status          TEXT,
            last_updated    TEXT,
            percentage      TEXT,
            extra_fields    TEXT,   -- JSON blob for unmapped columns
            is_resolved     INTEGER NOT NULL DEFAULT 0,
            archived_at     TEXT,
            imported_at     TEXT    NOT NULL
        );

        CREATE TABLE IF NOT EXISTS column_mapping (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            excel_col   TEXT NOT NULL UNIQUE,
            field_name  TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS upload_log (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            uploaded_at     TEXT NOT NULL,
            filename        TEXT,
            added           INTEGER DEFAULT 0,
            updated         INTEGER DEFAULT 0,
            skipped         INTEGER DEFAULT 0,
            archived        INTEGER DEFAULT 0,
            errors          INTEGER DEFAULT 0,
            total_rows      INTEGER DEFAULT 0
        );
    """)
    db.commit()
    # Migration: add percentage column to existing databases
    existing_cols = [row[1] for row in db.execute("PRAGMA table_info(tickets)").fetchall()]
    if "percentage" not in existing_cols:
        db.execute("ALTER TABLE tickets ADD COLUMN percentage TEXT")
        db.commit()


# ── Utility ────────────────────────────────────────────────────────────────
def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def is_resolved_status(status: str) -> bool:
    if not status:
        return False
    return status.strip().lower() in RESOLVED_STATUSES


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


# Core field names used internally
CORE_FIELDS = {
    "ticket_id", "title", "description", "created_date",
    "priority", "assigned_to", "category", "status", "last_updated", "percentage"
}

# Default guesses for auto-mapping when no saved mapping exists
DEFAULT_GUESSES = {
    "ticket id":       "ticket_id",
    "ticket number":   "ticket_id",
    "ticket_id":       "ticket_id",
    "ticketid":        "ticket_id",
    "id":              "ticket_id",
    "incident":        "ticket_id",
    "incident number": "ticket_id",
    "title":           "title",
    "subject":         "title",
    "ticket title":    "title",
    "summary":         "title",
    "description":     "description",
    "details":         "description",
    "created date":    "created_date",
    "created_date":    "created_date",
    "creation date":   "created_date",
    "open date":       "created_date",
    "priority":        "priority",
    "severity":        "priority",
    "assigned to":     "assigned_to",
    "assignee":        "assigned_to",
    "owner":           "assigned_to",
    "assigned_to":     "assigned_to",
    "category":        "category",
    "type":            "category",
    "ticket type":     "category",
    "status":          "status",
    "state":           "status",
    "last updated":    "last_updated",
    "last_updated":    "last_updated",
    "updated date":    "last_updated",
    "modified date":   "last_updated",
    "percentage":      "percentage",
    "% complete":      "percentage",
    "percent":         "percentage",
    "completion %":    "percentage",
    "progress":        "percentage",
}


def load_saved_mapping(db) -> dict:
    """Return {excel_col: field_name} from DB."""
    rows = db.execute("SELECT excel_col, field_name FROM column_mapping").fetchall()
    return {r["excel_col"]: r["field_name"] for r in rows}


def save_mapping(db, mapping: dict):
    """Upsert {excel_col: field_name} into DB."""
    for excel_col, field_name in mapping.items():
        db.execute(
            "INSERT INTO column_mapping (excel_col, field_name) VALUES (?, ?) "
            "ON CONFLICT(excel_col) DO UPDATE SET field_name=excluded.field_name",
            (excel_col, field_name),
        )
    db.commit()


def read_excel_rows(filepath: str) -> tuple[list[str], list[list]]:
    """Return (headers, rows) from an xlsx or xls file."""
    ext = filepath.rsplit(".", 1)[-1].lower()
    if ext == "xlsx":
        import openpyxl
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    else:
        import xlrd
        wb = xlrd.open_workbook(filepath)
        ws = wb.sheet_by_index(0)
        rows = [ws.row_values(i) for i in range(ws.nrows)]

    if not rows:
        return [], []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    data_rows = [list(r) for r in rows[1:]]
    return headers, data_rows


def apply_mapping(headers: list[str], row: list, mapping: dict) -> dict:
    """
    Map a raw row (list) into a dict using the column mapping.
    Unmapped columns go into 'extra_fields'.
    """
    record = {}
    extra = {}
    for i, header in enumerate(headers):
        value = row[i] if i < len(row) else None
        if value is None:
            value = ""
        else:
            value = str(value).strip()
        field = mapping.get(header)
        if field and field in CORE_FIELDS:
            record[field] = value
        else:
            extra[header] = value
    record["extra_fields"] = json.dumps(extra, ensure_ascii=False)
    return record


# ── Routes ─────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/mapping", methods=["GET"])
def get_mapping():
    db = get_db()
    mapping = load_saved_mapping(db)
    return jsonify({"mapping": mapping})


@app.route("/api/mapping", methods=["POST"])
def set_mapping():
    data = request.get_json(force=True)
    mapping = data.get("mapping", {})
    if not mapping:
        return jsonify({"error": "Empty mapping"}), 400
    db = get_db()
    save_mapping(db, mapping)
    return jsonify({"ok": True})


@app.route("/api/preview-headers", methods=["POST"])
def preview_headers():
    """Upload a file and return its headers + any auto-detected mapping."""
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    f = request.files["file"]
    if not f.filename or not allowed_file(f.filename):
        return jsonify({"error": "Invalid file type. Use .xlsx or .xls"}), 400

    save_path = app.config["UPLOAD_FOLDER"] / f.filename
    f.save(str(save_path))

    try:
        headers, _ = read_excel_rows(str(save_path))
    except Exception as e:
        return jsonify({"error": f"Could not read file: {e}"}), 400

    db = get_db()
    saved = load_saved_mapping(db)

    # Build a suggested mapping: saved wins, then auto-guess
    suggested = {}
    for h in headers:
        if h in saved:
            suggested[h] = saved[h]
        else:
            suggested[h] = DEFAULT_GUESSES.get(h.lower(), "")

    return jsonify({
        "headers": headers,
        "suggested_mapping": suggested,
        "filename": f.filename,
        "core_fields": sorted(CORE_FIELDS),
    })


@app.route("/api/upload", methods=["POST"])
def upload_tickets():
    """
    Full upload + upsert.
    Expects multipart form: file + optional mapping JSON string.
    """
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    f = request.files["file"]
    if not f.filename or not allowed_file(f.filename):
        return jsonify({"error": "Invalid file type. Use .xlsx or .xls"}), 400

    # Accept override mapping from form field
    mapping_override = request.form.get("mapping")
    if mapping_override:
        try:
            mapping_override = json.loads(mapping_override)
        except Exception:
            mapping_override = None

    save_path = app.config["UPLOAD_FOLDER"] / f.filename
    f.save(str(save_path))

    try:
        headers, data_rows = read_excel_rows(str(save_path))
    except Exception as e:
        return jsonify({"error": f"Could not read file: {e}"}), 400

    db = get_db()
    mapping = mapping_override or load_saved_mapping(db)

    if not mapping:
        return jsonify({
            "error": "No column mapping configured. Please map columns first.",
            "headers": headers,
            "needs_mapping": True,
        }), 422

    # Persist mapping if it came as an override
    if mapping_override:
        save_mapping(db, mapping_override)

    # Validate ticket_id column exists
    mapped_fields = set(mapping.values())
    if "ticket_id" not in mapped_fields:
        return jsonify({"error": "Mapping must include a 'ticket_id' field."}), 422

    stats = {"added": 0, "updated": 0, "skipped": 0, "archived": 0, "errors": 0, "total_rows": 0}
    now = now_iso()

    for raw_row in data_rows:
        # Skip completely empty rows
        if all((v is None or str(v).strip() == "") for v in raw_row):
            continue
        stats["total_rows"] += 1
        try:
            record = apply_mapping(headers, raw_row, mapping)
            tid = record.get("ticket_id", "").strip()
            if not tid:
                stats["errors"] += 1
                continue

            resolved = is_resolved_status(record.get("status", ""))

            existing = db.execute(
                "SELECT id, status, is_resolved FROM tickets WHERE ticket_id = ?", (tid,)
            ).fetchone()

            if existing is None:
                if resolved:
                    # New ticket already resolved — store as archived
                    db.execute(
                        """INSERT INTO tickets
                           (ticket_id, title, description, created_date, priority,
                            assigned_to, category, status, last_updated, percentage, extra_fields,
                            is_resolved, archived_at, imported_at)
                           VALUES (?,?,?,?,?,?,?,?,?,?,?,1,?,?)""",
                        (
                            tid,
                            record.get("title", ""),
                            record.get("description", ""),
                            record.get("created_date", ""),
                            record.get("priority", ""),
                            record.get("assigned_to", ""),
                            record.get("category", ""),
                            record.get("status", ""),
                            record.get("last_updated", ""),
                            record.get("percentage", ""),
                            record["extra_fields"],
                            now,
                            now,
                        ),
                    )
                    stats["archived"] += 1
                else:
                    db.execute(
                        """INSERT INTO tickets
                           (ticket_id, title, description, created_date, priority,
                            assigned_to, category, status, last_updated, percentage, extra_fields,
                            is_resolved, imported_at)
                           VALUES (?,?,?,?,?,?,?,?,?,?,?,0,?)""",
                        (
                            tid,
                            record.get("title", ""),
                            record.get("description", ""),
                            record.get("created_date", ""),
                            record.get("priority", ""),
                            record.get("assigned_to", ""),
                            record.get("category", ""),
                            record.get("status", ""),
                            record.get("last_updated", ""),
                            record.get("percentage", ""),
                            record["extra_fields"],
                            now,
                        ),
                    )
                    stats["added"] += 1
            else:
                # Existing ticket — compare & update
                old_status = (existing["status"] or "").strip().lower()
                new_status = (record.get("status", "") or "").strip().lower()
                same_data = (old_status == new_status)  # simplified change detection

                if same_data and not resolved:
                    stats["skipped"] += 1
                else:
                    archived_at = now if resolved else None
                    db.execute(
                        """UPDATE tickets SET
                               title=?, description=?, created_date=?, priority=?,
                               assigned_to=?, category=?, status=?, last_updated=?,
                               percentage=?, extra_fields=?, is_resolved=?, archived_at=?
                           WHERE ticket_id=?""",
                        (
                            record.get("title", ""),
                            record.get("description", ""),
                            record.get("created_date", ""),
                            record.get("priority", ""),
                            record.get("assigned_to", ""),
                            record.get("category", ""),
                            record.get("status", ""),
                            record.get("last_updated", ""),
                            record.get("percentage", ""),
                            record["extra_fields"],
                            1 if resolved else 0,
                            archived_at,
                            tid,
                        ),
                    )
                    if resolved:
                        stats["archived"] += 1
                    else:
                        stats["updated"] += 1

        except Exception as exc:
            stats["errors"] += 1
            app.logger.warning("Row error: %s", exc)

    db.execute(
        """INSERT INTO upload_log
           (uploaded_at, filename, added, updated, skipped, archived, errors, total_rows)
           VALUES (?,?,?,?,?,?,?,?)""",
        (now, f.filename, stats["added"], stats["updated"],
         stats["skipped"], stats["archived"], stats["errors"], stats["total_rows"]),
    )
    db.commit()

    return jsonify({"ok": True, "summary": stats})


@app.route("/api/tickets", methods=["GET"])
def list_tickets():
    """Return active (non-resolved) tickets."""
    db = get_db()
    search = request.args.get("q", "").strip()
    priority = request.args.get("priority", "").strip()
    assigned = request.args.get("assigned_to", "").strip()

    query = "SELECT * FROM tickets WHERE is_resolved = 0"
    params = []
    if search:
        query += " AND (ticket_id LIKE ? OR title LIKE ? OR assigned_to LIKE ?)"
        s = f"%{search}%"
        params += [s, s, s]
    if priority:
        query += " AND LOWER(priority) = LOWER(?)"
        params.append(priority)
    if assigned:
        query += " AND LOWER(assigned_to) = LOWER(?)"
        params.append(assigned)
    query += " ORDER BY imported_at DESC"

    rows = db.execute(query, params).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/tickets/resolved", methods=["GET"])
def list_resolved():
    """Return archived/resolved tickets."""
    db = get_db()
    search = request.args.get("q", "").strip()
    query = "SELECT * FROM tickets WHERE is_resolved = 1"
    params = []
    if search:
        query += " AND (ticket_id LIKE ? OR title LIKE ?)"
        s = f"%{search}%"
        params += [s, s]
    query += " ORDER BY archived_at DESC"
    rows = db.execute(query, params).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/tickets/<ticket_id>", methods=["GET"])
def get_ticket(ticket_id):
    db = get_db()
    row = db.execute("SELECT * FROM tickets WHERE ticket_id = ?", (ticket_id,)).fetchone()
    if row is None:
        return jsonify({"error": "Not found"}), 404
    return jsonify(dict(row))


@app.route("/api/tickets/<ticket_id>", methods=["PATCH"])
def update_ticket(ticket_id):
    """Manual field update from the UI."""
    data = request.get_json(force=True)
    db = get_db()
    row = db.execute("SELECT * FROM tickets WHERE ticket_id = ?", (ticket_id,)).fetchone()
    if row is None:
        return jsonify({"error": "Not found"}), 404

    allowed = {"title", "description", "priority", "assigned_to", "category", "status", "last_updated", "percentage"}
    updates = {k: v for k, v in data.items() if k in allowed}
    if not updates:
        return jsonify({"error": "No valid fields to update"}), 400

    resolved = is_resolved_status(updates.get("status", row["status"]))
    updates["is_resolved"] = 1 if resolved else 0
    updates["archived_at"] = now_iso() if resolved else None

    set_clause = ", ".join(f"{k}=?" for k in updates)
    db.execute(
        f"UPDATE tickets SET {set_clause} WHERE ticket_id=?",
        list(updates.values()) + [ticket_id],
    )
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/stats", methods=["GET"])
def stats():
    db = get_db()
    total_active = db.execute("SELECT COUNT(*) FROM tickets WHERE is_resolved=0").fetchone()[0]
    total_resolved = db.execute("SELECT COUNT(*) FROM tickets WHERE is_resolved=1").fetchone()[0]
    by_priority = db.execute(
        "SELECT COALESCE(NULLIF(priority,''),'Unknown') as priority, COUNT(*) as cnt "
        "FROM tickets WHERE is_resolved=0 GROUP BY priority ORDER BY cnt DESC"
    ).fetchall()
    by_assignee = db.execute(
        "SELECT COALESCE(NULLIF(assigned_to,''),'Unassigned') as assigned_to, COUNT(*) as cnt "
        "FROM tickets WHERE is_resolved=0 GROUP BY assigned_to ORDER BY cnt DESC LIMIT 10"
    ).fetchall()
    by_status = db.execute(
        "SELECT COALESCE(NULLIF(status,''),'Unknown') as status, COUNT(*) as cnt "
        "FROM tickets WHERE is_resolved=0 GROUP BY status ORDER BY cnt DESC"
    ).fetchall()
    last_upload = db.execute(
        "SELECT * FROM upload_log ORDER BY id DESC LIMIT 1"
    ).fetchone()
    return jsonify({
        "total_active": total_active,
        "total_resolved": total_resolved,
        "by_priority": [dict(r) for r in by_priority],
        "by_assignee": [dict(r) for r in by_assignee],
        "by_status": [dict(r) for r in by_status],
        "last_upload": dict(last_upload) if last_upload else None,
    })


@app.route("/api/upload-history", methods=["GET"])
def upload_history():
    db = get_db()
    rows = db.execute("SELECT * FROM upload_log ORDER BY id DESC LIMIT 20").fetchall()
    return jsonify([dict(r) for r in rows])


# ── Startup ────────────────────────────────────────────────────────────────
with app.app_context():
    init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)
