"""Persistent weekly snapshots for the Build Your Own workspace."""

import json
import sqlite3
import threading
import traceback
from contextlib import contextmanager
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import pandas as pd

DEMO = "__demo_claims__"
DEMO_NAME = "Medicare Claims - sample history"
_worker = None
_worker_lock = threading.Lock()


@contextmanager
def connect(path):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path), timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        with conn:
            yield conn
    finally:
        conn.close()


def initialize(path):
    with connect(path) as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS schedules (
                owner TEXT NOT NULL, dataset TEXT NOT NULL, key_column TEXT NOT NULL,
                next_run TEXT NOT NULL, error TEXT NOT NULL DEFAULT '',
                PRIMARY KEY (owner, dataset)
            );
            CREATE TABLE IF NOT EXISTS snapshots (
                id INTEGER PRIMARY KEY AUTOINCREMENT, owner TEXT NOT NULL,
                dataset TEXT NOT NULL, captured_at TEXT NOT NULL,
                key_column TEXT NOT NULL, data TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS snapshot_history
                ON snapshots(owner, dataset, captured_at, id);
            CREATE TABLE IF NOT EXISTS recipients (
                owner TEXT NOT NULL, dataset TEXT NOT NULL, email TEXT NOT NULL COLLATE NOCASE,
                PRIMARY KEY(owner, dataset, email)
            );
            CREATE TABLE IF NOT EXISTS email_outbox (
                id INTEGER PRIMARY KEY AUTOINCREMENT, owner TEXT NOT NULL, dataset TEXT NOT NULL,
                snapshot_id INTEGER NOT NULL, email TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'pending',
                attempts INTEGER NOT NULL DEFAULT 0, next_attempt TEXT NOT NULL,
                last_error TEXT NOT NULL DEFAULT '', sent_at TEXT,
                UNIQUE(snapshot_id, email)
            );
        """)
        # Seed once, without ever touching the user's dataset library.
        conn.execute("BEGIN IMMEDIATE")
        if not conn.execute("SELECT 1 FROM snapshots WHERE owner=? AND dataset=?", (DEMO, DEMO)).fetchone():
            seed_demo(conn)


def seed_demo(conn):
    columns = ["Claim ID", "Bene ID", "Claim From DT", "Claim Through DT",
               "Bene Pmt Amt", "Claim Pmt Amt", "Clm Effective Date", "Clm Obsolete dt",
               "Source ID", "Provider ID", "Diagnosis Code", "Claim Type",
               "Claim Final Action Indicator"]
    def row(number, payment, diagnosis):
        return [f"CLM-{number:04}", f"BENE-{number:03}", "2026-01-05", "2026-01-08",
                40, payment, "2026-01-09", None, "CMS-A", f"PRV-{number:03}",
                diagnosis, "Inpatient", "Y"]
    rows = [row(1001, 1200, "I10"), row(1002, 850, "E11.9"), row(1003, 475, "J44.9"),
            row(1004, 2100, "M54.5"), row(1005, 725, "I25.1")]
    weeks = [json.loads(json.dumps(rows))]
    rows[0][5] = 1325
    rows.append(row(1006, 990, "E78.5"))
    weeks.append(json.loads(json.dumps(rows)))
    rows[1][10] = None
    rows[2][4] = -25
    rows.append(row(1007, 640, "J18.9"))
    weeks.append(json.loads(json.dumps(rows)))
    rows[1][10] = "E11.65"
    rows[2][4] = 25
    rows[3][12] = "N"
    rows.pop(4)
    rows.append(row(1008, 1510, "I50.9"))
    weeks.append(json.loads(json.dumps(rows)))
    first = datetime.now(timezone.utc).date() - timedelta(days=21)
    for i, records in enumerate(weeks):
        captured = datetime.combine(first + timedelta(days=7 * i), datetime.min.time(), timezone.utc)
        conn.execute("INSERT INTO snapshots(owner,dataset,captured_at,key_column,data) VALUES(?,?,?,?,?)",
                     (DEMO, DEMO, captured.isoformat(), "Claim ID",
                      json.dumps({"columns": columns, "rows": [dict(zip(columns, r)) for r in records]})))


def frame_data(frame, key):
    columns = [str(c) for c in frame.columns]
    if len(set(columns)) != len(columns):
        raise ValueError("Column names must be unique.")
    if key not in columns:
        raise ValueError(f"Identity column '{key}' is missing from the dataset.")
    # Serialize through pandas to normalize timestamps, numpy values and NaN.
    rows = json.loads(frame.to_json(orient="records", date_format="iso"))
    values = [str(r[key]).strip() if r[key] is not None else "" for r in rows]
    if any(not v for v in values) or len(set(values)) != len(values):
        raise ValueError("Identity values must be non-empty and unique in every snapshot.")
    return {"columns": columns, "rows": rows}


def save_schedule(path, owner, dataset, key, start, frame):
    try:
        start_date = date.fromisoformat(start)
    except (TypeError, ValueError):
        raise ValueError("Choose a valid first refresh date.")
    if start_date < datetime.now(timezone.utc).date():
        raise ValueError("The first refresh date cannot be in the past (UTC).")
    data = frame_data(frame, key)
    next_run = datetime.combine(start_date, datetime.min.time(), timezone.utc).isoformat()
    now = datetime.now(timezone.utc).isoformat()
    with connect(path) as conn:
        conn.execute("BEGIN IMMEDIATE")
        existing = conn.execute("SELECT key_column FROM snapshots WHERE owner=? AND dataset=? ORDER BY id DESC LIMIT 1",
                                (owner, dataset)).fetchone()
        if existing and existing["key_column"] != key:
            raise ValueError("The identity column cannot change while this dataset has history.")
        if not existing:
            conn.execute("INSERT INTO snapshots(owner,dataset,captured_at,key_column,data) VALUES(?,?,?,?,?)",
                         (owner, dataset, now, key, json.dumps(data)))
        conn.execute("""INSERT INTO schedules(owner,dataset,key_column,next_run,error) VALUES(?,?,?,?, '')
                        ON CONFLICT(owner,dataset) DO UPDATE SET key_column=excluded.key_column,
                        next_run=excluded.next_run,error=''""", (owner, dataset, key, next_run))


def invalid_cells(row):
    invalid = set()
    for column in ("Bene Pmt Amt", "Claim Pmt Amt"):
        value = row.get(column)
        if value is not None:
            try:
                if float(value) < 0:
                    invalid.add(column)
            except (TypeError, ValueError):
                invalid.add(column)
    for column in ("Claim From DT", "Claim Through DT", "Clm Effective Date", "Clm Obsolete dt"):
        value = row.get(column)
        if value is not None and pd.isna(pd.to_datetime(value, errors="coerce")):
            invalid.add(column)
    start = pd.to_datetime(row.get("Claim From DT"), errors="coerce")
    end = pd.to_datetime(row.get("Claim Through DT"), errors="coerce")
    if pd.notna(start) and pd.notna(end) and start > end:
        invalid.add("Claim Through DT")
    return invalid


def compare(before, after, key):
    old = {str(r[key]).strip(): r for r in before["rows"]}
    new = {str(r[key]).strip(): r for r in after["rows"]}
    old_columns, new_columns = set(before["columns"]), set(after["columns"])
    added, removed = new.keys() - old.keys(), old.keys() - new.keys()
    changed_fields, changed_values, new_nulls, new_invalid = set(), 0, 0, 0
    field_counts = {}
    details = []
    drilldowns = {name: [] for name in ("changed_columns", "changed_values", "added", "removed",
                  "new_nulls", "new_invalid", "changed_rows", "resolved_nulls", "schema_changes")}
    for identity in sorted(old.keys() | new.keys()):
        a, b = old.get(identity), new.get(identity)
        fields = []
        if a is not None and b is not None:
            for col in before["columns"]:
                if col in new_columns and a.get(col) != b.get(col):
                    fields.append(col)
                    changed_fields.add(col)
                    changed_values += 1
                    field_counts[col] = field_counts.get(col, 0) + 1
                    drilldowns["changed_values"].append({"identity": identity, "column": col, "before": a.get(col), "after": b.get(col)})
                if col in new_columns and a.get(col) is None and b.get(col) is not None:
                    drilldowns["resolved_nulls"].append({"identity": identity, "column": col, "before": None, "after": b.get(col)})
        if b is not None:
            nulls = [c for c in after["columns"] if b.get(c) is None and
                     (a is None or c not in old_columns or a.get(c) is not None)]
            invalid = invalid_cells(b) - (invalid_cells(a) if a is not None else set())
            new_nulls += len(nulls)
            new_invalid += len(invalid)
            for col in nulls:
                drilldowns["new_nulls"].append({"identity": identity, "column": col, "before": a.get(col) if a else None, "after": None})
            for col in sorted(invalid):
                drilldowns["new_invalid"].append({"identity": identity, "column": col, "before": a.get(col) if a else None, "after": b.get(col)})
        details.append({"id": identity, "status": "Added" if a is None else "Removed" if b is None else "Changed" if fields else "Unchanged",
                        "fields": fields, "before": a, "after": b})
        if a is None:
            drilldowns["added"].append({"identity": identity, "column": "Entire row", "before": None, "after": b})
        elif b is None:
            drilldowns["removed"].append({"identity": identity, "column": "Entire row", "before": a, "after": None})
        elif fields:
            drilldowns["changed_rows"].append({"identity": identity, "column": ", ".join(fields), "before": a, "after": b})
    for column in sorted(changed_fields):
        drilldowns["changed_columns"].append({"identity": f"{field_counts[column]} changed values", "column": column, "before": None, "after": None})
    for col in sorted(new_columns - old_columns):
        drilldowns["schema_changes"].append({"identity": "Added column", "column": col, "before": "Not present", "after": "Present"})
    for col in sorted(old_columns - new_columns):
        drilldowns["schema_changes"].append({"identity": "Removed column", "column": col, "before": "Present", "after": "Not present"})
    return {"added": len(added), "removed": len(removed), "changed_rows": sum(d["status"] == "Changed" for d in details),
            "changed_values": changed_values, "changed_columns": len(changed_fields),
            "added_columns": sorted(new_columns - old_columns), "removed_columns": sorted(old_columns - new_columns),
            "new_nulls": new_nulls, "new_invalid": new_invalid,
            "resolved_nulls": len(drilldowns["resolved_nulls"]), "schema_changes": len(drilldowns["schema_changes"]),
            "row_count": len(new), "details": details, "drilldowns": drilldowns}


def history(path, owner, dataset, selected=None, offset=0, changed_only=False, metric="", metric_offset=0):
    actual_owner = DEMO if dataset == DEMO else owner
    with connect(path) as conn:
        schedule = conn.execute("SELECT key_column,next_run,error FROM schedules WHERE owner=? AND dataset=?",
                                (actual_owner, dataset)).fetchone()
        rows = conn.execute("SELECT id,captured_at,key_column,data FROM snapshots WHERE owner=? AND dataset=? ORDER BY captured_at,id",
                            (actual_owner, dataset)).fetchall()
    snapshots = [{"id": r["id"], "captured_at": r["captured_at"], "key": r["key_column"],
                  "data": json.loads(r["data"])} for r in rows]
    timeline = []
    for i, snap in enumerate(snapshots):
        summary = compare(snapshots[i - 1]["data"], snap["data"], snap["key"]) if i else None
        if summary:
            summary = {k: v for k, v in summary.items() if k not in ("details", "drilldowns")}
        timeline.append({"id": snap["id"], "captured_at": snap["captured_at"], "summary": summary})
    if selected is not None and not any(s["id"] == selected for s in snapshots):
        raise ValueError("Snapshot not found for this dataset.")
    index = next((i for i, s in enumerate(snapshots) if s["id"] == selected), len(snapshots) - 1)
    current = snapshots[index] if index >= 0 else None
    previous = snapshots[index - 1] if index > 0 else None
    comparison = compare(previous["data"], current["data"], current["key"]) if previous else None
    metric_data = None
    if comparison:
        drilldowns = comparison.pop("drilldowns")
        if metric:
            if metric not in drilldowns:
                raise ValueError("Unknown summary metric.")
            entries = drilldowns[metric]
            metric_offset = max(0, metric_offset)
            metric_data = {"metric": metric, "total": len(entries), "offset": metric_offset,
                           "rows": entries[metric_offset:metric_offset + 50]}
        details = comparison["details"]
        if changed_only:
            details = [row for row in details if row["status"] != "Unchanged"]
        comparison["details_total"] = len(details)
        comparison["offset"] = max(0, offset)
        comparison["details"] = details[comparison["offset"]:comparison["offset"] + 50]
    return {"demo": dataset == DEMO, "schedule": dict(schedule) if schedule else None,
            "timeline": timeline, "selected": current["id"] if current else None,
            "identity_column": current["key"] if current else None,
            "before_date": previous["captured_at"] if previous else None,
            "after_date": current["captured_at"] if current else None,
            "comparison": comparison,
            "metric_data": metric_data,
            "columns": list(dict.fromkeys((previous["data"]["columns"] if previous else []) +
                                           (current["data"]["columns"] if current else []))),
            "baseline_rows": current["data"]["rows"][:100] if current and not previous else []}


def recipients(path, owner, dataset):
    with connect(path) as conn:
        addresses = [row["email"] for row in conn.execute(
            "SELECT email FROM recipients WHERE owner=? AND dataset=? ORDER BY email", (owner, dataset))]
        latest = conn.execute("""SELECT status, last_error, sent_at, email FROM email_outbox
                                WHERE owner=? AND dataset=? ORDER BY id DESC LIMIT 1""", (owner, dataset)).fetchone()
        pending = conn.execute("""SELECT COUNT(*) FROM email_outbox WHERE owner=? AND dataset=?
                                  AND status IN ('pending','sending')""", (owner, dataset)).fetchone()[0]
        failed = conn.execute("""SELECT email, last_error FROM email_outbox WHERE owner=? AND dataset=?
                                 AND status IN ('pending','sending') AND last_error!=''
                                 ORDER BY id DESC LIMIT 1""", (owner, dataset)).fetchone()
    return {"addresses": addresses, "latest": dict(latest) if latest else None,
            "pending": pending, "failure": dict(failed) if failed else None}


def add_recipient(path, owner, dataset, email):
    with connect(path) as conn:
        conn.execute("INSERT OR IGNORE INTO recipients(owner,dataset,email) VALUES(?,?,?)",
                     (owner, dataset, email))


def remove_recipient(path, owner, dataset, email):
    with connect(path) as conn:
        conn.execute("DELETE FROM recipients WHERE owner=? AND dataset=? AND email=?",
                     (owner, dataset, email))
        conn.execute("DELETE FROM email_outbox WHERE owner=? AND dataset=? AND email=? AND status IN ('pending','sending')",
                     (owner, dataset, email))


def deliver_pending(path, sender, now):
    if sender is None:
        return
    with connect(path) as conn:
        ids = [row["id"] for row in conn.execute(
            "SELECT id FROM email_outbox WHERE status IN ('pending','sending') AND next_attempt<=? ORDER BY id",
            (now.isoformat(),))]
    for item_id in ids:
        with connect(path) as conn:
            conn.execute("BEGIN IMMEDIATE")
            item = conn.execute("SELECT * FROM email_outbox WHERE id=? AND status IN ('pending','sending') AND next_attempt<=?",
                                (item_id, now.isoformat())).fetchone()
            if not item:
                continue
            conn.execute("UPDATE email_outbox SET status='sending',next_attempt=?,attempts=attempts+1 WHERE id=?",
                         ((now + timedelta(minutes=10)).isoformat(), item_id))
        try:
            with connect(path) as conn:
                current = conn.execute("SELECT * FROM snapshots WHERE id=? AND owner=? AND dataset=?",
                                       (item["snapshot_id"], item["owner"], item["dataset"])).fetchone()
                previous = conn.execute("""SELECT * FROM snapshots WHERE owner=? AND dataset=? AND id<?
                                           ORDER BY id DESC LIMIT 1""",
                                        (item["owner"], item["dataset"], item["snapshot_id"])).fetchone()
            if not current or not previous:
                raise ValueError("The snapshots needed for this email no longer exist.")
            summary = compare(json.loads(previous["data"]), json.loads(current["data"]), current["key_column"])
            sender(item["email"], item["dataset"], previous["captured_at"], current["captured_at"], summary)
            with connect(path) as conn:
                conn.execute("UPDATE email_outbox SET status='sent',sent_at=?,last_error='' WHERE id=?",
                             (datetime.now(timezone.utc).isoformat(), item_id))
        except Exception as exc:
            traceback.print_exc()
            with connect(path) as conn:
                conn.execute("UPDATE email_outbox SET status='pending',next_attempt=?,last_error=? WHERE id=?",
                             ((now + timedelta(minutes=15)).isoformat(), str(exc), item_id))


def run_due(path, loader, now=None, sender=None):
    now = now or datetime.now(timezone.utc)
    with connect(path) as conn:
        due = conn.execute("SELECT * FROM schedules WHERE next_run<=?", (now.isoformat(),)).fetchall()
    for item in due:
        try:
            # Lock before reading, so concurrent workers cannot capture the same scheduled run.
            with connect(path) as conn:
                conn.execute("BEGIN IMMEDIATE")
                current = conn.execute("SELECT * FROM schedules WHERE owner=? AND dataset=?",
                                       (item["owner"], item["dataset"])).fetchone()
                if not current or current["next_run"] > now.isoformat():
                    continue
                data = frame_data(loader(current["dataset"]), current["key_column"])
                snapshot_id = conn.execute("INSERT INTO snapshots(owner,dataset,captured_at,key_column,data) VALUES(?,?,?,?,?)",
                             (current["owner"], current["dataset"], now.isoformat(), current["key_column"], json.dumps(data))).lastrowid
                conn.execute("""INSERT INTO email_outbox(owner,dataset,snapshot_id,email,next_attempt)
                                SELECT owner,dataset,?,email,? FROM recipients WHERE owner=? AND dataset=?""",
                             (snapshot_id, now.isoformat(), current["owner"], current["dataset"]))
                next_run = datetime.fromisoformat(current["next_run"])
                while next_run <= now:
                    next_run += timedelta(days=7)
                conn.execute("UPDATE schedules SET next_run=?,error='' WHERE owner=? AND dataset=?",
                             (next_run.isoformat(), current["owner"], current["dataset"]))
        except Exception as exc:
            traceback.print_exc()
            with connect(path) as conn:
                conn.execute("UPDATE schedules SET error=? WHERE owner=? AND dataset=?",
                             (str(exc), item["owner"], item["dataset"]))
    deliver_pending(path, sender, now)


def start_worker(path, loader, sender=None):
    global _worker
    with _worker_lock:
        if _worker and _worker.is_alive():
            return
        def loop():
            while True:
                try:
                    run_due(path, loader, sender=sender)
                except Exception:
                    traceback.print_exc()
                threading.Event().wait(60)
        _worker = threading.Thread(target=loop, name="dart-weekly-history", daemon=True)
        _worker.start()
