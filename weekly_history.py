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
        if not conn.execute("SELECT 1 FROM snapshots WHERE owner=? AND dataset=?", (DEMO_TRENDS, DEMO_TRENDS)).fetchone():
            seed_trends_demo(conn)


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


DEMO_TRENDS = "__demo_trends__"
DEMO_DATASETS = (DEMO, DEMO_TRENDS)


def seed_trends_demo(conn):
    """Ten synthetic weeks with planted anomalies so trends and flags have something to show."""
    import random
    rng = random.Random(20260105)
    columns = ["Claim ID", "Bene ID", "Claim From DT", "Provider ID", "Diagnosis Code",
               "Claim Type", "Claim Pmt Amt", "Bene Pmt Amt"]
    diagnoses = ["I10", "E11.9", "J44.9", "M54.5", "I25.1", "E78.5", "J18.9", "I50.9"]
    next_id = [2001]

    def new_row(payment_scale=1.0):
        number = next_id[0]
        next_id[0] += 1
        return {"Claim ID": f"CLM-{number:04}", "Bene ID": f"BENE-{number:04}",
                "Claim From DT": f"2026-01-{rng.randint(1, 28):02}", "Provider ID": f"PRV-{rng.randint(1, 40):03}",
                "Diagnosis Code": rng.choice(diagnoses), "Claim Type": rng.choice(["Inpatient", "Outpatient"]),
                "Claim Pmt Amt": round(rng.uniform(600, 1800) * payment_scale), "Bene Pmt Amt": rng.choice([40, 40, 40, 55])}

    rows = [new_row() for _ in range(80)]
    for r in rng.sample(rows, 3):
        r["Claim Pmt Amt"] = None
    weeks = []
    for week in range(10):
        if week:
            rows = [dict(r) for r in rows]
            for _ in range(rng.randint(2, 4)):
                rows.append(new_row())
            for r in rows:
                if r["Claim Pmt Amt"] is not None:
                    r["Claim Pmt Amt"] = max(100, r["Claim Pmt Amt"] + rng.randint(-60, 90)) if rng.random() < 0.05 else r["Claim Pmt Amt"]
            if rng.random() < 0.5:
                rng.choice(rows)["Provider ID"] = None
        if week == 5:
            for source in rng.sample(rows[:60], 6):
                copy = dict(source)
                copy["Claim ID"] = f"CLM-{next_id[0]:04}"
                next_id[0] += 1
                rows.append(copy)
        if week == 7:
            for r in rng.sample(rows, 14):
                r["Claim Pmt Amt"] = None
        if week == 9:
            rows = [r for r in rows if rng.random() > 0.3]
            for r in rows:
                if r["Claim Pmt Amt"] is not None:
                    r["Claim Pmt Amt"] = round(r["Claim Pmt Amt"] * 1.45)
        elif week == 8:
            for r in rows:
                if r["Claim Pmt Amt"] is None and rng.random() < 0.7:
                    r["Claim Pmt Amt"] = round(rng.uniform(600, 1800))
        weeks.append([dict(r) for r in rows])
    first = datetime.now(timezone.utc).date() - timedelta(days=63)
    for i, records in enumerate(weeks):
        captured = datetime.combine(first + timedelta(days=7 * i), datetime.min.time(), timezone.utc)
        conn.execute("INSERT INTO snapshots(owner,dataset,captured_at,key_column,data) VALUES(?,?,?,?,?)",
                     (DEMO_TRENDS, DEMO_TRENDS, captured.isoformat(), "Claim ID",
                      json.dumps({"columns": columns, "rows": records})))


def _is_null(value):
    return value is None or (isinstance(value, str) and not value.strip())


def _number(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return float(value)


def snapshot_metrics(data, key):
    rows, columns = data["rows"], data["columns"]
    count = len(rows)
    nulls = {c: sum(_is_null(r.get(c)) for r in rows) for c in columns}
    total_cells = count * len(columns)
    null_total = sum(nulls.values())
    others = [c for c in columns if c != key]
    signatures = [json.dumps([r.get(c) for c in others], default=str) for r in rows]
    duplicates = count - len(set(signatures))
    means = {}
    for c in others:
        present = [r.get(c) for r in rows if not _is_null(r.get(c))]
        numeric = [_number(v) for v in present]
        if present and all(v is not None for v in numeric):
            means[c] = sum(numeric) / len(numeric)
    return {"rows": count, "nulls": nulls, "null_total": null_total,
            "null_rate": round(100 * null_total / total_cells, 2) if total_cells else 0.0,
            "completeness": round(100 - 100 * null_total / total_cells, 2) if total_cells else 100.0,
            "duplicates": duplicates, "duplicate_rate": round(100 * duplicates / count, 2) if count else 0.0,
            "means": means}


def _zscore(series, i, floor):
    """Score the latest week-over-week change against the earlier changes in the series."""
    deltas = [series[j] - series[j - 1] for j in range(1, i + 1)]
    current, prior = deltas[-1], deltas[:-1]
    if len(prior) < 3:
        return None
    mean = sum(prior) / len(prior)
    spread = (sum((d - mean) ** 2 for d in prior) / len(prior)) ** 0.5
    return (current - mean) / max(spread, floor)


def _times(ratio):
    text = f"{ratio:.1f}".rstrip("0").rstrip(".")
    return f"{text}×"


def detect_anomalies(snapshots, metrics, threshold=2.0):
    found = []
    columns = list(dict.fromkeys(c for m in metrics for c in m["nulls"]))
    means_columns = list(dict.fromkeys(c for m in metrics for c in m["means"]))
    for i in range(1, len(snapshots)):
        snap, cur, prev = snapshots[i], metrics[i], metrics[i - 1]

        def add(metric, column, z, message, bad):
            if z is not None and abs(z) >= threshold:
                found.append({"snapshot_id": snap["id"], "date": snap["captured_at"], "metric": metric,
                              "column": column, "z": round(z, 1), "severity": "high" if abs(z) >= 3 else "medium",
                              "direction": "bad" if bad else "good", "message": message})

        series = [m["rows"] for m in metrics[:i + 1]]
        z = _zscore(series, i, max(1.0, 0.01 * prev["rows"]))
        if z is not None and prev["rows"]:
            change = 100 * (cur["rows"] - prev["rows"]) / prev["rows"]
            word = "dropped" if change < 0 else "jumped"
            add("rows", None, z, f"Row volume {word} {abs(change):.0f}% this week ({prev['rows']:,} → {cur['rows']:,}).", change < 0)

        for column in columns:
            before, after = prev["nulls"].get(column, 0), cur["nulls"].get(column, 0)
            rates = [100 * m["nulls"].get(column, 0) / m["rows"] if m["rows"] else 0 for m in metrics[:i + 1]]
            z = _zscore(rates, i, 0.75)
            if before == after:
                continue
            if after > before:
                text = (f"{column} nulls jumped {_times(after / before)} this week ({before:,} → {after:,} rows)."
                        if before else f"{column} nulls appeared this week: {after:,} rows are now empty.")
            else:
                text = f"{column} nulls fell from {before:,} to {after:,} rows this week."
            add("nulls", column, z, text, after > before)

        series = [m["duplicate_rate"] for m in metrics[:i + 1]]
        z = _zscore(series, i, 0.5)
        if prev["duplicates"] != cur["duplicates"]:
            add("duplicates", None, z, f"Duplicate rows {'rose' if cur['duplicates'] > prev['duplicates'] else 'fell'} "
                f"from {prev['duplicates']:,} to {cur['duplicates']:,} this week ({cur['duplicate_rate']:.1f}% of rows).",
                cur["duplicates"] > prev["duplicates"])

        series = [m["completeness"] for m in metrics[:i + 1]]
        z = _zscore(series, i, 0.3)
        if prev["completeness"] != cur["completeness"]:
            add("completeness", None, z, f"Completeness {'improved' if cur['completeness'] > prev['completeness'] else 'fell'} "
                f"from {prev['completeness']:.1f}% to {cur['completeness']:.1f}% this week.", cur["completeness"] < prev["completeness"])

        for column in means_columns:
            if column not in prev["means"] or column not in cur["means"] or not prev["means"][column]:
                continue
            values = [m["means"].get(column) for m in metrics[:i + 1]]
            if any(v is None for v in values):
                continue
            z = _zscore(values, i, max(0.02 * abs(prev["means"][column]), 1e-9))
            change = 100 * (cur["means"][column] - prev["means"][column]) / abs(prev["means"][column])
            add("distribution", column, z, f"Average {column} shifted {change:+.0f}% this week "
                f"({prev['means'][column]:,.0f} → {cur['means'][column]:,.0f}).", True)
    return found


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
    actual_owner = dataset if dataset in DEMO_DATASETS else owner
    with connect(path) as conn:
        schedule = conn.execute("SELECT key_column,next_run,error FROM schedules WHERE owner=? AND dataset=?",
                                (actual_owner, dataset)).fetchone()
        rows = conn.execute("SELECT id,captured_at,key_column,data FROM snapshots WHERE owner=? AND dataset=? ORDER BY captured_at,id",
                            (actual_owner, dataset)).fetchall()
    snapshots = [{"id": r["id"], "captured_at": r["captured_at"], "key": r["key_column"],
                  "data": json.loads(r["data"])} for r in rows]
    metrics = [snapshot_metrics(s["data"], s["key"]) for s in snapshots]
    anomalies = detect_anomalies(snapshots, metrics)
    trends = {"weeks": [{"id": s["id"], "captured_at": s["captured_at"], "rows": m["rows"],
                         "completeness": m["completeness"], "duplicate_rate": m["duplicate_rate"],
                         "duplicates": m["duplicates"], "null_rate": m["null_rate"]}
                        for s, m in zip(snapshots, metrics)],
              "anomalies": anomalies}
    timeline = []
    for i, snap in enumerate(snapshots):
        summary = compare(snapshots[i - 1]["data"], snap["data"], snap["key"]) if i else None
        if summary:
            summary = {k: v for k, v in summary.items() if k not in ("details", "drilldowns")}
        timeline.append({"id": snap["id"], "captured_at": snap["captured_at"], "summary": summary,
                         "anomaly_count": sum(a["snapshot_id"] == snap["id"] for a in anomalies)})
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
    return {"demo": dataset in DEMO_DATASETS, "trends": trends, "schedule": dict(schedule) if schedule else None,
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


def key_candidates(frame):
    """Columns whose values are non-empty and unique, so they can identify rows."""
    result = []
    for column in frame.columns:
        values = frame[column].astype(str).str.strip()
        if (frame[column].notna().all() and not (values == "").any() and values.is_unique
                and len(frame) > 0):
            result.append(str(column))
    return result


def capture_now(path, owner, dataset, frame):
    """Manual refresh: snapshot the dataset now without moving the weekly schedule."""
    with connect(path) as conn:
        conn.execute("BEGIN IMMEDIATE")
        last = conn.execute("SELECT key_column FROM snapshots WHERE owner=? AND dataset=? ORDER BY id DESC LIMIT 1",
                            (owner, dataset)).fetchone()
        if not last:
            raise ValueError("Save a weekly schedule first to capture the baseline.")
        data = frame_data(frame, last["key_column"])
        now = datetime.now(timezone.utc).isoformat()
        snapshot_id = conn.execute("INSERT INTO snapshots(owner,dataset,captured_at,key_column,data) VALUES(?,?,?,?,?)",
                                   (owner, dataset, now, last["key_column"], json.dumps(data))).lastrowid
        conn.execute("""INSERT INTO email_outbox(owner,dataset,snapshot_id,email,next_attempt)
                        SELECT owner,dataset,?,email,? FROM recipients WHERE owner=? AND dataset=?""",
                     (snapshot_id, now, owner, dataset))
        return snapshot_id


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
