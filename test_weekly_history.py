import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd

import weekly_history as history


class WeeklyHistoryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "history.db"
        history.initialize(self.path)

    def test_demo_has_four_weekly_snapshots_with_real_diffs(self):
        data = history.history(self.path, "user", history.DEMO)
        self.assertEqual(len(data["timeline"]), 4)
        self.assertEqual([x["summary"]["changed_values"] if x["summary"] else None
                          for x in data["timeline"]], [None, 1, 2, 3])
        self.assertEqual(data["comparison"]["added"], 1)
        self.assertEqual(data["comparison"]["removed"], 1)
        self.assertEqual(data["comparison"]["new_invalid"], 0)
        second_change = history.history(self.path, "user", history.DEMO, data["timeline"][2]["id"])
        self.assertEqual(second_change["comparison"]["new_nulls"], 2)
        self.assertEqual(second_change["comparison"]["new_invalid"], 1)
        dates = [datetime.fromisoformat(x["captured_at"]) for x in data["timeline"]]
        self.assertTrue(all(b - a == timedelta(days=7) for a, b in zip(dates, dates[1:])))
        history.initialize(self.path)
        self.assertEqual(len(history.history(self.path, "other", history.DEMO)["timeline"]), 4)

    def test_schedule_captures_real_data_and_skips_missed_weeks(self):
        first = pd.DataFrame([{"Claim ID": "A", "Claim Pmt Amt": 10}, {"Claim ID": "B", "Claim Pmt Amt": 20}])
        start = datetime.now(timezone.utc).date()
        history.save_schedule(self.path, "alice", "claims.csv", "Claim ID", start.isoformat(), first)
        later = pd.DataFrame([{"Claim ID": "A", "Claim Pmt Amt": -10}, {"Claim ID": "C", "Claim Pmt Amt": 30}])
        now = datetime.combine(start, datetime.min.time(), timezone.utc) + timedelta(days=15)
        history.run_due(self.path, lambda _: later, now)
        data = history.history(self.path, "alice", "claims.csv")
        self.assertEqual(len(data["timeline"]), 2)
        self.assertEqual(data["comparison"]["changed_values"], 1)
        self.assertEqual(data["comparison"]["new_invalid"], 1)
        self.assertEqual((data["comparison"]["added"], data["comparison"]["removed"]), (1, 1))
        self.assertEqual(datetime.fromisoformat(data["schedule"]["next_run"]).date(), start + timedelta(days=21))
        history.run_due(self.path, lambda _: later, now)
        self.assertEqual(len(history.history(self.path, "alice", "claims.csv")["timeline"]), 2)
        self.assertFalse(history.history(self.path, "bob", "claims.csv")["timeline"])

    def test_invalid_key_failure_and_retry_are_visible(self):
        frame = pd.DataFrame([{"id": 1, "value": "a"}])
        today = datetime.now(timezone.utc).date()
        history.save_schedule(self.path, "alice", "file.csv", "id", today.isoformat(), frame)
        now = datetime.combine(today, datetime.min.time(), timezone.utc) + timedelta(hours=1)
        duplicates = pd.DataFrame([{"id": 1}, {"id": 1}])
        history.run_due(self.path, lambda _: duplicates, now)
        failed = history.history(self.path, "alice", "file.csv")
        self.assertIn("unique", failed["schedule"]["error"])
        self.assertEqual(len(failed["timeline"]), 1)
        history.run_due(self.path, lambda _: frame, now)
        recovered = history.history(self.path, "alice", "file.csv")
        self.assertEqual(recovered["schedule"]["error"], "")
        self.assertEqual(len(recovered["timeline"]), 2)
        with self.assertRaisesRegex(ValueError, "identity column cannot change"):
            history.save_schedule(self.path, "alice", "file.csv", "value", today.isoformat(), frame)

    def test_pagination_and_schema_changes(self):
        before = {"columns": ["id", "old"], "rows": [{"id": i, "old": "a"} for i in range(120)]}
        after = {"columns": ["id", "new"], "rows": [{"id": i, "new": None} for i in range(120)]}
        result = history.compare(before, after, "id")
        self.assertEqual(result["added_columns"], ["new"])
        self.assertEqual(result["removed_columns"], ["old"])
        self.assertEqual(result["new_nulls"], 120)
        self.assertEqual(result["changed_values"], 0)
        now = datetime.now(timezone.utc)
        history.save_schedule(self.path, "a", "file.csv", "id", now.date().isoformat(), pd.DataFrame(before["rows"]))
        history.run_due(self.path, lambda _: pd.DataFrame(after["rows"]), now + timedelta(days=1))
        page = history.history(self.path, "a", "file.csv", offset=100)
        self.assertEqual(page["comparison"]["details_total"], 120)
        self.assertEqual(len(page["comparison"]["details"]), 20)
        self.assertEqual(len(history.history(self.path, "a", "file.csv", changed_only=True)["comparison"]["details"]), 0)

    def test_summary_drilldowns_match_counts(self):
        data = history.history(self.path, "user", history.DEMO)
        snapshot = data["timeline"][2]["id"]
        expected = {"changed_columns": 2, "changed_values": 2, "changed_rows": 2,
                    "added": 1, "removed": 0, "new_nulls": 2, "new_invalid": 1,
                    "resolved_nulls": 0, "schema_changes": 0}
        for metric, count in expected.items():
            result = history.history(self.path, "user", history.DEMO, snapshot, metric=metric)
            self.assertEqual(result["comparison"][metric], count)
            self.assertEqual(result["metric_data"]["total"], count)
            self.assertEqual(len(result["metric_data"]["rows"]), count)
        invalid = history.history(self.path, "user", history.DEMO, snapshot, metric="new_invalid")
        self.assertEqual(invalid["metric_data"]["rows"][0]["column"], "Bene Pmt Amt")
        self.assertEqual(invalid["metric_data"]["rows"][0]["after"], -25)
        with self.assertRaisesRegex(ValueError, "Unknown summary metric"):
            history.history(self.path, "user", history.DEMO, snapshot, metric="missing")

    def test_emails_queue_on_refresh_retry_and_are_scoped(self):
        today = datetime.now(timezone.utc).date()
        now = datetime.combine(today, datetime.min.time(), timezone.utc) + timedelta(hours=1)
        frame = pd.DataFrame([{"id": "1", "value": "before"}])
        history.save_schedule(self.path, "alice", "file.csv", "id", today.isoformat(), frame)
        history.add_recipient(self.path, "alice", "file.csv", "one@domain.test")
        history.add_recipient(self.path, "alice", "file.csv", "ONE@domain.test")
        history.add_recipient(self.path, "alice", "file.csv", "two@domain.test")
        history.add_recipient(self.path, "bob", "file.csv", "other@domain.test")
        history.add_recipient(self.path, "alice", "other.csv", "other@domain.test")
        self.assertEqual(len(history.recipients(self.path, "alice", "file.csv")["addresses"]), 2)
        changed = pd.DataFrame([{"id": "1", "value": "after"}, {"id": "2", "value": None}])
        calls = []
        def failing(email, dataset, before, after, summary):
            calls.append(email)
            raise RuntimeError("SMTP unavailable")
        history.run_due(self.path, lambda _: changed, now, failing)
        status = history.recipients(self.path, "alice", "file.csv")
        self.assertEqual(status["pending"], 2)
        self.assertIn("SMTP unavailable", status["failure"]["last_error"])
        self.assertEqual(len(history.history(self.path, "alice", "file.csv")["timeline"]), 2)
        history.remove_recipient(self.path, "alice", "file.csv", "two@domain.test")
        delivered = []
        def sender(email, dataset, before, after, summary):
            delivered.append((email, dataset, summary["changed_values"], summary["added"]))
        history.run_due(self.path, lambda _: changed, now + timedelta(minutes=16), sender)
        self.assertEqual(delivered, [("one@domain.test", "file.csv", 1, 1)])
        self.assertEqual(history.recipients(self.path, "alice", "file.csv")["pending"], 0)
        history.run_due(self.path, lambda _: changed, now + timedelta(minutes=17), sender)
        self.assertEqual(len(delivered), 1)
        self.assertEqual(history.recipients(self.path, "bob", "file.csv")["pending"], 0)


if __name__ == "__main__":
    unittest.main()
