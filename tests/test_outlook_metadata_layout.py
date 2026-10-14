"""Real-world GetArray layouts, including square batches reported in issue #53."""

from datetime import datetime, timedelta, timezone
from types import SimpleNamespace
import unittest

from outlook_rag import outlook
from verify_outlook_metadata import probe


class MetadataTable:
    def __init__(self, rows, layout):
        self.rows = rows
        self.layout = layout
        self.position = 0
        self.calls = 0
        self.columns = []
        self.Columns = SimpleNamespace(RemoveAll=self.columns.clear, Add=self.columns.append)

    def Sort(self, *_):
        pass

    @property
    def EndOfTable(self):
        return self.position >= len(self.rows)

    def GetArray(self, count):
        self.calls += 1
        rows = self.rows[self.position:self.position + count]
        self.position += len(rows)
        return tuple(zip(*rows)) if self.layout == "columns" else tuple(rows)

    def MoveToStart(self):
        self.position = 0

    def GetNextRow(self):
        values = dict(zip(self.columns, self.rows[self.position]))
        self.position += 1
        return SimpleNamespace(Item=values.__getitem__)


def source_rows(count):
    now = datetime(2026, 10, 14, 0, 30, tzinfo=timezone.utc)
    return [
        (f"entry-{i}", "IPM.Note.Custom" if i % 2 else "IPM.Appointment",
         now - timedelta(minutes=i), now + timedelta(minutes=i))
        for i in range(count)
    ]


def metadata(table):
    return outlook.mail_metadata(SimpleNamespace(GetTable=lambda _: table), [])


class OutlookMetadataLayout(unittest.TestCase):
    def test_both_layouts_preserve_every_field_and_batch_boundary(self):
        for layout in ("rows", "columns"):
            for count in (0, 1, 2, 3, 4, 63, 64, 65, 68, 128, 129):
                with self.subTest(layout=layout, count=count):
                    source = source_rows(count)
                    table = MetadataTable(source, layout)
                    actual = list(metadata(table))
                    self.assertEqual(len(actual), count)
                    self.assertEqual(table.calls, (count + 63) // 64)
                    for (row, kind), expected in zip(actual, source):
                        self.assertEqual(kind, "GetTable")
                        self.assertEqual(row, dict(
                            entry_id=expected[0], cursor_id=expected[0],
                            mail=expected[1].startswith("IPM.Note"),
                            received=outlook.iso(expected[2], utc=True),
                            modified=outlook.iso(expected[3], utc=True), item=None,
                        ))

    def test_square_string_dates_are_not_mistaken_for_ids(self):
        rows = [(entry, kind, received.isoformat(), modified.isoformat())
                for entry, kind, received, modified in source_rows(4)]
        for layout in ("rows", "columns"):
            with self.subTest(layout=layout):
                actual = list(metadata(MetadataTable(rows, layout)))
                self.assertEqual([row["entry_id"] for row, _ in actual], [r[0] for r in rows])
                self.assertEqual([row["received"] for row, _ in actual], [r[2] for r in rows])
                self.assertEqual([row["modified"] for row, _ in actual], [r[3] for r in rows])

    def test_non_mail_and_missing_dates_preserve_their_slots(self):
        rows = [("a", "IPM.Note", None, ""), ("b", "REPORT.IPM.Note.NDR", "", None),
                ("c", "IPM.Appointment", None, None), ("d", "IPM.Note.Custom", "", "")]
        for layout in ("rows", "columns"):
            with self.subTest(layout=layout):
                actual = list(metadata(MetadataTable(rows, layout)))
                self.assertEqual([r["entry_id"] for r, _ in actual], list("abcd"))
                self.assertEqual([r["mail"] for r, _ in actual], [True, False, False, True])
                self.assertTrue(all(r["received"] == r["modified"] == "" for r, _ in actual))

    def test_invalid_or_ambiguous_batch_yields_nothing_and_hides_values(self):
        for values in (
            "private-folder-name",
            ("private-mail-id", "IPM.Note", "2026-10-14", "2026-10-14"),
            (("a", "IPM.Note", "2026-10-14"),),
            (("a", "IPM.Note", "2026-10-14", "2026-10-14"), ("b",)),
            (("a", "IPM.Note", "2026-10-14", "2026-10-14"),
             ("private-mail-id", "IPM.Note", "not-a-date", "2026-10-14")),
            (("2026-10-14",) * 4,) * 4,
        ):
            with self.subTest(shape=type(values).__name__):
                table = MetadataTable([], "rows")
                table.rows = [None]
                table.GetArray = lambda _: values
                with self.assertRaises(ValueError) as raised:
                    next(metadata(table))
                self.assertNotIn("private-", str(raised.exception))

    def test_read_only_probe_handles_both_layouts_and_square_batches(self):
        for layout in ("rows", "columns"):
            for count in (4, 65, 68):
                with self.subTest(layout=layout, count=count):
                    table = MetadataTable(source_rows(count), layout)
                    result = probe(SimpleNamespace(GetTable=lambda _: table), max_rows=count)
                    self.assertTrue(result["passed"])
                    self.assertEqual(result["rows_checked"], count)
                    self.assertFalse(result["embedding_api_used"])
                    self.assertFalse(result["mail_bodies_read"])


if __name__ == "__main__":
    unittest.main()
