"""Tests for M2: xlsx reading/writing, parsing rules, input template and loader."""

from __future__ import annotations

import contextlib
import datetime as dt
import io
import os
import sys
import tempfile
import unittest
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from statusreport import labels, parse  # noqa: E402
from statusreport.__main__ import main  # noqa: E402
from statusreport.loader import InputError, load  # noqa: E402
from statusreport.pptx.lint import lint_package  # noqa: E402
from statusreport.pptx.opc import Package  # noqa: E402
from statusreport.template import template_bytes  # noqa: E402
from statusreport.xlsx_read import Percent, Workbook, XlsxError, is_date_format  # noqa: E402
from statusreport.xlsx_write import Sheet, Styled, workbook_bytes  # noqa: E402

FX = os.path.join(ROOT, "tests", "fixtures")
D = dt.date


def fixture(name: str) -> str:
    return os.path.join(FX, name)


class TempFiles(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = self._tmp.name

    def tearDown(self):
        self._tmp.cleanup()

    def write(self, name: str, data: bytes) -> str:
        path = os.path.join(self.tmp, name)
        with open(path, "wb") as fh:
            fh.write(data)
        return path

    def workbook(self, tasks_rows, settings_rows=None, extra=None) -> str:
        header = [f.header for f in labels.FIELDS]
        sheets = [Sheet(labels.SHEET_TASKS, [header] + tasks_rows)]
        if settings_rows is not None:
            sheets.append(Sheet(labels.SHEET_SETTINGS, [list(labels.SETTINGS_HEADERS)] + settings_rows))
        sheets += extra or []
        return self.write("input.xlsx", workbook_bytes(sheets))


def task_row(**kw):
    return [kw.get(f.key) for f in labels.FIELDS]


# ---- parsing (R-IN-6, R-IN-8) ------------------------------------------------------------------


class ParseTest(unittest.TestCase):
    def test_dates(self):
        for raw in (D(2026, 11, 2), dt.datetime(2026, 11, 2, 14, 30), "2026-11-02", "02.11.2026",
                    "2.11.2026", "02/11/2026", "2026-11-02 08:00", 46328.0):
            self.assertEqual(parse.date(raw), D(2026, 11, 2), raw)
        self.assertIsNone(parse.date(None))
        self.assertIsNone(parse.date("  "))
        for bad in ("31.02.2026", "next week", 5.0, "11/2026"):
            with self.assertRaises(ValueError, msg=bad):
                parse.date(bad)

    def test_hours(self):
        cases = {
            "1w 2d 4h 30m": 60.5, "3d": 24.0, "12h": 12.0, "12 h": 12.0, "4h 30m": 4.5,
            "1,5d": 12.0, "1.5 d": 12.0, 7: 7.0, "4,5": 4.5, "0h": 0.0,
        }
        for raw, expected in cases.items():
            self.assertAlmostEqual(parse.hours(raw), expected, msg=raw)
        self.assertEqual(parse.hours("2d", hours_per_day=7.5), 15.0)
        self.assertIsNone(parse.hours(""))
        for bad in ("two days", "3x", "-4", "1d and 2h"):
            with self.assertRaises(ValueError, msg=bad):
                parse.hours(bad)

    def test_percent(self):
        self.assertEqual(parse.percent(65), 65)
        self.assertEqual(parse.percent("65 %"), 65)
        self.assertEqual(parse.percent("65,5%"), 65.5)
        self.assertEqual(parse.percent(Percent(0.65)), 65)
        self.assertEqual(parse.percent(Percent(1.0)), 100)
        self.assertEqual(parse.percent(0.5), 0.5)  # plain numbers are percentage points
        for bad in (101, -1, "half", Percent(1.2)):
            with self.assertRaises(ValueError, msg=bad):
                parse.percent(bad)

    def test_light(self):
        for raw, expected in (("Green", "green"), (" RED ", "red"), ("o", "orange"), ("Amber", "orange"), ("yellow", "orange")):
            self.assertEqual(parse.light(raw), expected)
        self.assertIsNone(parse.light(None))
        with self.assertRaises(ValueError):
            parse.light("blue")

    def test_text_and_lists(self):
        self.assertEqual(parse.text(123.0), "123")
        self.assertEqual(parse.text(1.25), "1.25")
        self.assertEqual(parse.text("  x "), "x")
        self.assertEqual(parse.word_list("done, Closed;resolved,, "), ["done", "Closed", "resolved"])

    def test_natural_sort(self):
        versions = ["1.10.0", "1.2.0", "1.9.1", "2.0.0", "1.2.0-rc1"]
        self.assertEqual(sorted(versions, key=parse.natural_key), ["1.2.0", "1.2.0-rc1", "1.9.1", "1.10.0", "2.0.0"])


# ---- xlsx reader / writer (T-2.1, T-2.2) ---------------------------------------------------------


class XlsxTest(unittest.TestCase):
    def test_date_format_detection(self):
        for code, expected in (
            ("yyyy-mm-dd", True), ("dd.mm.yyyy", True), ("[$-407]dd/mm/yyyy", True), ("h:mm", True),
            ("General", False), ("0.00", False), ('0.00" days"', False), ("[Red]0.00", False), ("#,##0", False),
        ):
            self.assertEqual(is_date_format(code), expected, code)

    def test_roundtrip_with_gaps_and_types(self):
        data = workbook_bytes([Sheet("S", [["a", None, "c"], [], [1.5, True, D(2026, 11, 10)], [None, Styled("x", "bold")]])])
        rows = Workbook(data).rows("S")
        self.assertEqual(rows, [["a", None, "c"], [], [1.5, True, D(2026, 11, 10)], [None, "x"]])

    def test_writer_output_is_a_clean_package(self):
        self.assertEqual(lint_package(Package.from_bytes(template_bytes())), [])

    def test_excel_saved_file(self):
        wb = Workbook.open(fixture("excel_filled.xlsx"))
        self.assertEqual(wb.sheet_names, ["Read me", "Tasks", "Import", "Mapping", "Settings"])
        rows = wb.rows("Tasks")
        self.assertEqual(rows[0][:3], ["ID", "Title", "Version"])
        self.assertEqual(rows[2][4], D(2026, 11, 2))  # real date cell
        self.assertEqual(rows[3][4], "01.10.2026")  # date typed as text
        self.assertIsInstance(rows[4][8], Percent)
        self.assertEqual(rows[4][2], "1.10.0")  # text column keeps the version

    def test_1904_date_system(self):
        wb = Workbook.open(fixture("excel_1904.xlsx"))
        self.assertTrue(wb.date1904)
        self.assertEqual(wb.rows("Dates")[1][0], D(2026, 11, 10))

    def test_rejects_xls_and_non_excel_files(self):
        with self.assertRaisesRegex(XlsxError, "old .xls"):
            Workbook.open(fixture("polarion_export.xls"))
        with self.assertRaisesRegex(XlsxError, "not an Excel"):
            Workbook(b"just text", "notes.xlsx")
        with self.assertRaisesRegex(XlsxError, "cannot open"):
            Workbook.open(fixture("does_not_exist.xlsx"))


# ---- template (R-IN-1 ... R-IN-5) -----------------------------------------------------------------------


class TemplateTest(TempFiles):
    def test_template_structure(self):
        wb = Workbook(template_bytes())
        self.assertEqual(wb.sheet_names, ["Read me", "Tasks", "Import", "Mapping", "Settings"])
        self.assertEqual(wb.rows("Tasks")[0], [f.header for f in labels.FIELDS])
        self.assertTrue(wb.rows("Tasks")[1][0].startswith("#"))
        self.assertEqual(wb.rows("Import"), [])
        mapping = {r[0]: (r[1] if len(r) > 1 else None) for r in wb.rows("Mapping")[1:]}
        self.assertEqual(mapping["ID"], "ID")
        self.assertEqual(mapping["Estimate (h)"], "Initial Estimate")
        with zipfile.ZipFile(io.BytesIO(template_bytes())) as zf:
            tasks_xml = zf.read("xl/worksheets/sheet2.xml").decode()
        self.assertIn('"Green,Orange,Red"', tasks_xml)
        self.assertIn('state="frozen"', tasks_xml)

    def test_fresh_template_loads_without_tasks(self):
        path = self.write("t.xlsx", template_bytes())
        result = load(path, today=D(2026, 11, 10))
        self.assertEqual(result.tasks, [])
        self.assertEqual(result.errors, [])
        self.assertEqual(result.settings.project, "My Project")
        self.assertEqual(result.settings.report_date, D(2026, 11, 10))


# ---- loader (R-IN-7, R-OP-3) -----------------------------------------------------------------------------


class LoaderFixtureTest(unittest.TestCase):
    def test_filled_by_excel(self):
        r = load(fixture("excel_filled.xlsx"))
        tasks = {t.id: t for t in r.tasks}
        self.assertEqual(sorted(tasks), ["PRJ-101", "PRJ-102", "PRJ-103", "PRJ-104", "PRJ-106"])
        t = tasks["PRJ-101"]
        self.assertEqual((t.start, t.due, t.estimate_h, t.spent_h), (D(2026, 11, 2), D(2026, 11, 27), 24.0, 10.0))
        self.assertEqual((t.row, t.source), (3, "Tasks"))
        self.assertEqual((tasks["PRJ-102"].start, tasks["PRJ-102"].due), (D(2026, 10, 1), D(2026, 11, 13)))
        self.assertEqual(tasks["PRJ-103"].progress_manual, 90)
        self.assertEqual(tasks["PRJ-103"].version, "1.10.0")
        self.assertEqual(tasks["PRJ-104"].light_override, "red")
        self.assertEqual(tasks["PRJ-106"].state, "Done")
        s = r.settings
        self.assertEqual((s.project, s.customer, s.report_date), ("Falcon", "ACME Automotive", D(2026, 11, 10)))
        self.assertEqual(s.version_order, ["1.2.0", "1.10.0"])
        messages = [str(i) for i in r.issues]
        self.assertEqual(len(r.errors), 4, messages)
        self.assertTrue(any("row 7" in m and "31.02.2026" in m for m in messages))
        self.assertTrue(any("row 8" in m and "duplicate ID PRJ-101" in m for m in messages))
        self.assertTrue(any("row 9" in m and "ID missing" in m for m in messages))
        self.assertEqual([i.row for i in r.issues], sorted(i.row for i in r.issues))

    def check_merged(self, r):
        tasks = {t.id: t for t in r.tasks}
        self.assertEqual(sorted(tasks), ["PRJ-101", "PRJ-201", "PRJ-202", "PRJ-203", "PRJ-301"])
        t = tasks["PRJ-101"]
        self.assertEqual(t.source, "Import+Tasks")
        self.assertEqual(t.title, "Login dialog with SSO")  # from the import (Tasks cell empty)
        self.assertEqual(t.comment, "waiting for API")  # from Tasks
        self.assertEqual(t.start, D(2026, 11, 2))  # only in Tasks
        self.assertEqual(t.spent_h, 10.0)  # "1d 2h" from the import
        self.assertEqual(t.state, "in progress")  # Tasks wins over "In Progress"
        self.assertEqual(tasks["PRJ-201"].estimate_h, 56.0)  # 1w 2d
        self.assertEqual(tasks["PRJ-201"].due, D(2026, 10, 30))
        self.assertEqual(tasks["PRJ-202"].estimate_h, 4.5)
        self.assertEqual(tasks["PRJ-202"].owner, "Hana")
        self.assertEqual(tasks["PRJ-202"].spent_h, 0.0)  # empty Time Spent in the export (R-IN-11)
        self.assertEqual(tasks["PRJ-301"].source, "Tasks")
        self.assertIsNone(tasks["PRJ-301"].spent_h)  # Tasks-only row: empty stays unknown
        self.assertEqual(r.errors, [])

    def test_import_sheet(self):
        r = load(fixture("excel_with_import.xlsx"))
        self.assertEqual(r.import_source, "Import sheet")
        self.check_merged(r)

    def test_separate_polarion_file(self):
        r = load(fixture("excel_with_import.xlsx"), polarion=fixture("polarion_export.xlsx"))
        self.assertTrue(r.import_source.endswith("polarion_export.xlsx"))
        self.check_merged(r)


class LoaderUnitTest(TempFiles):
    def test_missing_tasks_sheet(self):
        path = self.write("x.xlsx", workbook_bytes([Sheet("Other", [["a"]])]))
        with self.assertRaisesRegex(InputError, '"Tasks" is missing'):
            load(path)

    def test_missing_id_column(self):
        path = self.write("x.xlsx", workbook_bytes([Sheet("Tasks", [["Title", "Version"], ["a", "1"]])]))
        with self.assertRaisesRegex(InputError, 'column "ID" not found'):
            load(path)

    def test_required_fields_and_comment_rows(self):
        path = self.workbook([
            task_row(id="A-1", title="ok", version="1.0"),
            task_row(id="A-2", version="1.0"),
            task_row(id="#A-3", title="ignored", version="1.0"),
            task_row(id=7.0, title="numeric id", version=2.0),
        ])
        r = load(path)
        self.assertEqual([t.id for t in r.tasks], ["A-1", "7"])
        self.assertEqual(r.tasks[1].version, "2")
        self.assertTrue(any("A-2: Title missing" in str(i) for i in r.errors))

    def test_all_errors_of_a_row_are_reported(self):
        path = self.workbook([task_row(id="A-1", title="t", version="1", due="soon", progress=150, light_override="blue")])
        r = load(path)
        self.assertEqual(r.tasks, [])
        text = "\n".join(str(i) for i in r.errors)
        for part in ("Due:", "Progress (%):", "Light override:", "row skipped"):
            self.assertIn(part, text)

    def test_start_after_due_is_a_warning(self):
        path = self.workbook([task_row(id="A-1", title="t", version="1", start="2026-12-01", due="2026-11-01")])
        r = load(path)
        self.assertEqual(len(r.tasks), 1)
        self.assertTrue(any(i.level == "warning" and "Start is after Due" in i.message for i in r.issues))

    def test_settings_validation(self):
        path = self.workbook(
            [task_row(id="A-1", title="t", version="9.9")],
            [
                ["Orange threshold (pp)", 30], ["Red threshold (pp)", 20],
                ["Rows per slide", 2.5], ["Hours per day", "eight"],
                ["Colour", "blue"], ["Done states", "Done, Accepted"], ["Version order", "1.0"],
            ],
        )
        r = load(path)
        s = r.settings
        self.assertEqual((s.orange_pp, s.red_pp, s.rows_per_slide, s.hours_per_day), (10, 25, 12, 8))
        self.assertEqual(s.done_states, ["done", "accepted"])
        warnings = "\n".join(str(i) for i in r.issues if i.level == "warning")
        for part in ("Orange threshold is above", "Rows per slide", "Hours per day", 'unknown setting "Colour"', '"9.9" is not in "Version order"'):
            self.assertIn(part, warnings)

    def test_import_with_wrong_mapping(self):
        mapping = Sheet(labels.SHEET_MAPPING, [list(labels.MAPPING_HEADERS), ["ID", "Key"], ["Title", "Summary"], ["Version", "Release"]])
        imp = Sheet(labels.SHEET_IMPORT, [["Key", "Summary", "Fix Version"], ["X-1", "a", "1.0"]])
        path = self.workbook([], extra=[imp, mapping])
        with self.assertRaisesRegex(InputError, r'"release" \(Version\) not found.*"fix version"'):
            load(path)

    def test_empty_imported_time_spent_is_zero(self):
        """R-IN-11: only for imported rows with Time Spent mapped; a Tasks value still wins."""
        mapping = Sheet(labels.SHEET_MAPPING, [
            list(labels.MAPPING_HEADERS), ["ID", "Key"], ["Title", "Summary"], ["Version", "Release"], ["Spent (h)", "Time Spent"],
        ])
        imp = Sheet(labels.SHEET_IMPORT, [["Key", "Summary", "Release", "Time Spent"], ["X-1", "a", "1"], ["X-2", "b", "1"], ["X-3", "c", "1", "2h"]])
        path = self.workbook([task_row(id="X-2", spent_h=5), task_row(id="M-1", title="manual", version="1")], extra=[imp, mapping])
        tasks = {t.id: t for t in load(path).tasks}
        self.assertEqual(tasks["X-1"].spent_h, 0.0)
        self.assertEqual(tasks["X-2"].spent_h, 5.0)
        self.assertEqual(tasks["X-3"].spent_h, 2.0)
        self.assertIsNone(tasks["M-1"].spent_h)

    def test_unmapped_time_spent_stays_unknown(self):
        mapping = Sheet(labels.SHEET_MAPPING, [list(labels.MAPPING_HEADERS), ["ID", "Key"], ["Title", "Summary"], ["Version", "Release"]])
        imp = Sheet(labels.SHEET_IMPORT, [["Key", "Summary", "Release"], ["X-1", "a", "1"]])
        (t,) = load(self.workbook([], extra=[imp, mapping])).tasks
        self.assertIsNone(t.spent_h)

    def test_import_header_below_title_rows(self):
        mapping = Sheet(labels.SHEET_MAPPING, [list(labels.MAPPING_HEADERS), ["ID", "Key"], ["Title", "Summary"], ["Version", "Release"]])
        imp = Sheet(labels.SHEET_IMPORT, [["Export of 2026-11-10"], [], ["Key", "Summary", "Release"], ["X-1", "a", "1.0"], ["X-1", "b", "1.0"]])
        r = load(self.workbook([], extra=[imp, mapping]))
        self.assertEqual([(t.id, t.title, t.row) for t in r.tasks], [("X-1", "a", 4)])
        self.assertTrue(any("row 5" in str(i) and "duplicate ID X-1" in str(i) for i in r.errors))


class CliTest(TempFiles):
    def run_cli(self, *args) -> int:
        with contextlib.redirect_stdout(io.StringIO()) as out:
            code = main(list(args))
        self.output = out.getvalue()
        return code

    def test_init_and_check(self):
        path = os.path.join(self.tmp, "new.xlsx")
        self.assertEqual(self.run_cli(path, "--init"), 0)
        self.assertTrue(os.path.exists(path))
        self.assertEqual(self.run_cli(path, "--init"), 2)  # no silent overwrite
        self.assertIn("already exists", self.output)
        self.assertEqual(self.run_cli(path, "--check"), 0)
        self.assertEqual(self.run_cli(fixture("excel_filled.xlsx"), "--check"), 1)
        self.assertIn("4 error(s)", self.output)
        self.assertEqual(self.run_cli(fixture("excel_filled.xlsx"), "--check", "--date", "2026-12-24"), 1)
        self.assertIn("CW 52/2026", self.output)
        self.assertEqual(self.run_cli(fixture("polarion_export.xls"), "--check"), 2)

    def test_input_is_not_modified(self):
        path = fixture("excel_filled.xlsx")
        with open(path, "rb") as fh:
            before = fh.read()
        self.run_cli(path, "--check")
        with open(path, "rb") as fh:
            self.assertEqual(fh.read(), before)  # R-OP-4


if __name__ == "__main__":
    unittest.main()
