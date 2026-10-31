"""Tests for the received-date extractor. Standard library only (unittest).

Run from the repo root:   python3 -m unittest discover -s tests -v
Running the tests (re)writes test_report.csv in the repo root with the columns
checks, input, expected, actual, result.

Real manuscripts are optional: if samples/merged_*.docx exist (or SAMPLES_DIR is set)
they are checked against tests/answer_key.csv; otherwise those checks are skipped.
"""
import csv
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tests"))

import extract                      # noqa: E402
from docx_builder import make_docx  # noqa: E402

REPORT = []


def _clean(x):
    return str(x).replace("\r", " ").replace("\n", " / ")


def read_csv_rows(path, as_dict=False):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f)) if as_dict else list(csv.reader(f))


def record(check, inp, expected, actual, result):
    REPORT.append([check, _clean(inp), _clean(expected), _clean(actual), result])


def tearDownModule():
    with open(ROOT / "test_report.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["checks", "input", "expected", "actual", "result"])
        w.writerows(REPORT)


def run_doc(items, header=None, footer=None):
    """Build a docx from items, run the extractor, return (YYYY-MM-DD or 'Invalid', note)."""
    with tempfile.TemporaryDirectory() as d:
        path = Path(d) / "t.docx"
        make_docx(path, items, header, footer)
        try:
            date, note = extract.extract_received(path)
        except Exception as e:                      # same behaviour as the CLI
            return "Invalid", f"unreadable ({type(e).__name__})"
    return (date.isoformat() if date else "Invalid"), note


class Base(unittest.TestCase):
    def check(self, name, inp, expected, actual):
        ok = actual == expected
        record(name, inp, expected, actual, "pass" if ok else "fail")
        self.assertEqual(actual, expected, f"{name}: {inp}")

    def text_cases(self, cases):
        for name, text, expected in cases:
            with self.subTest(check=name, text=text):
                actual, _ = run_doc(["Title of paper", text, "Abstract"])
                self.check(name, text, expected, actual)


class LabelVariations(Base):
    def test_labels(self):
        self.text_cases([
            ("label: colon", "Received: 12 March 2023", "2023-03-12"),
            ("label: no colon", "Received 12 March 2023", "2023-03-12"),
            ("label: Date received", "Date received: 12 March 2023", "2023-03-12"),
            ("label: Manuscript received", "Manuscript received March 12, 2023", "2023-03-12"),
            ("label: 'on'", "Received on 12th March, 2023", "2023-03-12"),
            ("label: upper case", "RECEIVED: 12 March 2023", "2023-03-12"),
            ("label: typo recieved", "Recieved: 12 March 2023", "2023-03-12"),
            ("label: abbreviation Recvd.", "Recvd. 12 Mar 2023", "2023-03-12"),
            ("label: article history prefix", "Article history: Received: 2 Jan 2026; Revised: 4 Mar 2026; Accepted: 10 Apr 2026", "2026-01-02"),
            ("label: one line, semicolons, no colons", "Received 12 Mar 2023; Revised 5 May 2023; Accepted 1 Jun 2023", "2023-03-12"),
            ("label: pipe separated", "Received 2 Jan 2026 | Accepted 10 Apr 2026 | Published online: 15 May 2026", "2026-01-02"),
        ])


class DateFormats(Base):
    def test_formats(self):
        self.text_cases([
            ("format: 2 Jan 2026", "Received: 2 Jan 2026", "2026-01-02"),
            ("format: 2 January 2026", "Received: 2 January 2026", "2026-01-02"),
            ("format: 04 Mar 2026", "Received: 04 Mar 2026", "2026-03-04"),
            ("format: ordinal 2nd January 2026", "Received: 2nd January 2026", "2026-01-02"),
            ("format: month first with comma", "Received: January 2, 2026", "2026-01-02"),
            ("format: dashes 12-Mar-2023", "Received: 12-Mar-2023", "2023-03-12"),
            ("format: ISO", "Received: 2026-01-02", "2026-01-02"),
            ("format: ISO with slashes", "Received: 2023/03/12", "2023-03-12"),
            ("format: dotted 21.12.2026", "Received: 21.12.2026", "2026-12-21"),
            ("format: compact 20230312", "Received: 20230312", "2023-03-12"),
            ("format: Sept. with dot", "Received: 4th Sept. 2026", "2026-09-04"),
            ("format: Sep. 15", "Received: Sep. 15, 2026", "2026-09-15"),
            ("format: two-digit year", "Received: 12/05/26", "2026-05-12"),
            ("format: extra spaces", "Received:    2   Jan   2026", "2026-01-02"),
            ("format: leap day valid", "Received: 29 Feb 2024", "2024-02-29"),
        ])


class Ambiguity(Base):
    def test_numeric_dates(self):
        self.text_cases([
            ("ambiguity: day above 12 is day-first", "Received: 13/04/2023", "2023-04-13"),
            ("ambiguity: second number above 12 is month-first", "Received: 04/13/2023", "2023-04-13"),
            ("ambiguity: same day and month", "Received: 03/03/2026", "2026-03-03"),
            ("ambiguity: no clue defaults to day-first", "Received: 03/04/2026", "2026-04-03"),
        ])

    def test_chronology_picks_month_first(self):
        items = ["Received: 03/04/2026", "Revised: 12 March 2026", "Accepted: 28 March 2026"]
        actual, note = run_doc(items)
        self.check("ambiguity: chronology resolves to month-first", " ; ".join(items), "2026-03-04", actual)

    def test_chronology_picks_day_first(self):
        items = ["Received: 03/04/2026", "Revised: 15 April 2026", "Accepted: 20 June 2026"]
        actual, note = run_doc(items)
        self.check("ambiguity: both fit, day-first", " ; ".join(items), "2026-04-03", actual)

    def test_review_flag(self):
        _, note = run_doc(["Received: 03/04/2026; Revised: 05/06/2026; Accepted: 07/08/2026"])
        self.check("ambiguity: unresolved date is flagged for review",
                   "Received: 03/04/2026; Revised: 05/06/2026; Accepted: 07/08/2026",
                   "note contains NEEDS REVIEW", "note contains NEEDS REVIEW" if "NEEDS REVIEW" in note else note)


class Traps(Base):
    def test_traps(self):
        self.text_cases([
            ("trap: Received in revised form only", "Received in revised form: 4 March 2026", "Invalid"),
            ("trap: original beside revised form", "Received 2 January 2026; Received in revised form 4 March 2026; Accepted 10 April 2026; Available online 15 May 2026", "2026-01-02"),
            ("trap: prose 'received' with a date later", "We received 120 samples on 5 May 2025.", "Invalid"),
            ("trap: prose 'received support'", "The team received additional support on 2 May 2025 from the council.", "Invalid"),
            ("trap: only Accepted", "Accepted: 10 April 2026", "Invalid"),
            ("trap: only Date accepted", "Date accepted: 2026-04-10", "Invalid"),
            ("trap: Published and Issue date only", "Published online: 15 May 2026; Issue date: 1 July 2026", "Invalid"),
            ("trap: Accessed date only", "[9] Guidance note. Accessed: 12 January 2026.", "Invalid"),
            ("trap: Submitted is not Received", "Submitted: 12 November 2025", "Invalid"),
        ])

    def test_decoys_around_real_date(self):
        items = ["We received 120 samples on 5 May 2025.",
                 "An earlier version was accepted for oral presentation on 14 June 2025 and submitted on 12 November 2025.",
                 "Protocol revised 1 February 2025 was used.",
                 "Accessed: 12 January 2026",
                 "Received: 2 Jan 2026; Revised: 4 Mar 2026; Accepted: 10 Apr 2026"]
        actual, _ = run_doc(items)
        self.check("trap: real label amid prose decoys", " ; ".join(items), "2026-01-02", actual)

    def test_two_received_first_wins(self):
        items = ["Received: 2 Jan 2026", "Received: 5 Mar 2026"]
        actual, _ = run_doc(items)
        self.check("assumption: first Received wins", " ; ".join(items), "2026-01-02", actual)


class InvalidDates(Base):
    def test_invalid(self):
        self.text_cases([
            ("invalid: impossible 31/04/2026", "Received: 31/04/2026", "Invalid"),
            ("invalid: 31 February 2023", "Received: 31 February 2023", "Invalid"),
            ("invalid: 29 Feb non-leap year", "Received: 29 February 2026", "Invalid"),
            ("invalid: month 13 ISO", "Received: 2026-13-02", "Invalid"),
            ("invalid: partial Mar 2026", "Received: Mar 2026", "Invalid"),
            ("invalid: label without date", "Received:", "Invalid"),
            ("invalid: no received at all", "Revised: 4 Mar 2026; Accepted: 10 Apr 2026", "Invalid"),
        ])

    def test_future_year_on_other_label_is_ignored(self):
        actual, _ = run_doc(["Received: 12/05/26; Revised: Mar 2026; Accepted: 10 Apr 2031"])
        self.check("invalid: future year on Accepted does not affect Received",
                   "Received: 12/05/26; Revised: Mar 2026; Accepted: 10 Apr 2031", "2026-05-12", actual)


class Layouts(Base):
    def test_layouts(self):
        cases = [
            ("layout: label and date in separate runs",
             ["Title", ("split", "Received: 2 Jan 2026"), "Revised: 4 Mar 2026"], None, None, "2026-01-02"),
            ("layout: label and value in separate table cells",
             ["Title", ("table", [["Accepted", "21 April 2026"], ["Received", "2026-01-02"], ["Revised", "04 Mar 2026"]])], None, None, "2026-01-02"),
            ("layout: label and date on separate lines",
             ["Title", "Received", "2 January 2026"], None, None, "2026-01-02"),
            ("layout: date only in header",
             ["Title", "Body text"], "Received 2 Jan 2026 | Accepted 10 Apr 2026", None, "2026-01-02"),
            ("layout: date only in footer",
             ["Title", "Body text"], None, "Received: 2 Jan 2026", "2026-01-02"),
            ("layout: after references",
             ["Title", "References", "[1] A. Author. 2018.", "Received: 2 January 2026; Revised: 4 March 2026"], None, None, "2026-01-02"),
            ("layout: one date per paragraph",
             ["Title", "Accepted: 10 April 2026", "Received: 2 January 2026", "Revised: 4 March 2026"], None, None, "2026-01-02"),
        ]
        for name, items, header, footer, expected in cases:
            with self.subTest(check=name):
                actual, _ = run_doc(items, header, footer)
                self.check(name, " ; ".join(i if isinstance(i, str) else str(i[0]) for i in items)
                           + (f" | header: {header}" if header else "") + (f" | footer: {footer}" if footer else ""),
                           expected, actual)


class BadFiles(Base):
    def test_corrupt_file(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "bad.docx"
            p.write_bytes(b"this is not a zip file")
            out = Path(d) / "o.csv"
            extract.main(["x", d, str(out)])
            rows = read_csv_rows(out)
        self.check("file: corrupt docx", "bad.docx (not a zip)", "bad.docx,Invalid", ",".join(rows[1]))

    def test_zip_without_document_xml(self):
        import zipfile
        with tempfile.TemporaryDirectory() as d:
            with zipfile.ZipFile(Path(d) / "odd.docx", "w") as z:
                z.writestr("hello.txt", "hi")
            out = Path(d) / "o.csv"
            extract.main(["x", d, str(out)])
            rows = read_csv_rows(out)
        self.check("file: zip without word/document.xml", "odd.docx", "odd.docx,Invalid", ",".join(rows[1]))

    def test_empty_document(self):
        actual, _ = run_doc([])
        self.check("file: empty document", "(no text)", "Invalid", actual)


class CommandLine(Base):
    def _run(self, *args):
        return subprocess.run(["bash", str(ROOT / "run.sh"), *args], capture_output=True, text=True)

    def test_csv_output(self):
        with tempfile.TemporaryDirectory() as d:
            inp = Path(d) / "in"
            inp.mkdir()
            make_docx(inp / "a.docx", ["Received: 12 March 2023"])
            (inp / "b_corrupt.docx").write_bytes(b"junk")
            make_docx(inp / "c.DOCX", ["Revised: 1 May 2023"])
            make_docx(inp / "~$lock.docx", ["Received: 1 Jan 2020"])
            (inp / "notes.txt").write_text("Received: 1 Jan 2020")
            out = Path(d) / "out.csv"
            r = self._run(str(inp), str(out))
            text = out.read_text(encoding="utf-8")
        self.check("cli: exit code 0", "./run.sh in out.csv", "0", str(r.returncode))
        self.check("cli: header, one row per .docx, sorted, ignores other files",
                   "a.docx, b_corrupt.docx, c.DOCX, ~$lock.docx, notes.txt",
                   "file,received_date / a.docx,2023-03-12 / b_corrupt.docx,Invalid / c.DOCX,Invalid",
                   text.strip().replace("\n", " / "))

    def test_empty_folder(self):
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / "in").mkdir()
            out = Path(d) / "out.csv"
            r = self._run(str(Path(d) / "in"), str(out))
            text = out.read_text(encoding="utf-8")
        self.check("cli: empty folder gives header only", "(no files)", "file,received_date", text.strip())

    def test_bad_usage(self):
        r = self._run()
        self.check("cli: no arguments exits with 2", "./run.sh", "2", str(r.returncode))

    def test_missing_folder(self):
        with tempfile.TemporaryDirectory() as d:
            r = self._run(str(Path(d) / "nope"), str(Path(d) / "o.csv"))
        self.check("cli: missing input folder exits with 1", "./run.sh nope out.csv", "1", str(r.returncode))


class RealManuscripts(Base):
    def test_samples_match_answer_key(self):
        samples = Path(os.environ.get("SAMPLES_DIR", ROOT / "samples"))
        files = sorted(samples.glob("merged_*.docx")) if samples.is_dir() else []
        if not files:
            record("real files: samples present", str(samples), "merged_*.docx files", "not found", "skip")
            self.skipTest("no samples folder")
        key = {r["file"]: r["expected"]
               for r in read_csv_rows(ROOT / "tests" / "answer_key.csv", as_dict=True)}
        with tempfile.TemporaryDirectory() as d:
            out = Path(d) / "out.csv"
            extract.main(["x", str(samples), str(out)])
            got = {r["file"]: r["received_date"] for r in read_csv_rows(out, as_dict=True)}
        for name, expected in sorted(key.items()):
            with self.subTest(file=name):
                self.check(f"real file: {name}", name, expected, got.get(name, "(missing)"))


if __name__ == "__main__":
    unittest.main()