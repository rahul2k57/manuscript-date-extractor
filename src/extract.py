#!/usr/bin/env python3
"""Find the "received" date in manuscript .docx files and write it as YYYY-MM-DD.

Usage: python3 src/extract.py INPUT_FOLDER OUTPUT.csv

Standard library only (Python 3.12). A .docx file is a zip of XML parts, so we
read the XML directly instead of depending on python-docx.

Pipeline:  read_lines -> scan (find "Received" + date) -> resolve (to a real date)
Anything we cannot turn into a real date is written as "Invalid".
Review notes (ambiguous dates, reasons for Invalid) go to stderr, never the CSV.
"""
import csv
import datetime as dt
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TEXT_PART = re.compile(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")


# ---------------------------------------------------------------- 1. reader
def _paragraph_text(p):
    """Join every run of one paragraph, so a bold label + plain date still match."""
    out = []
    for el in p.iter():
        if el.tag == W + "t":
            out.append(el.text or "")
        elif el.tag == W + "tab":
            out.append(" ")
        elif el.tag in (W + "br", W + "cr"):
            out.append("\n")
        elif el.tag == W + "noBreakHyphen":
            out.append("-")
    return "".join(out)


def read_lines(path):
    """Return one string per paragraph from body, headers, footers and footnotes.

    Table cells are paragraphs too, so a label cell is followed by its value cell.
    Raises zipfile.BadZipFile / KeyError for corrupt or non-docx files.
    """
    lines = []
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if TEXT_PART.match(n)]
        if "word/document.xml" not in names:
            raise KeyError("word/document.xml missing")
        # body first, then headers/footers/footnotes
        names.sort(key=lambda n: (n != "word/document.xml", n))
        for name in names:
            root = ET.fromstring(z.read(name))
            for p in root.iter(W + "p"):
                text = _paragraph_text(p).replace("\xa0", " ")
                if text.strip():
                    lines.append(text)
    return lines


# ---------------------------------------------------------------- 2. finder
MONTHS = {}
for _i, _name in enumerate(
    ["january", "february", "march", "april", "may", "june", "july",
     "august", "september", "october", "november", "december"], 1):
    for _n in range(3, len(_name) + 1):      # jan, janu, ..., also "sept"
        MONTHS[_name[:_n]] = _i

RE_ISO = re.compile(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?!\d)")
RE_NUM = re.compile(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4}|\d{2})(?!\d)")
RE_DMY = re.compile(r"(\d{1,2})(?:st|nd|rd|th)?(?:\s+of)?[\s\-./,]*([A-Za-z]{3,9})\.?[\s\-./,]*(\d{4})(?!\d)")
RE_MDY = re.compile(r"([A-Za-z]{3,9})\.?[\s\-./,]*(\d{1,2})(?:st|nd|rd|th)?[\s,\-./]*(\d{4})(?!\d)")
RE_MY = re.compile(r"([A-Za-z]{3,9})\.?[\s\-./,]*(\d{4})(?!\d)")
RE_COMPACT = re.compile(r"(\d{4})(\d{2})(\d{2})(?!\d)")
SKIP = re.compile(r"[\s:|\-\u2013\u2014]*(?:on\s+)?", re.I)


def read_date(text, pos):
    """Try to read a date starting exactly at pos. Returns (kind, groups) or None."""
    for kind, rx in (("iso", RE_ISO), ("num", RE_NUM), ("dmy", RE_DMY),
                     ("mdy", RE_MDY), ("partial", RE_MY), ("compact", RE_COMPACT)):
        m = rx.match(text, pos)
        if not m:
            continue
        g = m.groups()
        if kind == "dmy" and g[1].lower() not in MONTHS:
            continue
        if kind == "mdy" and g[0].lower() not in MONTHS:
            continue
        if kind == "partial" and g[0].lower() not in MONTHS:
            continue
        return kind, g
    return None


# "received" (typos allowed), but not "received in revised form"
RECEIVED = re.compile(
    r"(?<![A-Za-z])(?:rec(?:ei|ie)ved|recvd\.?)(?![A-Za-z])(?!\s+in\s+(?:revised|final)\b)", re.I)
# other dates in the file, used only to resolve DD/MM vs MM/DD
CONTEXT = re.compile(
    r"(?<![A-Za-z])(?:received\s+in\s+revised\s+form|revised|accepted|(?:published|available)\s+online)"
    r"(?:\s*\([^)]*\))?(?![A-Za-z])", re.I)


def scan(text, label_re):
    """Every date that sits immediately after a label match.

    'Received: 2 Jan 2026' counts; 'we received 120 samples on 5 May 2025' does not,
    because a date must follow the label directly.
    """
    found = []
    for m in label_re.finditer(text):
        pos = SKIP.match(text, m.end()).end()
        r = read_date(text, pos)
        if r:
            found.append(r)
    return found


# ---------------------------------------------------------------- 3. parser
def to_date(kind, g, order="dmy"):
    """Convert a scanned date to datetime.date, or None if impossible/partial."""
    try:
        if kind == "iso":
            y, m, d = int(g[0]), int(g[1]), int(g[2])
        elif kind == "compact":
            y, m, d = int(g[0]), int(g[1]), int(g[2])
        elif kind == "num":
            a, b, y = int(g[0]), int(g[1]), int(g[2])
            if len(g[2]) == 2:
                y += 2000
            d, m = (a, b) if order == "dmy" else (b, a)
        elif kind == "dmy":
            d, m, y = int(g[0]), MONTHS[g[1].lower()], int(g[2])
        elif kind == "mdy":
            m, d, y = MONTHS[g[0].lower()], int(g[1]), int(g[2])
        else:                       # partial: no day
            return None
        return dt.date(y, m, d)     # rejects 31/04, 29 Feb 2026, ...
    except (ValueError, KeyError):
        return None


# ---------------------------------------------------------------- 4. ambiguity
def resolve(kind, g, context):
    """Return (date or None, note). note is '' when nothing needs review."""
    if kind == "partial":
        return None, "partial date with no day"
    if kind != "num":
        d = to_date(kind, g)
        return (d, "") if d else (None, "impossible date")

    dates = {o: to_date(kind, g, o) for o in ("dmy", "mdy")}
    orders = [o for o, d in dates.items() if d]
    if not orders:
        return None, "impossible date"
    if len({dates[o] for o in orders}) == 1:      # unambiguous (day > 12, or 03/03)
        return dates[orders[0]], ""

    # both readings are real dates: received must not be after revised/accepted/...
    def consistent(o):
        later = [to_date(ck, cg, o) for ck, cg in context]
        return all(c is None or c >= dates[o] for c in later)

    ok = [o for o in orders if consistent(o)]
    if len(ok) == 1:
        return dates[ok[0]], f"ambiguous numeric date, resolved by chronology as {ok[0]}"
    pick = "dmy" if "dmy" in orders else orders[0]
    return dates[pick], "ambiguous numeric date, chronology cannot decide: day-first guess, NEEDS REVIEW"


def extract_received(path):
    lines = read_lines(path)
    text = "\n".join(lines)
    cands = scan(text, RECEIVED)
    if not cands:
        return None, "no received date found"
    context = scan(text, CONTEXT)
    notes = []
    for kind, g in cands:                          # first usable candidate wins
        d, note = resolve(kind, g, context)
        if d:
            return d, note
        notes.append(note)
    return None, "; ".join(notes)


# ---------------------------------------------------------------- CLI
def main(argv):
    if len(argv) != 3:
        print("Usage: extract.py INPUT_FOLDER OUTPUT.csv", file=sys.stderr)
        return 2
    folder, out = Path(argv[1]), Path(argv[2])
    if not folder.is_dir():
        print(f"Not a folder: {folder}", file=sys.stderr)
        return 1
    files = sorted((p for p in folder.iterdir()
                    if p.is_file() and p.suffix.lower() == ".docx" and not p.name.startswith("~$")),
                   key=lambda p: p.name)
    rows = []
    for p in files:
        try:
            d, note = extract_received(p)
        except Exception as e:                     # corrupt, encrypted, not a zip, ...
            d, note = None, f"unreadable file ({type(e).__name__})"
        rows.append((p.name, d.isoformat() if d else "Invalid"))
        if note:
            print(f"{p.name}: {note}", file=sys.stderr)
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["file", "received_date"])
        w.writerows(rows)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))