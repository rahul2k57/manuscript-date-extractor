# Received-date extractor

Finds the date a manuscript was **received** inside each `.docx` file and writes it as
`YYYY-MM-DD`. Anything that cannot be turned into a real date is written as `Invalid`.

Standard library only: needs **Python 3.12** (works on 3.8+), no `pip install`.

## Run

```bash
./run.sh input_folder output.csv
```

Reads every `.docx` in `input_folder` (top level) and writes a CSV:

```
file,received_date
merged_01.docx,2026-01-02
merged_08.docx,Invalid
```

One row per file, sorted by file name. Notes for a human reviewer (ambiguous dates, why a
file is `Invalid`) are printed to **stderr**, not into the CSV.

## Run the tests

```bash
python3 -m unittest discover -s tests -v
```

This also regenerates `test_report.csv` (columns `checks, input, expected, actual, result`).
Tests build their own small `.docx` files, so they run on a clean clone.
If the 20 manuscripts are in `samples/` (or `SAMPLES_DIR=/path`), they are also checked
against `tests/answer_key.csv`; without them those checks are reported as `skip`.

## How it works

1. **Read**: a `.docx` is a zip of XML. We read `word/document.xml` plus headers, footers and
   footnotes, and join all text runs of each paragraph (so a bold label followed by a plain date
   still matches). Table cells are paragraphs, so label and value cells are adjacent.
2. **Find**: the word *received* (case-insensitive, `recieved`/`Recvd.` accepted) followed
   directly by a date. `Received in revised form` is skipped. Prose such as
   "we received 120 samples on 5 May 2025" is ignored because no date follows the word.
3. **Parse**: handles `2 Jan 2026`, `2nd January 2026`, `January 2, 2026`, `12-Mar-2023`,
   `2026-01-02`, `2023/03/12`, `03/04/2026`, `21.12.2026`, `12/05/26`, `20230312`, `Sept.`.
   `datetime.date` rejects impossible dates such as 31 April.
4. **Resolve DD/MM vs MM/DD** for numeric dates where both readings are real dates: pick the
   reading that keeps Received on or before the Revised / Accepted / Published dates. If both
   fit (or neither), use day-first and log `NEEDS REVIEW`.

## Assumptions

- Only the original **Received** date counts. Submitted, Accepted, Revised, Published online,
  Available online, Issue date and Accessed are not the received date.
- If several valid "Received" dates exist, the first in the document wins.
- Two-digit years mean 20xx.
- Output is `Invalid` for: no received date, corrupt/unreadable file, impossible date
  (`31/04/2026`), partial date with no day (`Mar 2026`).
- Ambiguous numeric dates that chronology cannot decide are read **day-first**
  (affects sample files 05, 07 and 18).
- A future year on another label (e.g. Accepted 2031 in file 18) does not affect Received.

## Layout

```
run.sh                 entry point: ./run.sh input_folder output.csv
src/extract.py         the tool
tests/test_extract.py  unittest suite (writes test_report.csv)
tests/docx_builder.py  builds small .docx files for the tests
tests/answer_key.csv   hand-checked expected value for each of the 20 samples
test_report.csv        output of the last test run
```