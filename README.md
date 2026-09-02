# GSE Defect Data Consolidation Pipeline

> **Work in Progress:** The core pipeline is functional and can be run
> end-to-end using the included synthetic dataset. This repository is still
> being actively refined, with further testing, documentation, and
> improvements planned.

Merges monthly Ground Support Equipment (GSE) defect-report sheets into a
single master dataset, resolves what's missing, standardizes inconsistent
text, and produces a dashboard/BI-ready Excel file with a full audit trail.

Built to solve a real problem: multiple monthly source sheets totalled
several thousand more rows than the existing "master" file — thousands of
rows had never been merged in, and the ones that had were riddled with
typos and inconsistent formatting.

> **Note:** This is a portfolio version of a project originally built for an
> internal operational dataset. The source data and output files are
> excluded from this repository (confidential); a small synthetic sample
> dataset is included instead so the pipeline can be run end-to-end.

## Skills demonstrated

- **Data reconciliation logic** — designed a composite-key matching
  approach (rather than a naive single-column join) to correctly identify
  missing records without creating false duplicates
- **Defensive / non-destructive data cleaning** — the pipeline documents
  ambiguous anomalies for human review instead of silently guessing at a fix
- **Systematic QA over manual spot-checking** — built a word-frequency-based
  discovery tool to catch data-quality issues at scale, rather than relying
  on eyeballing thousands of rows
- **Reusable, config-driven pipeline design** — built for handoff to future
  users with no code changes required for routine new-data updates
- **End-to-end documentation** — technical README, plus a separate
  stakeholder-facing one-pager for non-technical audiences

## Why this exists / the problem

- Several monthly report sheets, covering a multi-month period, totaling
  well over ten thousand data rows combined
- 1 "master" sheet that was supposed to already contain everything, but was
  missing a large share of that data
- Thousands of rows were simply missing from the master and needed to be
  identified and appended correctly
- The remaining data had dozens of recurring misspellings, inconsistent
  punctuation, inconsistent categorical values, and embedded non-English
  text mixed into English defect descriptions

Manually reconciling thousands of rows by eye wasn't realistic, and neither
was trusting a naive "row count" comparison — the numbers needed to be
exactly right before anyone could trust a dashboard built on top of them.

## Key design decisions (and why)

**1. Equipment number alone is not a safe match key.**
The same equipment can have multiple, separate defect entries on the same
day (e.g. a tractor gets both a tyre puncture and a battery fault logged
separately). Matching on equipment number alone would either wrongly drop
legitimate repeat entries or wrongly treat every repeat as "already in
master." The match key used is **Date + Equipment No + Defect
Description**, normalized (whitespace collapsed, case-insensitive) — this
correctly tells apart "same equipment, different problem" (keep both) from
"same equipment, same problem, entered twice" (a real duplicate).

**2. Reconciliation is count-based, not set-based.**
For every match key, the pipeline compares *how many times* it appears in
the monthly source vs. in Master. If a key appears 3 times in the source but
only 1 time in Master, exactly 2 extra rows get appended — not 0, not all 3.
This correctly handles legitimate repeats without either losing data or
creating duplicates.

**3. Never touch equipment codes when fixing spelling.**
The spelling-correction dictionary only ever replaces exact known-misspelled
words. It never runs a general spellchecker across the text, which would
risk "correcting" an equipment code into something else.

**4. A systematic discovery process, not a fixed list.**
Rather than relying only on a hand-built list of known misspellings, a
separate tool scans every word across the full dataset by frequency —
rare words are surfaced for human review rather than assumed correct. This
makes "how do you know you caught everything" answerable with a repeatable
process, not a one-time guess.

**5. Non-English text is flagged, never auto-translated.**
A number of rows contain non-English words embedded in otherwise-English
defect descriptions. Automatically translating free text risks changing the
technical meaning of a safety-relevant record. These rows are instead
highlighted for a human with domain knowledge to review later.

**6. Every anomaly is logged, not silently fixed.**
Where the pipeline found something questionable but ambiguous — a
pre-existing date typo in the original master file, a handful of rows with
a clearly wrong categorical value — it does **not** guess and silently
correct it. It copies the source value through as-is and documents the
issue in the QA sheet, so a human with the actual context makes the final
call.

## What the pipeline does, step by step

1. **Read** the Master sheet and each monthly sheet, normalizing each into
   the same column schema (handling cases where different months use
   slightly different header names for the same field)
2. **Reconcile**: find every row present in a monthly sheet but missing from
   Master, using the composite key described above, and append only those
3. **Standardize text**: fix dozens of recurring misspellings, punctuation
   inconsistencies, and mixed letter case — all case-preserving, all
   word-boundary safe
4. **Standardize categorical fields**: collapse near-duplicate values into
   one canonical value
5. **Flag non-English text** for later manual translation (does not
   translate)
6. **Convert the Date column** from text to a real date type, so BI tools /
   Excel can sort, filter, and do time-intelligence without extra measures
7. **Write the output workbook**: the cleaned Master data (with color-coded
   flags), an audit list of every appended row, a plain-English QA/anomalies
   write-up, and a Change Log with exact before/after counts for every text
   fix applied

## Usage

```bash
pip install openpyxl

# Step 1 (recommended before every run on new data): scan for misspellings
# the pipeline doesn't already know about
python discover_typos.py

# Step 2: run the actual merge + clean
python run_pipeline.py
```

Run against `sample_data.xlsx` (included) to see the pipeline work
end-to-end without needing the real dataset.

### `discover_typos.py` — "how do you know you caught every misspelling?"

The correction dictionary is a fixed list, so on its own it can only fix
misspellings it already knows about — a brand new typo in a future batch of
data wouldn't be caught automatically. `discover_typos.py` closes that gap:
it scans every word across every description row, counts how often each one
occurs, and prints anything that occurs rarely for a human to review. A word
used thousands of times is essentially never a typo; a word used once or
twice, especially one resembling a common word already in the data, has
exactly the profile of a typo.

This turns "we eyeballed it and think we got everything" into "we
systematically scanned every word in the dataset, and here's the reviewed
list of what we found" — a repeatable process rather than a one-time guess.

## Project structure

```text
gse_defect_data_pipeline/
├── config.py
├── pipeline.py
├── write_output.py
├── run_pipeline.py
├── discover_typos.py
├── sample_data.xlsx
└── README.md
```

## Limitations / things intentionally left for a human

- Non-English defect descriptions are flagged, not translated
- Anomalies with an ambiguous cause (e.g. a likely date typo in the source)
  are documented but not auto-corrected
- Rows with a clearly wrong categorical value are copied through as-is and
  flagged, not corrected, since the true value isn't knowable from the data
  alone
