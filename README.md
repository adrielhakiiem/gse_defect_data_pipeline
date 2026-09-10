# EDR CMMS Cleaning

This project cleans and publishes the current CMMS export:

`EDR (JAN-AUG) 2026.xls`

It does not merge monthly files, reconcile a master workbook, or categorize equipment.
After local review, the cleaned workbook can be published to the configured
SharePoint-synced folder.

## Process

There are three scripts. Run them from `gse_pipeline/`.

### 1. Read and inspect the source

```bash
cd gse_pipeline
python read_source_v2.py
```

This reads the `Worksheet` sheet and prints:

- total non-blank rows
- rows without an equipment code from either the description or Asset Number
- examples of unresolved descriptions
- row counts by Work Order Type
- row counts by Work Status

The cleaned records are returned as Python dictionaries by `read_and_normalize()`. The output fields are:

`Permit No`, `Work Status`, `Equipment Code`, `Asset Number`, `Asset Description`, `Defects Description`, `Work Order Type`, `Work Order Trade`, `Action Taken`, `Date`, `By Month`

### 2. Review rare words

```bash
python discover_typos_v2.py
```

This reads the same source, counts words in `Defects Description`, and prints words used at most eight times. It excludes words already present in `config.SPELLING_MAP` or `config.MALAY_WORDS`.

Use a different threshold when needed:

```bash
python discover_typos_v2.py --threshold 3
```

Rare words are suggestions for human review. They are not automatically changed.

### 3. Clean and create the local review workbook

```bash
python clean_v2.py
```

This runs the cleaning steps in order and writes `EDR_JANAUG_2026_CLEANED.xlsx`
to the active output path. The current configuration publishes to the synced
SharePoint folder after local review.

The workbook contains:

- `Cleaned Data` — normalized records, with audit flag columns retained
- `QA Summary` — counts for every cleaning step
- `Non-Defect Review` — administrative rows flagged but not deleted
- `Near-Duplicate Review` — groups that need human confirmation
- `Exact Duplicate Review` — exact duplicates removed from the cleaned data
- `Malay Review` — rows containing configured Malay terms

Known entries in `SPELLING_MAP` are corrected only as whole words. Valid terms
such as `TIRE`, `TIRES`, and `BOGEY` are listed in `KNOWN_CORRECT_WORDS` and
are not silently changed. Operational codes such as `KE`, `KUL`, and `ALG` are
also excluded from Malay and typo flags. Rare words are printed as candidates;
they are never automatically added as corrections.
Review the candidate list and the review tabs before using the workbook in
Power BI. Exact duplicates are removed, but near-duplicates and administrative
rows are retained or flagged so no potentially meaningful record is silently lost.

The QA Summary also reports equipment-code normalization separately for
`Equipment Code` and `Asset Number`, plus the number of identifiers still
containing a decimal point. That residual count must be zero before dashboard use.

## Configuration

Install the only required package:

```bash
python -m pip install -r requirements.txt
```

The active paths are in `gse_pipeline/config.py`:

- The synced SharePoint source path is active by default and resolves to a local
	file through `config_local.py`.
- A repository-local source line is directly below it and commented out.
- The local output line is active by default.
- The SharePoint output line is directly below it and commented out.

The SharePoint output path is active after the local workbook was inspected.
To create another local review copy, comment the SharePoint output line and
uncomment the repository-local output line in `config.py`.

The source workbook, cleaned workbook, private paths, and generated caches are
excluded from Git. Only code, configuration templates, and documentation are
intended for GitHub.

## Files

```text
gse_pipeline/
├── config.py                 # active paths and review word lists
├── config_local.example.py   # private SharePoint path template
├── config_local.py           # private local settings; ignored by Git
├── read_source_v2.py         # source reader and normalization summary
├── discover_typos_v2.py      # rare-word discovery report
├── clean_v2.py                # cleaning, QA flags, deduplication, and local export
└── requirements.txt           # xlrd and openpyxl
```

Do not run any other script. The old consolidation pipeline has been removed.
