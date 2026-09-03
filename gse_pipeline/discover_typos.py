import argparse
import re
from collections import Counter

import openpyxl

import config
import pipeline

_WORD_RE = re.compile(r"[A-Za-z']+")


def collect_all_descriptions(wb):
    """Pull every Defects Description across Master + every monthly sheet."""
    descs = []
    master_rows = pipeline.read_master(wb)
    descs += [r["Defects Description"] for r in master_rows]
    for month in config.MONTH_ORDER:
        rows = pipeline.read_month(wb, month)
        descs += [r["Defects Description"] for r in rows]
    return [d for d in descs if d]


def word_frequencies(descriptions):
    """Word-frequency count across all descriptions. Equipment codes are
    excluded by stripping the leading 'CODE / ' or 'CODE - ' pattern before
    tokenizing, same as how misspellings were originally found."""
    counter = Counter()
    for d in descriptions:
        body = re.sub(r"^\s*[A-Za-z0-9.\-/ ]{2,15}?[/\-]\s*", "", d, count=1)
        for w in _WORD_RE.findall(body):
            counter[w.upper()] += 1
    return counter


def find_candidates(counter, threshold, already_known):
    """Words occurring <= threshold times, excluding ones already handled
    (either as a known-bad spelling or as the correct target spelling)."""
    known_upper = {k.upper() for k in already_known}
    known_upper |= {v.upper() for v in config.SPELLING_MAP.values()}
    known_upper |= {w.upper() for w in config.MALAY_WORDS}

    candidates = [
        (w, c) for w, c in counter.items()
        if c <= threshold and w not in known_upper and len(w) > 2
    ]
    return sorted(candidates, key=lambda x: (-x[1], x[0]))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--threshold", type=int, default=8,
                         help="Flag words occurring this many times or fewer (default 8)")
    args = parser.parse_args()

    print(f"Loading {config.SOURCE_FILE}...")
    wb = openpyxl.load_workbook(config.SOURCE_FILE, data_only=True)

    descriptions = collect_all_descriptions(wb)
    print(f"Scanning {len(descriptions)} Defects Description entries...")

    counter = word_frequencies(descriptions)
    print(f"Distinct words found: {len(counter)}")

    known_bad = set(config.SPELLING_MAP.keys())
    candidates = find_candidates(counter, args.threshold, known_bad)

    print(f"\n{len(candidates)} words occur <= {args.threshold} times and are NOT already "
          f"in config.SPELLING_MAP or config.MALAY_WORDS.")
    print("Review this list manually - not every rare word is a typo (some are proper")
    print("nouns, brand names, or genuinely rare-but-correct terms). For anything you")
    print("confirm as a typo, add it to SPELLING_MAP in config.py as 'BADWORD': 'GOODWORD'.\n")

    print(f"{'Count':>6}  Word")
    print("-" * 30)
    for w, c in candidates:
        print(f"{c:6d}  {w}")


if __name__ == "__main__":
    main()
