"""Find rare words in the v2 CMMS defect descriptions."""

import argparse
import re
from collections import Counter

import config
import read_source_v2


_WORD_RE = re.compile(r"[A-Za-z']+")


def word_frequencies(descriptions):
    counter = Counter()
    for description in descriptions:
        for word in _WORD_RE.findall(description or ""):
            counter[word.upper()] += 1
    return counter


def find_candidates(counter, threshold):
    known_upper = {word.upper() for word in config.SPELLING_MAP}
    known_upper.update(value.upper() for value in config.SPELLING_MAP.values())
    known_upper.update(word.upper() for word in config.KNOWN_CORRECT_WORDS)
    known_upper.update(word.upper() for word in config.MALAY_WORDS)
    return sorted(
        ((word, count) for word, count in counter.items()
         if count <= threshold and len(word) > 2 and word not in known_upper),
        key=lambda item: (-item[1], item[0]),
    )


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--threshold", type=int, default=8,
                        help="Flag words occurring this many times or fewer (default 8)")
    parser.add_argument("--source", default=config.SOURCE_FILE)
    args = parser.parse_args()

    print(f"Loading {args.source}...")
    rows, _ = read_source_v2.read_and_normalize(args.source)
    descriptions = [row["Defects Description"] for row in rows]
    print(f"Scanning {len(descriptions)} Defects Description entries...")

    counter = word_frequencies(descriptions)
    print(f"Distinct words found: {len(counter)}")
    candidates = find_candidates(counter, args.threshold)
    print(f"\n{len(candidates)} words occur <= {args.threshold} times and are NOT already "
          "in config.SPELLING_MAP or config.MALAY_WORDS.")
    print("Review this list manually - not every rare word is a typo (some are proper")
    print("nouns, brand names, or genuinely rare-but-correct terms). For anything you")
    print("confirm as a typo, add it to SPELLING_MAP in config.py as 'BADWORD': 'GOODWORD'.\n")
    print(f"{'Count':>6}  Word")
    print("-" * 30)
    for word, count in candidates:
        print(f"{count:6d}  {word}")


if __name__ == "__main__":
    main()