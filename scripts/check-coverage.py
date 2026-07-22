#!/usr/bin/env python3
"""Aggregate line coverage across all cobertura files in coverage/ and enforce a threshold.

Weighted by valid lines (sum covered / sum valid), not an average of rates.
Exit 0 = pass, 1 = below threshold, 2 = no coverage files found.
"""

import argparse
import glob
import os
import sys
import xml.etree.ElementTree as ET


def parse_file(path):
    tree = ET.parse(path)
    root = tree.getroot()
    valid = covered = 0
    for line in root.iter("line"):
        valid += 1
        if int(line.get("hits", "0")) > 0:
            covered += 1
    if valid == 0:
        rate = float(root.get("line-rate", "0"))
        valid = 1
        covered = int(round(rate))
    return valid, covered


def project_name(path, base):
    rel = os.path.relpath(path, base)
    parts = rel.split(os.sep)
    return parts[-2] if len(parts) >= 2 else rel


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--threshold", type=float, default=60.0,
                    help="minimum total line coverage percent (default: 60)")
    ap.add_argument("--dir", default="coverage",
                    help="coverage results directory (default: coverage)")
    args = ap.parse_args()

    files = sorted(glob.glob(os.path.join(args.dir, "**", "coverage.cobertura.xml"),
                             recursive=True))
    if not files:
        print(f"ERROR: no coverage.cobertura.xml files found under '{args.dir}/'",
              file=sys.stderr)
        sys.exit(2)

    total_valid = total_covered = 0
    print(f"{'project':<50} {'covered':>10} {'valid':>10} {'rate':>8}")
    print("-" * 82)
    for f in files:
        valid, covered = parse_file(f)
        total_valid += valid
        total_covered += covered
        rate = 100.0 * covered / valid
        print(f"{project_name(f, args.dir):<50} {covered:>10} {valid:>10} {rate:>7.2f}%")
    print("-" * 82)

    total_rate = 100.0 * total_covered / total_valid
    print(f"{'TOTAL':<50} {total_covered:>10} {total_valid:>10} {total_rate:>7.2f}%")

    if total_rate < args.threshold:
        print(f"\nFAIL: total line coverage {total_rate:.2f}% < threshold "
              f"{args.threshold:.2f}%")
        sys.exit(1)
    print(f"\nPASS: total line coverage {total_rate:.2f}% >= threshold "
          f"{args.threshold:.2f}%")
    sys.exit(0)


if __name__ == "__main__":
    main()
