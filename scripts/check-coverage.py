#!/usr/bin/env python3
"""Aggregate line coverage across all cobertura files via merged-union counting.

Parses every coverage.cobertura.xml, unions <line> entries by (filename, line-number)
keeping max hits, then groups by first path segment for per-project display.
Exit 0 = pass, 1 = below threshold, 2 = no coverage files found.
"""

import argparse
import glob
import os
import sys
import xml.etree.ElementTree as ET

SEP = os.sep


def _canonical_name_map(all_filenames):
    """Build a mapping from every filename to its canonical (longest-suffix) form.

    Covertura filenames are relative to each test project's directory. The same
    source file can appear as 'TypstBridge/managed/Foo.cs' in one report and
    'Foo.cs' in another (when the test project is TypstBridge/managed).
    The canonical form is the longest filename whose path suffix matches the
    short form; this merges lines from both reports under one key.
    """
    names = sorted(all_filenames, key=lambda x: -len(x))  # longest first
    canonical = {}
    for name in names:
        if name in canonical:
            continue
        parts = name.split("/")
        # Try to find shorter versions that are suffixes of this name
        for candidate in names:
            if candidate == name:
                continue
            if candidate in canonical:
                continue
            cand_parts = candidate.split("/")
            if len(cand_parts) >= len(parts):
                continue
            if parts[-len(cand_parts):] == cand_parts:
                canonical[candidate] = name
        canonical[name] = name
    return canonical


def union_merge(xml_files):
    """Parse all cobertura files and return merged per-file line data.

    Returns:
        file_stats: dict[filename] = {"covered": int, "valid": int, "rate": float}
        total_covered: int  (lines with hits > 0 across the union)
        total_valid: int    (unique instrumented lines across the union)
    """
    # Pass 1: collect all unique filenames and build canonical mapping
    all_filenames = set()
    for path in xml_files:
        tree = ET.parse(path)
        root = tree.getroot()
        for cls in root.iter("class"):
            fn = cls.get("filename", "")
            if fn:
                all_filenames.add(fn.replace("\\", "/"))

    name_map = _canonical_name_map(all_filenames)

    # Pass 2: build merged line data using canonical names
    global_lines = {}  # filename -> {line_number: max_hits}

    for path in xml_files:
        tree = ET.parse(path)
        root = tree.getroot()
        for cls in root.iter("class"):
            filename = cls.get("filename", "")
            if not filename:
                continue
            canonical = name_map.get(filename.replace("\\", "/"), filename)
            line_map = global_lines.setdefault(canonical, {})
            for line in cls.iter("line"):
                num = line.get("number")
                hits = int(line.get("hits", "0"))
                if num is None:
                    continue
                num = int(num)
                prev = line_map.get(num, -1)
                if hits > prev:
                    line_map[num] = hits

    file_stats = {}
    total_covered = 0
    total_valid = 0

    for filename, line_map in global_lines.items():
        valid = len(line_map)
        covered = sum(1 for h in line_map.values() if h > 0)
        file_stats[filename] = {
            "covered": covered,
            "valid": valid,
            "rate": 100.0 * covered / valid if valid else 0.0,
        }
        total_covered += covered
        total_valid += valid

    return file_stats, total_covered, total_valid


def project_from_filename(filename):
    """Extract the top-level project directory from a source filename.

    Example: 'PptxEditor.Core/Generation/Foo.cs' -> 'PptxEditor.Core'
    """
    parts = filename.replace("\\", "/").split("/")
    return parts[0] if parts else "unknown"


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--threshold", type=float, default=60.0,
        help="minimum total line coverage percent (default: 60)",
    )
    ap.add_argument(
        "--dir", default="coverage",
        help="coverage results directory (default: coverage)",
    )
    ap.add_argument(
        "--per-file", action="store_true",
        help="print the 20 lowest-covered files by valid lines",
    )
    args = ap.parse_args()

    xml_files = sorted(
        glob.glob(os.path.join(args.dir, "**", "coverage.cobertura.xml"), recursive=True)
    )

    if not xml_files:
        print(
            f"ERROR: no coverage.cobertura.xml files found under '{args.dir}/'",
            file=sys.stderr,
        )
        sys.exit(2)

    file_stats, total_covered, total_valid = union_merge(xml_files)

    # Group by project (first path segment)
    project_data = {}  # project -> {"covered": int, "valid": int}
    for filename, stats in file_stats.items():
        proj = project_from_filename(filename)
        pd = project_data.setdefault(proj, {"covered": 0, "valid": 0})
        pd["covered"] += stats["covered"]
        pd["valid"] += stats["valid"]

    # Sort projects by name
    sorted_projects = sorted(project_data.items(), key=lambda x: x[0].lower())

    print(f"{'project':<50} {'covered':>10} {'valid':>10} {'rate':>8}")
    print("-" * 82)
    for proj, data in sorted_projects:
        rate = 100.0 * data["covered"] / data["valid"] if data["valid"] else 0.0
        print(f"{proj:<50} {data['covered']:>10} {data['valid']:>10} {rate:>7.2f}%")
    print("-" * 82)

    total_rate = 100.0 * total_covered / total_valid if total_valid else 0.0
    print(f"{'TOTAL':<50} {total_covered:>10} {total_valid:>10} {total_rate:>7.2f}%")

    if args.per_file and file_stats:
        print()
        sorted_files = sorted(
            file_stats.items(),
            key=lambda kv: (kv[1]["rate"], -kv[1]["valid"]),
        )[:20]
        print("Lowest-covered files (top 20 by valid lines):")
        print(f"{'  file':<80} {'covered':>10} {'valid':>10} {'rate':>8}")
        print("-" * 112)
        for fname, stats in sorted_files:
            print(
                f"  {fname:<78} {stats['covered']:>10} {stats['valid']:>10}"
                f" {stats['rate']:>7.2f}%"
            )
        print()

    if total_rate < args.threshold:
        print(
            f"\nFAIL: total line coverage {total_rate:.2f}% < threshold"
            f" {args.threshold:.2f}%"
        )
        sys.exit(1)

    print(
        f"\nPASS: total line coverage {total_rate:.2f}% >= threshold"
        f" {args.threshold:.2f}%"
    )
    sys.exit(0)


if __name__ == "__main__":
    main()
