#!/usr/bin/env python3
"""
RMSE measurement tool for OfficeEditor PPTX render validation.

Compares OfficeEditor's PPTX→PNG render against a reference PDF (via pdftocairo)
on a per-slide basis using ImageMagick's RMSE metric.

Stdlib + subprocess only — no extra dependencies.

Usage:
  python3 scripts/rmse.py <deck.pptx> [--format md|json] [--out DIR]
                           [--no-render] [--force] [--size WxH]
"""

import argparse
import json
import math
import os
import re
import shutil
import statistics
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CONVERT_TOOL = ["dotnet", "run", "--project", str(REPO_ROOT / "tools" / "convert-pptx")]

ASPECT_TARGETS = {
    (2000, 1125): (2000, 1125),   # 16:9
    (1600, 1200): (1600, 1200),   # 4:3
    (1500, 1125): (1600, 1200),   # 4:3 (alt resolution from old renders)
}


def _die(msg: str, code: int = 2) -> None:
    print(f"ERROR: {msg}", file=sys.stderr)
    raise SystemExit(code)


def _which(cmd: str) -> str | None:
    return shutil.which(cmd)


def _ensure(cmd: str) -> str:
    path = _which(cmd)
    if not path:
        _die(f"{cmd} not found on PATH — please install it")
    return path


def _slide_number(filepath: Path) -> int | None:
    m = re.search(r"(?:slide|page)[_-](\d+)", filepath.name)
    return int(m.group(1)) if m else None


def _sort_slides(paths: list[Path]) -> list[Path]:
    return sorted(paths, key=lambda p: _slide_number(p) or 0)


def _detect_aspect(png_path: Path) -> tuple[int | None, int | None]:
    """Return (width, height) of a PNG using sips (macOS)."""
    r = subprocess.run(
        ["sips", "-g", "pixelWidth", "-g", "pixelHeight", str(png_path)],
        capture_output=True, text=True,
    )
    w = h = None
    for line in r.stdout.splitlines():
        m = re.search(r"pixelWidth:\s*(\d+)", line)
        if m:
            w = int(m.group(1))
        m = re.search(r"pixelHeight:\s*(\d+)", line)
        if m:
            h = int(m.group(1))
    return w, h


def _classify_aspect(w: int, h: int) -> tuple[int, int]:
    """Return (target_w, target_h) based on detected pixel dimensions."""
    key = (w, h)
    if key in ASPECT_TARGETS:
        return ASPECT_TARGETS[key]
    ratio = w / h if h else 1.0
    if ratio >= 1.5:
        return (2000, 1125)
    else:
        return (1600, 1200)


def _find_sibling_pdf(pptx_path: Path) -> Path | None:
    """Look for a .pdf next to the .pptx with the same stem."""
    pdf = pptx_path.with_suffix(".pdf")
    if pdf.is_file():
        return pdf
    deck_dir = pptx_path.parent
    for candidate in deck_dir.glob("*.pdf"):
        if candidate.stem == pptx_path.stem:
            return candidate
    return None


def _find_ref_png_dir(pptx_path: Path) -> Path | None:
    """Look for <deckdir>/ref-png/."""
    ref_dir = pptx_path.parent / "ref-png"
    if ref_dir.is_dir():
        slides = list(ref_dir.glob("slide-*.png"))
        if slides:
            return ref_dir
    return None


def _render_ours(pptx_path: Path, ours_dir: Path) -> list[Path]:
    """Render deck via convert-pptx → ours_dir. Returns list of generated PNGs."""
    ours_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [*CONVERT_TOOL, str(pptx_path), str(ours_dir), "--format", "png"],
        capture_output=True, text=True, timeout=300,
        cwd=REPO_ROOT,
    )
    if result.returncode != 0:
        _die(f"convert-pptx failed:\n{result.stderr}\n{result.stdout}")
    # convert-pptx writes one page-NNN.png per slide (shared naming with convert-xlsx).
    return _sort_slides(list(ours_dir.glob("page-*.png")))


def _slide_count_from_pdf(pdf_path: Path) -> int:
    """Use pdfinfo to get page count."""
    r = subprocess.run(
        ["pdfinfo", str(pdf_path)], capture_output=True, text=True
    )
    m = re.search(r"Pages:\s+(\d+)", r.stdout)
    if m:
        return int(m.group(1))
    return 0


def _build_refs(pdf_path: Path, work_ref_dir: Path, slide_count: int) -> list[Path]:
    """Generate ref PNGs from PDF via pdftocairo. Normalize naming to slide-NNN.png."""
    if work_ref_dir.exists():
        shutil.rmtree(work_ref_dir)
    work_ref_dir.mkdir(parents=True, exist_ok=True)

    pdftocairo = _ensure("pdftocairo")
    subprocess.run(
        [pdftocairo, "-png", "-r", "144", str(pdf_path), str(work_ref_dir / "slide")],
        capture_output=True, text=True, timeout=120, check=True,
    )
    # pdftocairo generates slide-1.png, slide-2.png, ...
    # Normalize to slide-NNN.png (3-digit, zero-padded)
    for raw in sorted(work_ref_dir.glob("slide-*.png"), key=lambda p: _slide_number(p) or 0):
        num = _slide_number(raw)
        if num is None:
            continue
        target = work_ref_dir / f"slide-{num:03d}.png"
        if raw != target:
            raw.rename(target)

    refs = _sort_slides(list(work_ref_dir.glob("slide-*.png")))
    if len(refs) != slide_count:
        raise RuntimeError(
            f"Ref PNG count mismatch: expected {slide_count}, got {len(refs)}"
        )
    return refs


def _resolve_refs(pptx_path: Path, work_ref_dir: Path, slide_count: int) -> list[Path]:
    """Resolve reference PNGs: prefer existing ref-png/ (copied), else pdftocairo from sibling PDF."""
    # Check for a sibling PDF
    pdf_path = _find_sibling_pdf(pptx_path)
    if pdf_path is None:
        _die(f"No sibling PDF found next to {pptx_path}")

    pdf_slides = _slide_count_from_pdf(pdf_path)
    if pdf_slides and pdf_slides != slide_count:
        _die(
            f"Slide count mismatch: PPTX has {slide_count} slides, "
            f"PDF has {pdf_slides} pages"
        )

    ref_png_dir = _find_ref_png_dir(pptx_path)
    if ref_png_dir is not None:
        # Copy/link existing ref-png into work area
        if work_ref_dir.exists():
            shutil.rmtree(work_ref_dir)
        work_ref_dir.mkdir(parents=True, exist_ok=True)
        for src in _sort_slides(list(ref_png_dir.glob("slide-*.png"))):
            dst = work_ref_dir / src.name
            shutil.copy2(src, dst)
        refs = _sort_slides(list(work_ref_dir.glob("slide-*.png")))
        if len(refs) != slide_count:
            _die(
                f"Ref PNG count mismatch: expected {slide_count}, "
                f"got {len(refs)} in {ref_png_dir}"
            )
        return refs

    # Build refs from PDF
    print(f"  Building refs from {pdf_path.name} via pdftocairo...", file=sys.stderr)
    return _build_refs(pdf_path, work_ref_dir, slide_count)


def _match_slides(ours: list[Path], refs: list[Path]) -> list[tuple[int, Path, Path]]:
    """Match ours and ref PNGs by slide number. Returns [(slide_num, ours_path, ref_path)]."""
    ours_map = {}
    for p in ours:
        n = _slide_number(p)
        if n is not None:
            ours_map[n] = p

    ref_map = {}
    for p in refs:
        n = _slide_number(p)
        if n is not None:
            ref_map[n] = p

    common = sorted(set(ours_map) & set(ref_map))
    if not common:
        _die("No matching slide numbers between ours and ref images")
    if len(common) < len(ours_map) or len(common) < len(ref_map):
        ours_only = sorted(set(ours_map) - set(ref_map))
        ref_only = sorted(set(ref_map) - set(ours_map))
        if ours_only:
            print(f"  Warning: ours-only slides: {ours_only}", file=sys.stderr)
        if ref_only:
            print(f"  Warning: ref-only slides: {ref_only}", file=sys.stderr)

    return [(n, ours_map[n], ref_map[n]) for n in common]


def _rmse_compare(ours_png: Path, ref_png: Path, size_spec: str, tmp_dir: Path) -> float | None:
    """Resize both to same size, then compute RMSE via magick compare."""
    ours_resized = tmp_dir / f"_rmse_ours_{ours_png.stem}.png"
    ref_resized = tmp_dir / f"_rmse_ref_{ref_png.stem}.png"

    subprocess.run(
        ["magick", str(ours_png), "-resize", size_spec, str(ours_resized)],
        capture_output=True, timeout=30, check=True,
    )
    subprocess.run(
        ["magick", str(ref_png), "-resize", size_spec, str(ref_resized)],
        capture_output=True, timeout=30, check=True,
    )

    r = subprocess.run(
        ["magick", "compare", "-metric", "RMSE", str(ours_resized), str(ref_resized), "null:"],
        capture_output=True, text=True, timeout=60,
    )
    # Output on stderr: "NNNNN.N (0.NNNNNN)"
    output = r.stderr or r.stdout
    m = re.search(r"\(([\d.]+)\)", output)
    if m:
        return float(m.group(1))
    return None


def _compute_stats(values: dict[int, float]) -> dict:
    """Compute aggregate statistics from per-slide RMSE values (raw 0..1 floats)."""
    vals = sorted(values.values())
    if not vals:
        raise RuntimeError("No RMSE values to compute stats from")

    n = len(vals)
    mean = sum(vals) / n
    median = statistics.median(vals) if n % 2 else (
        vals[n // 2 - 1] + vals[n // 2]) / 2

    def pctile(p: float) -> float:
        idx = int(n * p)
        return vals[min(idx, n - 1)]

    p85 = pctile(0.85)
    p90 = pctile(0.90)

    count_lt10 = sum(1 for v in vals if v < 0.10)
    count_lt15 = sum(1 for v in vals if v < 0.15)

    worst_5 = sorted(values.items(), key=lambda x: x[1], reverse=True)[:5]

    return {
        "n": n,
        "mean": mean,
        "median": median,
        "p85": p85,
        "p90": p90,
        "count_lt10": count_lt10,
        "count_lt15": count_lt15,
        "worst_5": worst_5,
    }


def _format_pct(val: float) -> str:
    return f"{val * 100:.1f}%"


def _format_md(deck_name: str, per_slide: dict[int, float], stats: dict) -> str:
    lines = []
    lines.append(f"## RMSE — {deck_name}")
    lines.append("")
    lines.append(f"Slides: {stats['n']} | "
                 f"Mean: {_format_pct(stats['mean'])} | "
                 f"Median: {_format_pct(stats['median'])} | "
                 f"P85: {_format_pct(stats['p85'])} | "
                 f"P90: {_format_pct(stats['p90'])}")
    lines.append(f"<10%: {stats['count_lt10']} | "
                 f"<15%: {stats['count_lt15']}")
    lines.append("")
    lines.append("### Worst 5")
    lines.append("")
    lines.append("| # | RMSE % |")
    lines.append("|---|--------|")
    for slide_num, rmse in stats["worst_5"]:
        lines.append(f"| {slide_num} | {_format_pct(rmse)} |")
    lines.append("")
    lines.append("### All slides")
    lines.append("")
    lines.append("| # | RMSE % |")
    lines.append("|---|--------|")
    for num, rmse in per_slide.items():
        lines.append(f"| {num} | {_format_pct(rmse)} |")
    return "\n".join(lines)


def _format_json(deck_name: str, per_slide: dict[int, float], stats: dict) -> str:
    output = {
        "deck": deck_name,
        "slides": stats["n"],
        "mean": round(stats["mean"] * 100, 1),
        "median": round(stats["median"] * 100, 1),
        "p85": round(stats["p85"] * 100, 1),
        "p90": round(stats["p90"] * 100, 1),
        "count_lt10": stats["count_lt10"],
        "count_lt15": stats["count_lt15"],
        "per_slide": {
            str(k): round(v * 100, 1) for k, v in sorted(per_slide.items())
        },
        "worst_5": [
            {"slide": num, "rmse": round(rmse * 100, 1)}
            for num, rmse in stats["worst_5"]
        ],
    }
    return json.dumps(output, indent=2)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Measure per-slide RMSE between OfficeEditor render and reference PDF"
    )
    parser.add_argument("deck", type=Path, help="Path to .pptx file")
    parser.add_argument(
        "--format", choices=["md", "json"], default="md",
        help="Output format (default: md)"
    )
    parser.add_argument(
        "--out", type=Path, default=None,
        help="Working output directory (default: /tmp/rmse-<deckname>)"
    )
    parser.add_argument(
        "--no-render", action="store_true",
        help="Skip rendering; expect ours/ already populated in --out"
    )
    parser.add_argument(
        "--force", action="store_true",
        help="Force re-render even if ours/ already has the right slide count"
    )
    parser.add_argument(
        "--size", type=str, default=None,
        help="Override resize target, e.g. 2000x1125! (the ! forces exact size)"
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()

    # --- Prerequisites ---
    _ensure("magick")
    if not args.no_render:
        _ensure("dotnet")
    _ensure("sips")

    deck_path = args.deck.resolve()
    if not deck_path.is_file():
        _die(f"Deck not found: {deck_path}")
    if deck_path.suffix.lower() != ".pptx":
        _die(f"Expected a .pptx file, got: {deck_path}")

    deck_name = deck_path.stem
    work_dir = args.out or Path(f"/tmp/rmse-{deck_name}")
    ours_dir = work_dir / "ours"
    ref_dir = work_dir / "ref"
    tmp_dir = work_dir / "_tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)

    # --- 1. Render ours ---
    if args.no_render:
        ours_pngs = _sort_slides(list(ours_dir.glob("page-*.png")))
        if not ours_pngs:
            _die(f"--no-render but no page-*.png found in {ours_dir}")
        print(f"  Using existing render: {len(ours_pngs)} slides in {ours_dir}", file=sys.stderr)
    else:
        existing = _sort_slides(list(ours_dir.glob("page-*.png")))
        need_render = args.force or not existing
        if not need_render:
            # Check slide count matches PDF
            pdf = _find_sibling_pdf(deck_path)
            if pdf:
                pdf_slides = _slide_count_from_pdf(pdf)
                if pdf_slides and len(existing) != pdf_slides:
                    need_render = True
            if not need_render:
                print(
                    f"  Skipping render — {len(existing)} slides exist in {ours_dir} "
                    f"(use --force to re-render)",
                    file=sys.stderr,
                )
        if need_render:
            print(f"  Rendering {deck_path.name}...", file=sys.stderr)
            ours_pngs = _render_ours(deck_path, ours_dir)
        else:
            ours_pngs = existing

    slide_count = len(ours_pngs)
    if slide_count == 0:
        _die("No rendered slides found after render step")

    # --- 2. Resolve references ---
    ref_pngs = _resolve_refs(deck_path, ref_dir, slide_count)
    print(f"  References: {len(ref_pngs)} slides", file=sys.stderr)

    # --- 3. Match slides ---
    matched = _match_slides(ours_pngs, ref_pngs)
    print(f"  Matched: {len(matched)} slides", file=sys.stderr)

    # --- 4. Detect aspect & determine resize target ---
    first_ours = matched[0][1]
    w, h = _detect_aspect(first_ours)
    if w is None or h is None:
        _die(f"Could not detect dimensions of {first_ours}")
    target_w, target_h = _classify_aspect(w, h)
    if args.size:
        size_spec = args.size
    else:
        size_spec = f"{target_w}x{target_h}!"
    print(f"  Resize target: {size_spec} (ours: {w}x{h})", file=sys.stderr)

    # --- 5. Measure ---
    print(f"  Measuring RMSE...", file=sys.stderr)
    per_slide = {}
    for i, (slide_num, ours_png, ref_png) in enumerate(matched):
        rmse = _rmse_compare(ours_png, ref_png, size_spec, tmp_dir)
        if rmse is None:
            print(f"  Warning: RMSE failed for slide {slide_num}", file=sys.stderr)
            per_slide[slide_num] = 1.0  # worst case
        else:
            per_slide[slide_num] = rmse
        if (i + 1) % 10 == 0 or (i + 1) == len(matched):
            print(f"    {i + 1}/{len(matched)} slides done", file=sys.stderr)

    # --- 6. Stats ---
    stats = _compute_stats(per_slide)
    per_slide_sorted = dict(sorted(per_slide.items()))

    # --- 7. Output ---
    if args.format == "json":
        print(_format_json(deck_name, per_slide_sorted, stats))
    else:
        print(_format_md(deck_name, per_slide_sorted, stats))

    # Cleanup tmp
    shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    try:
        main()
    except (subprocess.TimeoutExpired, subprocess.CalledProcessError) as e:
        _die(f"Subprocess error: {e}")
    except KeyboardInterrupt:
        _die("Interrupted", code=130)
