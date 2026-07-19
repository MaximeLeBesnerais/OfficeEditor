# OfficeEditor — Pitch Deck

> For internal presentation: why this open-source project replaces our
> Python + LibreOffice rendering pipeline, what it delivers today, the
> metrics to prove it, and what's next.

---

## 1. The one-liner

**OfficeEditor** replaces our legacy slide-rendering microservice with an
in-process C# library. Same output, milliseconds instead of seconds, zero
external services.

---

## 2. The thesis that drives everything

> Instant preview beats perfect preview.
> &lt;500ms at ~12% RMSE wins over 15 seconds at 2% RMSE.

The preview is for **iteration**. The rebuilt PPTX is the **deliverable**.
Fidelity work serves the preview loop first; editing correctness serves the
download second. That priority order is settled.

---

## 3. The one chart that matters

Performance comparison: our warm rendering pipeline vs headless LibreOffice
(macOS arm64, both REF decks).

| Deck | OfficeEditor warm | LibreOffice warm | Speedup |
|---|---|---|---|
| REMOVED (8 slides, 1440×810pt) | **105.1ms** | 1336.6ms | **12.7×** |
| pres-pro (16 slides, 960×540pt) | **117.2ms** | 1807.5ms | **15.4×** |

Per-slide derived: OfficeEditor 13.1/7.3ms vs LibreOffice 167.1/113.0ms.

Cached single-slide preview via the HTTP API: **~1.5ms** (ETag/cache hit).
Cold first request: ~230ms. Both under the 500ms product target with wide margin.

Methodology: `Stopwatch` + `Process` only, no third-party benchmark deps.
LibreOffice measured as `soffice --headless --convert-to pdf` cold vs warm
(median of N=5). Raw data committed at `tools/pptx-benchmark/baselines/`.

---

## 4. Visual fidelity today

Pixel-diff (ImageMagick RMSE, both REF decks at 150 PPI vs PowerPoint reference PDFs):

| Deck | Avg RMSE | Page 1 RMSE | Notes |
|---|---|---|---|
| REMOVED | **10.52%** | 8.61% | Background gradient + circle artwork correct |
| pres-pro | **17.19%** | 13.23% | Text-heavy financial deck; font metrics dominate |
| pitch-deck | **16.03%** | 18.17% | 12-slide corporate format, this document's source |

Methodology: `visual-diff --suite pptx --generate` → `pdftocairo` page
rasters @ 150 PPI → `compare -metric RMSE` per page. The metrics.json
baseline is committed and the gate script (`--check --baseline --margin`)
fails on regression. No script tolerates silent drift.

---

## 5. What shipped (Phase 4)

**12 parallel workstreams, 4 staged merge waves. 844 unit/integration tests.
One forced regression caught and fixed before closing.**

### Rendering backend

| Capability | Status |
|---|---|
| Per-slide PNG/SVG preview (`GET /api/decks/{id}/slides/{n}/preview`) | Done |
| Process-wide system font cache | Done |
| Per-slide Typst emission (compile only the edited slide) | Done |
| Persistent warm TypstBridge sessions (ABI v3; Rust + managed) | Done |
| comemo eviction policy (bounded memory in long-lived server) | Done |
| Unsupported-element placeholders (charts, SmartArt → visible badge + warnings[]) | Done |

### Editing surface

| Capability | Status |
|---|---|
| `/api/decks/{id}/anatomy` — elements with xfrm positions | Done |
| `/api/decks/{id}/instructions` — replaceText/Image/Table, move/dup/delete slides | Done |
| Instruction engine: validate-then-execute, per-op errors, revision counter | Done |
| Style-preserving text replacement | Done |
| Shared-ImagePart-safe image replacement | Done |
| Image fit modes: fill (center-crop srcRect), crop, contain | Done |
| `TextFitService`: wrap simulation + normAutofit emulation | Done |

### Infrastructure

| Capability | Status |
|---|---|
| `OfficeEditor.Mcp` stdio JSON-RPC host (AI agent tools: anatomize/edit/render) | Done |
| `tools/visual-diff` CLI: PNG-pair + `--suite pptx` + threshold gate | Done |
| `tools/pptx-benchmark` harness with committed pre-optimization baseline | Done |
| 5 test assemblies in the solution (844 tests, serial execution, REF-deck smoke gate) | Done |
| ReorderSlide off-by-one: confirmed, test-first fix | Done |

### One gap

- **F9 brand.json endpoint not wired** — `BrandProfileExtractor` exists + tested + snapshot fixtures committed; ~20 lines of `Program.cs` missing. Priority: wire it.

---

## 6. Development process (12 workstreams, 4 waves)

How we built it — reproducible, conflict-free, gated at every merge.

```
W1   replacer bug fixes          ──┐
W2   slide ops + builder surface  ──┤
W8   benchmark harness + baseline ──┤  Wave 1 (zero deps, parallel)
W9   visual-diff tooling          ──┤
W10  test infrastructure          ──┘
           │
           ▼  Phase 0 merge + W8/W9/W10 merged first (W8 must land before W3)
           │
W3   converter: latency + correctness (CRITICAL PATH)
W6   API: deck store + preview
W7   API: anatomy + instruction engine
W5   TextFitService (differentiator)
W4   BrandProfileExtractor
           │
           ▼  Sequential merge W3→W6→W7→W5→W4; catalog-integration wiring
           │
W11  Rust bridge ABI v3 (font cache, persistent World)
W12  MCP host
           │
           ▼  1 regression caught post-merge (image native-size bug), fixed, re-gated
```

Method:
- One git worktree + feature branch per workstream. Strict file ownership;
  cross-file needs go through `/integration/*.patch`.
- `dotnet test DocxEditor.sln` after every merge. Both REF-deck smoke tests
  (REMOVED + pres-pro) as the non-negotiable gate.
- `visual-diff --suite pptx --baseline --margin` after converter-touching
  merges. RMSE may only move intentionally with documented reason.

Result: **0 integration-revert cycles**, 844 tests green per merge, no
conflicts requiring manual intervention beyond the pre-planned catalog
wiring step. One real bug (native image sizing) caught by visual-diff and
fixed within 1 commit.

---

## 7. Architecture (the employer's integration story)

**OfficeEditor.Api** is the integration surface:

```
POST /api/decks                              → deckId
GET  /api/decks/{id}/anatomy                 → element map (id/type/position/text)
POST /api/decks/{id}/instructions            → apply edits, return changed slide URLs
GET  /api/decks/{id}/slides/{n}/preview      → PNG/SVG with Etag/304
GET  /api/decks/{id}/file                    → rebuilt .pptx download
```

What the employer's existing Python service does:

```python
deck_id = post("/api/decks", files={"file": template_pptx})
anatomy = get(f"/api/decks/{deck_id}/anatomy")
instructions = {"operations": [{"type": "replaceText", "slide": 1, "elementId": 7, "text": new_title}]}
result = post(f"/api/decks/{deck_id}/instructions", json=instructions)
for s in result["changedSlides"]:
    png = get(s["previewUrl"])  # ~1.5ms cached
pptx_bytes = get(f"/api/decks/{deck_id}/file")
```

No LibreOffice. No licensing. No process management. No font drift between
serve and render environments. Stateless Core, stateful API layer (30-min
sliding TTL, per-deck lock).

---

## 8. Open issues (honest)

1. **RMSE is machine-dependent.** The committed baseline is macOS arm64,
   ImageMagick 7, poppler 26. Re-run on the employer's target Linux/x64 and
   update the doc.
2. **F9 brand.json endpoint not wired** — 20 lines of `Program.cs` missing.
3. **No LibreOffice comparison numbers on the employer's target platform.**
   soffice is installed here; the LO numbers above are macOS, not the
   actual deployment Linux image. Re-run there.
4. **Aptos font gap.** New Office default, no metric-compatible open clone.
   Carlito works for Calibri but NOT for Aptos. This is the "font bakery" Phase.

---

## 9. Next: font metric-fitting ("the bakery")

Preview fidelity breaks when the deck's fonts aren't embedded AND aren't on
the server. Aptos is the new Office default — it has NO open clone.

Plan:
1. **Font census first.** Run BrandProfileExtractor over all available
   client decks; count fonts. Likely Aptos/Calibri/Segoe UI/Arial/Georgia
   ≈95%. Calibri/Arial/Georgia already have open clones (Carlito, Liberation,
   Gelasio). The bakery likely only needs AptosCompat + SegoeCompat.
2. **Build adapted clones** (offline, per source font): source metrics
   (OpenTypeFontMetricsReader) + OFL donor (Lato/Open Sans for Latin) →
   weight adjust → per-glyph width fit → vertical metrics → adapted TTF.
   Metrics aren't IP; outlines are. Precedent: Carlito/Caladea. Never copy
   proprietary outlines. OFL + rename.
3. **Preview-only.** Adapted fonts never enter the output PPTX. PowerPoint
   does its own substitution on the downloaded file.
4. **Regression-tested.** Advance widths within tolerance. Line-break
   diff <1% on a wrap corpus (TextFitService machinery as the harness).

Estimated scope: L. Depends on the font census.

---

## 10. Cost summary

| Item | Units |
|---|---|
| Development workstreams | 12 |
| Test assemblies | 5 |
| Tests | 844 |
| Effort pattern | 1 coordinator + agents per workstream, 4 merge waves |
| External services | 0 (TypstBridge native lib ~37MB, statically linked) |
| New dependencies | 1 flagged (`comemo` direct dep, already transitive at same version) |
| Production dependencies | .NET 9, DocumentFormat.OpenXml, no LibreOffice |

---

[Generated 2026-07-18, all metrics from committed artifacts in this repository.]
