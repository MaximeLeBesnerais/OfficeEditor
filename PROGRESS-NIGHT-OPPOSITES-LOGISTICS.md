# Night Opposites / Logistics Progress

## 2026-07-28

- Baseline rendered and measured both local decks with `scripts/rmse.py`:
  - Opposites: 26 slides, mean 8.5%, worst 12.9% (slide 26).
  - Logistics: 18 slides, mean 8.7%, worst 12.9% (slide 18).
- Visually swept all 44 reference/render slide pairs. The material non-font issues were
  faceted custom freeforms, omitted default-width connector rules, and unhandled regular
  shape `flipH`/`flipV` transforms. Opposites slide 5 grouped icons are present after the
  earlier `grpFill`/style-reference fixes already on this branch.
- Implemented a focused converter fix and synthetic XML regression coverage. The post-fix
  deck render/RMSE sweep was not completed before the 12:00 CEST cutoff; this branch is
  intentionally left as WIP.
