# HiDPI remediation plan (Library Organizer Configure)

Code review findings from PR [#25](https://github.com/Stonepaw/comicrack-library-organizer/pull/25) and their resolution in **2.1.16**.

## Summary

| ID | Priority | Finding | Remedy (2.1.16) |
|----|----------|---------|-----------------|
| 1 | P1 | Options month/illegal rows overlap at 200% | **2.1.15** — vertical `y` chain with `coords_scaled=True` |
| 2 | P1 | Search tab restore drops label layout | **2.1.15** — `_apply_insert_control_layout` + `RefreshLabelLayout` |
| 3 | P1 | Non-`Point` `Tag` crashes relayout | **2.1.15** — `isinstance(baseline, Point)` guard |
| 4 | P2 | Rules / Yes-No / Calculated tabs use fixed coordinates | `relayout_rules_page()`, single-column stack for Yes-No/Calculated, metadata `apply_hidpi_metrics` |
| 5 | P2 | Empty values tab uses fixed Y positions | Vertical flow chain in `relayout_options_page()` |
| 6 | P2 | DPI at construct time without form `owner` | `apply_hidpi_metrics(owner)` on all insert/metadata controls at relayout |
| 7 | P2 | No relayout on monitor/DPI change while open | Documented in `lodpi.py` module doc — reopen Configure |
| 8 | P2 | Fixed `WIDE_ROW_WIDTH` / `NARROW_ROW_WIDTH` heuristics | `measure_column_widths()` from `PreferredSize` per relayout pass |
| 9 | P3 | `get_scale()` called per control | Single `scale` cached per relayout method; passed into `layout_row` / grid helpers |
| 10 | P3 | `create_insert_controls` uses `scale_int` without `owner` | `owner=self` for insert tab shell metrics |
| 11 | P3 | No automated `lodpi` tests | Manual quickstart below (IronPython + WinForms; no headless harness) |

## Architecture

```
configure_form_load / change_page / insert_controls_selecting
        │
        ▼
  relayout_all_hidpi()
        ├── relayout_insert_control_grids()  ← measure widths, two-column + single-column
        ├── relayout_options_page()          ← Options + Empty values vertical flow
        └── relayout_rules_page()            ← folder list + metadata header + exclude controls
```

DPI source is always the Configure form handle (`owner=self`).

## Manual validation (150% and 200%)

1. Open **Library Organizer → Configure** (dialog title shows `2.1.16`).
2. **Files / Folders** — Text, Number, Multiple Value insert tabs: no column overlap; Search tab round-trip restores layout.
3. **Yes/No Fields** — Manga / Series Complete stack without overlap; instructions below controls.
4. **Calculated** — tall rows (e.g. Read %) stack; info label sits to the right of the column.
5. **Rules** — Folder Rules list + Add/Remove buttons aligned; Metadata Rules header row and rule rows fit panel width.
6. **Options** — month names, illegal characters, empty-folder list unchanged from 2.1.15 pass.
7. **Empty values** — substitution block and failed-empty block chain vertically without overlap.

## Known limitation

Moving the Configure window to a display with a different scale while the dialog stays open does **not** trigger relayout. Close and reopen Configure after changing Windows display scale or moving across monitors.
