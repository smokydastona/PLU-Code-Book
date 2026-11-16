# Complete Till Binder — Audit Report

Status: **REVIEW ONLY — NOT APPROVED FOR OPERATIONAL USE.** Technical checks pass; production blockers remain open (`PRODUCTION_BLOCKERS.csv`).
Machine-readable results: `AUDIT_REPORT.json`, `VALIDATION.json`, `PRINT_LAYOUT_REPORT.csv`.

## Current onion-image update (2026-11-16)

The owner's supplied onion image fills the final two blank entries, PLUs 4093 and 94093, using exact WorkingList identity snapshots. Only the matching organic onion reuses this image, with the standard pink badge. All 546 previous embedded images remain unchanged; the regenerated workbook and PDF now have images for all 548 Produce entries. The original upload is preserved unchanged, with white-background JPEG library copies and transparent binder derivatives. Supplied-image provenance is recorded as unverified, without assuming ownership or AI/photographic origin. Technical checks pass; production publication remains blocked. Earlier coverage and blocker counts below are historical.

## Current missing-image update (2026-11-16)

The regenerated workbook and PDF fill 33 previously blank Produce entries using label-free crops from the owner's supplied sheet. Each mapping pins the exact WorkingList identity, originating source IDs and record fingerprint. All 513 previous embedded images remain unchanged; 546 entries now have images, with only PLUs 4093 and 94093 blank. Transparent binder copies and pink organic badges are preserved. The new illustrations retain unverified rights, photographic-origin and variety concerns; they do not grant production approval. Technical checks pass and publication remains blocked. Historical blocker counts below precede this update; current counts are in `PRODUCTION_BLOCKERS.csv`.

## Current weighed-guide update (2026-11-16)

The current binder has 30 pages, including 13 Produce pages. Page 3 now uses the owner-uploaded weighed-items guide with muted binder-green colors. All original instructions, rounding examples and label values were compared. At the owner's request, the displayed total is corrected from $134.31 to $134.32 ($35.07/kg x 3.830 kg = $134.3181, rounded to cents). The upload is retained unchanged; only the final price digit is changed in the label. Other label pixels and red warnings are unchanged in the derivative. The rendered page was inspected for readability and its complete border; automated print verification checks the full-resolution image pixels, placement and portrait Letter page. The build and post-build technical audit pass. Publication remains blocked.

The baseline, defects and page numbers below describe the earlier 27-page audit, not a new inspection of every current page.

## 1. Baseline (regenerated against the current repository, not copied from old reports)

* `PickList.xml` SHA-256 `38732869…8727` unchanged; original binder SHA-256 `6b0aeaa1…0f63` unchanged.
* 223 existing tests pass; the committed binder regenerates in about 60 s and reproduces its previous 27-page structure.
* Original workbook: 14 sheets, 104 package parts, 37 media parts.

## 2. Confirmed defects and fixes

| # | Area | Defect (confirmed visually in the old 27-page PDF) | Fix |
|---|------|------------------------------------------------------|-----|
| 1 | Bulk p.20 | "Coconut Shredded Unsweetened" and "Trail Mix Mountain Man Unsalted" were clipped (two wrapped lines in a one-line row). Cause: `ceil(len/29)` character estimate. | New `tools/text_fit.py` measures DejaVu Sans widths against the column as rendered; Bulk row heights now come from the measurement. |
| 2 | Bakery p.6 | Uneven row heights, broken pale-green striping, tiny effective text (about 7 pt after fit-to-page), mixed fonts. Cause: row heights shared between the two panels and sized by a 20-character estimate. | Uniform 24 pt rows, DejaVu Sans 12 pt, widths calibrated so no name wraps, one-page portrait Letter retained. |
| 3 | Produce p.8-17 | About 25 % of each landscape page was blank and photos were small. | Row height 25→31 pt and photo box 48×30→54×38 px (no upscaling beyond the source). Page count still 10. |
| 4 | Pipeline | No automated package, per-page PDF or blocker audit. | New `tools/audit_till_binder.py` plus tests (`test_text_fit.py`, `test_audit_till_binder.py`); CI workflow runs it after the build. |

## 3. Findings recorded but not changed

* Deli (p.7): a few rows are slightly taller than neighbours (owner-supplied Deli source). Content, 52 entries and 10 headings are exact; left as is.
* Pages 1-5, 24-27 are owner artwork or original scans, preserved byte-for-byte. Heading colours on the Weighed Items page are original.
* 154 embedded produce photos have an edge below 200 px (smallest 114 px). They come from approved legacy/booklet sources and are **not** upscaled.
* 70 image hashes are shared by more than one PLU (conventional/organic pairs and legacy reuse). Organic reuse of a matching conventional photo is allowed; the remaining reuse is in the open `PROHIBITED_IMAGE_REUSE` blocker.
* `Bulk (2)`-`Bulk (4)` are intentionally merged into one six-page Bulk sheet (400 entries verified by the build).
* The code-quality hotspots (2.6 k-line `build_till_binders.py`, duplicated theme/XML helpers) are documented only; refactoring the OOXML-preserving writer was deliberately avoided to protect package fidelity.

## 4. Data and package integrity

* 548 Produce records, 37 Bakery, 52 Deli (10 headings), 400 Bulk entries, exact PLU text preserved (checked by the build validators and printed-text comparison).
* Package: 0 parts removed, all 37 original media parts byte-identical (barcode artwork intact), 3 original Bulk sheets merged by design, workbook zip integrity OK. Changed parts are the themed sheets, styles, workbook and content types.
* Conversions: owner chart preserved; $4.99/lb = $11.00/kg (the $11.02 draft value is absent).
* Determinism: two consecutive builds produced identical image mappings and identical package contents (only zip timestamps differ).

## 5. Print validation

27 pages: Stretch 1, Conversions 1, Weighed 1, Soil/Floral 1, Misc 1, Bakery 1, Deli 1, Produce 10, Bulk 6, Delivery 2, Directory 2. Portrait Letter 612×792 for pages 1-7 and 24-25; landscape 792×612 for 8-23 and 26-27. No blank pages, no text in the margin guard. See `PRINT_LAYOUT_REPORT.csv` and `PDF_VISUAL_INSPECTION.md`.

## 6. Remaining blockers

See `PRODUCTION_BLOCKERS.csv`: 35 missing images, 35 image reviews after identity change, 323 prohibited-reuse flags, 5 classifications, 5 section placements, 248 WorkingList reconciliation records (0 decisions recorded), 90 organic placements. No owner decisions were invented, and the separate produce-booklet publication gate is unchanged.
