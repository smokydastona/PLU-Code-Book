# Cashier Till Binder

This repository builds and reviews one complete, photo-inclusive cashier binder. The editable workbook, sequential full-binder PDF and print-imposed half-fold booklet PDF are generated together; the booklet is a compact print format of the same binder, not a separate set of product content.

## Current output

Build artifacts are kept in `outputs/till-binder-review/`:

| File | Purpose |
| --- | --- |
| `Complete Till Binder - Updated - Photos.xlsx` | Editable complete binder |
| `Complete Till Binder - Updated - Photos.pdf` | Full binder in reading order |
| `Pocket Guide.pdf` | Imposed 2-up, duplex binder guide for folding and saddle-stitching |
| `OldTimers_Produce.pdf` | Large-print, text-only list: Fruits, then Vegetables ending with the Asian Produce group |
| `Sprinkler Room Map.pdf` | Standalone map with A/B divider and sprinkler-room locations; no emergency intercom codes |
| `BINDER_BOOKLET_IMPOSITION.csv` | Source-page placement for each physical sheet side |
| `VALIDATION.json`, `AUDIT_REPORT.json`, `PRODUCTION_BLOCKERS.csv`, `OWNER_REVIEW_FINDINGS.csv`, and CSV reports | Separate build, technical, data, image, rights and release evidence |

These are review outputs, not authorization for operational publication. Release blockers are mandatory unresolved reconciliation or product-identity decisions, missing exact-identity image selection, explicit image rejection/supersession, known usage restrictions, and required classification/section decisions. Low resolution, incomplete optional provenance, unverified rights documentation and duplicate fingerprints are reported separately as findings; they do not by themselves revoke an approval or block release. The original source records and the last approved production materials are not replaced by a review build.

The old `records_needing_review: 232` count was the number of distinct rows among 594 included catalogue records that had at least one pre-M1 blocker. It was not a count of unresolved PickList records: 219 rows had duplicate/reuse findings, 26 had image-quality, rights or provenance findings, and 13 rows appeared in both groups (`219 + 26 - 13 = 232`). `PICKLIST_SYNC.json` now records the recomputed component sets and documents each counter. The separate count of 50 merged/standardized WorkingList identities records applied owner decisions, not unresolved decisions or image findings. Current explicit rejection findings still block release and are listed in `PRODUCTION_BLOCKERS.csv`.

## Build and audit

Use Python 3.12, the pinned packages in `requirements.txt`, LibreOffice Calc and Writer, and DejaVu Sans fonts. From the repository root:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
sudo apt-get install -y libreoffice-calc libreoffice-writer fonts-dejavu-core fonts-dejavu-extra

.venv/bin/python tools/build_till_binders.py \
  --binder sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx \
  --output /path/to/empty/till-binder-review \
  --verify-print
.venv/bin/python tools/audit_till_binder.py \
  --binder sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx \
  --output /path/to/that/till-binder-review
```

The output directory must be empty. `--verify-print` exports and checks the complete workbook, then creates the imposed booklet PDF and its page map. The audit checks workbook package preservation, printed page count and geometry, image coverage, booklet imposition, and open production blockers. Technical success never clears owner-review gates or approves a binder for cashier use.

Regenerate the standalone sprinkler map separately with
`.venv/bin/python tools/build_sprinkler_map.py`. It uses the unchanged original
directory artwork and the owner's markings in
`sources/binders/Sprinkler_Room_Map_Reference.pdf`: side A is left of the divider,
side B is right, room A is at the front-left wall, and room B is outside the
upper-right angled wall. This replaces the former `Store_Directory_Map.pdf`
deliverable; the binder's Store Directory and emergency intercom panel are unchanged.

## Management app

The read-only browser app at `management-app/` loads the generated catalogue,
WorkingList reconciliation and production readiness reports. See
`management-app/README.md` for local serving, Node.js tests, GitHub Pages
deployment and the concrete server-side authentication prerequisite for any
future write feature. The app does not simulate approvals or writes.

## Review packages and production publication

Run the **Package and publish PLU binder** workflow manually with `review` to
produce a complete ZIP review artifact, or `production` to request a
production candidate. The package includes the editable workbook, complete
PDF, imposed booklet, OldTimers PDF, validation and audit reports, reconciliation
evidence, and `RELEASE_MANIFEST.json` with input/output checksums, generator
identity, dependency versions, available approval-basis metadata and distinct
gate results. A review build is retained as a workflow artifact even if a
mandatory gate fails.

Production packaging fails closed on mandatory readiness blockers or failed
technical/print validation. Quality, rights-documentation, provenance and
duplicate/reuse findings remain separately reported and are not converted into
approval revocations. A passing automated gate still requires approval through
the GitHub `production` environment and a run from the repository default
branch; configure required reviewers for that environment before enabling
publication. Publication creates a new
commit-tagged release and never updates or deletes an existing release. The
packager refuses to overwrite an existing ZIP, and a blocked run leaves the
last approved production release intact.

The owner-supplied sheets are retained for review; suitable crops are references, not production image approvals. Newer suitable crops take precedence over older sheet references only for the same exact WorkingList identity; superseded and unselected panels remain in the reference audit. Invalid approved assignments may fall back to a review reference without clearing their validation errors or production gates.

Run `python tools/build_temporary_produce_review.py` to refresh the single
`outputs/review-booklets/Produce_Image_Review_Temporary.pdf` from the recorded
image screening decisions. It excludes verified prior exact-image approvals,
uses the original screened candidate checksums even when library paths have
changed, and blocks regeneration if those bytes or product identities are
unavailable. Replacement-required images have no approval checkbox. The
adjacent `TEMPORARY_PRODUCE_IMAGE_REVIEW.csv` and JSON record the current pages,
candidate paths, checksums and prior approval exclusions. Only older
`Produce_Image_Review_Temporary*.pdf` copies in that output directory are
removed; checked owner uploads, other review booklets and published binders
remain unchanged.

`data/review-references/lettuce_image_exact.pdf` is a lettuce-identification guide only. Its sidecar records the source checksum and reference-only authorization; the labelled chart is not assigned to cashier entries and grants no image approvals.

The individual `santaclaus-1.jpg` and `Romaine.jpg` references supersede their earlier sheet crops for PLUs 4336 and 4640. Their manifests record owner-confirmed permission; native sizes of 480 x 352 and 488 x 488 remain below the production resolution requirement.

## Half-fold booklet printing

The imposed PDF arranges the complete binder in saddle-stitch order: the first and last pages share the outside sheet, with inner pages paired toward the center fold. It includes every page in the full binder, including wide landscape sections. Pages are fitted proportionally into portrait half-page panels on landscape US Letter sheets. Produce, Bulk and Store Directory pages are rotated 90° counterclockwise within their panels to make better use of the panel area; turn the guide sideways to read those sections. If the binder page count is not divisible by four, blank panels are added to complete the final signature; the page map records those blanks and rotations.

Print the imposed PDF on US Letter paper in landscape orientation, 2-up duplex, flip on the short edge, at 100% scale. Fold each sheet once down the center and nest the sheets in order; saddle-stitch through the center fold. Wide pages are reduced to fit their half-page panels, so verify their legibility on a physical proof before printing a full run. Do not use the printer driver's booklet/imposition feature on this already-imposed PDF.

## Content and source authority

- `sources/plu/PickList.xml` is the authoritative imported record source. Only @smokydastona may update it. Scripts must report XML errors, never repair or normalize it, and preserve exact PLU strings, including leading zeros.
- `sources/plu/WorkingList.xml` is regenerated from the PickList and reconciled with owner-approved decisions. Preserve original XML records and identities in audit reports. Apply decisions only when their original identities still match; flag changed, new or ambiguous records for review.
- Cashier sheets are generated from the reviewed WorkingList and owner-maintained binder sources. Product-name decisions do not approve images; changed identities require fresh image review. Keep provenance and approval metadata in audit evidence, not cashier-facing entries.
- Production approval is identity- and asset-specific. Explicit rejection or supersession overrides prior approval; replacing an approved asset or changing its product identity requires fresh review. Keep production approval, image quality, rights/permission status, provenance completeness and duplicate/reuse review as separate findings.
- Block production publication for unresolved mandatory reconciliation, identity, image-selection, known rights/permission, classification or section decisions. A duplicate fingerprint alone, a low-resolution warning, or missing optional metadata does not revoke approval or block release. Keep the last published production materials intact while blocked.

The build keeps the original complete workbook in `sources/binders/` unchanged. Produce, Bakery, Deli and Bulk updates preserve source values and print layouts; PLUs appear to the left of product descriptions. Standalone “Organic” in product names displays as “Org.” while the “Organics” heading and organic image badges remain unchanged. Bulk exclusions and source identities are recorded separately from cashier content.

Misc Codes keeps its editable entries and adds a clean No Purchase Entry panel below Floral Codes on the same page. The heading, instructions, and barcode are generated without using pixels from the scan. The build preserves `sources/binders/Misc_Codes_No_Purchase_Upload.jpg` as source evidence and verifies that the scan, generated barcode, and rendered binder page decode to `0770974140186`. Check a physical print with the till scanner before use.

On the Helpful Conversion page, the build enlarges the Weight Conversions table by 8% without changing its values or the other tables. The original upload is preserved.

The Store Directory map includes a compact Emergency Intercom Pages panel for Code Adam, Cashier 20, and Code Zero above the original map markings.

The Weighed more $100 section prints the weighed-items guide followed by `sources/binders/Rain_Check_Instructions_Updated.png` as its second page. `tools/prepare_raincheck_image.py` deterministically prepares this corrected 1024 x 1536 PNG from the preserved owner upload after verifying its checksum. It records the corrections to the date format, catalogue-matched example product and PLU, missing store fields, and corner text: the unsupported accuracy claim is removed and the remaining instruction reads “Follow these steps to complete the raincheck/voucher correctly.” The original upload and `sources/binders/Rain_Check.pdf` remain unchanged and checksummed as comparison evidence. The poster is centered horizontally and vertically on its US Letter page without distorting its aspect ratio; any changed input requires review before rebuilding.

The OldTimers list is generated with every displayed binder produce entry, without photos. All text is at least 12 points in a two-column layout. Each produce group stays on one page, using both columns if needed; a group that will not fit in the remaining space starts on the next page. A group too large for a full page blocks generation rather than shrinking text or splitting across pages. Fruits come first, then Vegetables, with Asian Produce as the last group within Vegetables, not a separate section or a forced new page. That group includes all entries from the binder's Asian Produce section. It uses the binder's existing fruit/vegetable classifications and keeps melons together under fruit, in the style of the original `sources/binders/OldTimers_Produce_codes.pdf`. Unclassified products follow the binder's review classification; this presentation does not approve their classifications or release the list for operational use.

## Repository map

| Path | Contents |
| --- | --- |
| `sources/binders/` | Original binder template and owner-reviewed section sources |
| `sources/plu/` | Authoritative PickList and generated, reviewed WorkingList |
| `sources/image-library/` | Local image assets and separate approval/provenance indexes |
| `data/` | Identity-pinned decisions, image assignments and review metadata |
| `tools/build_till_binders.py` | Complete binder and print-output builder |
| `tools/audit_till_binder.py` | Post-build technical and production-readiness audit |
| `outputs/till-binder-review/` | Latest generated review binder and evidence |

The single `sources/image-library/` directory holds approved WorkingList images,
owner candidates and separately indexed legacy and review copies. Run
`python tools/build_image_library.py --check` and
`python tools/used_image_library.py --check` to verify its approved inventory
and cached assets. `python tools/used_image_library.py` refreshes the binder
and legacy associations without changing source images or granting approval.
Explicitly rejected or superseded images remain in the audit history, are not
offered as approved choices, and are left blank in newly built binders.

Run `python tools/consolidate_image_library.py` to store identical JPEG bytes
once. `IMAGE_ALIASES.json` resolves logical filenames to their stored images;
each original `.jpg.json` sidecar and approval index remains separate and
unchanged. A logical filename in an index may therefore have no standalone
JPEG: use `tools.library_assets.resolve_image` to find its physical asset.
Binder builds and local image lookup resolve aliases automatically. Refreshes
retain consolidation; images with different bytes (including organic badges)
are never merged, and sharing storage does not approve product-image reuse.
`outputs/reports/IMAGE_LIBRARY_ALIASES.csv` and
`IMAGE_LIBRARY_CONSOLIDATION.json` record the storage savings and metadata
preservation.

The owner-uploaded `data/review-references/2011-PLU-Listing1.pdf` is a historical
reference for future image-library coverage. Its checksummed sidecar records
reference-only status. Run `python tools/build_future_plu_coverage.py` to generate
`outputs/reports/FUTURE_PLU_IMAGE_COVERAGE.csv` and its JSON summary.
`PLU_REFERENCE_SOURCE_TEXT.json` preserves the extracted page text for audit;
each coverage row retains its original text, page and code group. Superscript
footnotes are excluded from codes, not mistaken for extra PLU digits. Retailer
assignments, unspecified commodities and unlisted ranges are retained as
review placeholders, never expanded into invented product identities.
Same-code library images are only search leads: every future product identity
and image needs owner review. This 2011 document is not an exhaustive current
PLU catalogue. Coverage generation does not update PickList or WorkingList,
assign images, or grant image approvals.

The `Fetch unreviewed produce candidates` Action has an `external-leads` mode for the identity-pinned URLs in `data/review-references/external-image-source-leads.json`. It reads page metadata only; it does not download image previews or PDF bodies, verify license claims, approve images, or change assignments. Search-result URLs are recorded without querying them. Fetchers retry DNS lookup failures up to three times with bounded backoff and report persistent failures as `DNS_ERROR`; HTTP errors such as 403 are not retried or bypassed. The Action runs on a GitHub-hosted runner, so it does not depend on the local machine's DNS configuration.
