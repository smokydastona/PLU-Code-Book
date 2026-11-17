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

The `Fetch unreviewed produce candidates` Action has an `external-leads` mode for the identity-pinned URLs in `data/review-references/external-image-source-leads.json`. It reads page metadata only; it does not download image previews or PDF bodies, verify license claims, approve images, or change assignments. Search-result URLs are recorded without querying them. Fetchers retry DNS lookup failures up to three times with bounded backoff and report persistent failures as `DNS_ERROR`; HTTP errors such as 403 are not retried or bypassed. The Action runs on a GitHub-hosted runner, so it does not depend on the local machine's DNS configuration.
