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
| `VALIDATION.json`, `AUDIT_REPORT.json`, and CSV reports | Build, print-layout, image and production-readiness evidence |

These are review outputs, not authorization for operational publication. Production remains blocked until WorkingList reconciliation and all applicable image, identity, rights, classification and reuse reviews pass. The original source records and the last approved production materials are not replaced by a review build.

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

The owner-supplied sheets are retained for review; suitable crops are references, not production image approvals. Newer suitable crops take precedence over older sheet references only for the same exact WorkingList identity; superseded and unselected panels remain in the reference audit. Invalid approved assignments may fall back to a review reference without clearing their validation errors or production gates.

`data/review-references/lettuce_image_exact.pdf` is a lettuce-identification guide only. Its sidecar records the source checksum and reference-only authorization; the labelled chart is not assigned to cashier entries and grants no image approvals.

## Half-fold booklet printing

The imposed PDF arranges the complete binder in saddle-stitch order: the first and last pages share the outside sheet, with inner pages paired toward the center fold. It includes every page in the full binder, including wide landscape sections. Pages are fitted proportionally into portrait half-page panels on landscape US Letter sheets. Produce, Bulk and Store Directory pages are rotated 90° counterclockwise within their panels to make better use of the panel area; turn the guide sideways to read those sections. If the binder page count is not divisible by four, blank panels are added to complete the final signature; the page map records those blanks and rotations.

Print the imposed PDF on US Letter paper in landscape orientation, 2-up duplex, flip on the short edge, at 100% scale. Fold each sheet once down the center and nest the sheets in order; saddle-stitch through the center fold. Wide pages are reduced to fit their half-page panels, so verify their legibility on a physical proof before printing a full run. Do not use the printer driver's booklet/imposition feature on this already-imposed PDF.

## Content and source authority

- `sources/plu/PickList.xml` is the authoritative imported record source. Only @smokydastona may update it. Scripts must report XML errors, never repair or normalize it, and preserve exact PLU strings, including leading zeros.
- `sources/plu/WorkingList.xml` is regenerated from the PickList and reconciled with owner-approved decisions. Preserve original XML records and identities in audit reports. Apply decisions only when their original identities still match; flag changed, new or ambiguous records for review.
- Cashier sheets are generated from the reviewed WorkingList and owner-maintained binder sources. Product-name decisions do not approve images; changed identities require fresh image review. Keep provenance and approval metadata in audit evidence, not cashier-facing entries.
- Block production publication when reconciliation, identity, image, rights or reuse gates fail. Keep the last published production materials intact while blocked.

The build keeps the original complete workbook in `sources/binders/` unchanged. Produce, Bakery, Deli and Bulk updates preserve source values and print layouts; PLUs appear to the left of product descriptions. Standalone “Organic” in product names displays as “Org.” while the “Organics” heading and organic image badges remain unchanged. Bulk exclusions and source identities are recorded separately from cashier content.

The Weighed more $100 section prints the weighed-items guide followed by the owner-supplied `sources/binders/Rain_Check.pdf` as its second page. The PDF is retained unchanged and rendered into the workbook; changes to it require review before rebuilding.

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

The `Fetch unreviewed produce candidates` Action has an `external-leads` mode for the identity-pinned URLs in `data/review-references/external-image-source-leads.json`. It reads page metadata only; it does not download image previews or PDF bodies, verify license claims, approve images, or change assignments. Search-result URLs are recorded without querying them.
