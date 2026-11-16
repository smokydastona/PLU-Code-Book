# PDF Page-by-Page Visual Inspection

Editable-sheet heading update (2026-11-16): Bulk headings are centered, and all 12 continuing Bulk categories print their full non-bold, half-size `(continued)` label directly after the centered name. The longest category is slightly reduced in size so its label fits without clipping. Bakery, Deli and Misc Codes headings are centered in their existing cells; Soil and Floral headings remain centered. The owner-approved image-based artwork is unchanged. The updated Bulk first and final pages were inspected, and automated checks measured every continuation label.

Produce heading update (2026-11-16): the first Produce page was rendered and inspected. Letter and subgroup names remain centered. The 17 repeated subgroup headings have a native-text, non-bold `(continued)` overlay immediately beside the name, separated by one normal word space rather than a separate cell. Declared font sizes are exactly half the title sizes. All printed labels were checked for inline word gaps, regular weight and half-size rendering within Calc font quantization. The Apples heading centers are unchanged across the three column blocks.

Current update (2026-11-16): page 3 of the current 30-page binder was rendered and inspected. The uploaded weighed-items guide has binder-green headings, unchanged red warnings, legible instructions and examples, and a complete border within a single portrait Letter page. Its contents were compared with the original worksheet and embedded label. At the owner's request, only the final price digit is changed to correct the total from $134.31 to $134.32; all other label pixels are unchanged. The complete binder also passed automated print and package checks. The table below records the earlier 27-page inspection; its later page ranges are historical, not current.

Every page of the regenerated 27-page PDF was rendered (70-110 dpi contact sheets; pages with defects at 200 dpi) and inspected by eye. Automated checks (`PRINT_LAYOUT_REPORT.csv`) are supplementary, not a substitute.

| Page | Section | Result |
|---|---|---|
| 1 | Time To Stretch | OK. Owner colour artwork, correct aspect ratio. |
| 2 | Conversions | OK. Owner chart; $4.99/lb → 11.00/kg. |
| 3 | Weighed items over $100 | OK. Original content, readable. |
| 4 | Soil / Floral Codes | OK. Barcodes intact; Floral Codes beside Soil, striped white/pale green, no overlap. |
| 5 | Misc Codes | OK. Green theme, categories clear, one page. |
| 6 | Bakery | **Fixed.** Uniform rows, consistent stripes, larger text; 37 entries, one page. |
| 7 | Deli | OK, minor row-height unevenness from the owner source (recorded). |
| 8-17 | Produce | OK. Three blocks, A-Z, red shared-PLU highlight, pink ORGANIC badges, missing images blank, no truncated names. Larger rows and photos; bottom whitespace reduced. |
| 18 | Bulk (Candy) | OK. |
| 19 | Bulk (Chocolate/Snacking) | OK. `DC` entries kept; wrapped names fit their merged rows. |
| 20 | Bulk (Baking/Nut Mixes) | **Fixed.** Previously clipped names now fit. |
| 21-23 | Bulk | OK. 11 pt, 100 %, repeated category headings at column breaks. |
| 24-25 | Delivery Charges | OK. Both original pages and barcodes, consecutive. |
| 26-27 | Store Directory | OK. Index and map; exits, muster-point text, Front Parking Lot (green) and Back Parking Lot (red) labels unchanged. |

Residual notes: Produce and Bulk columns end with some bottom whitespace by design (fixed 20-line blocks); no clipping, overlap or illegible text found.
Grayscale: dark-green headers and pale-green stripes keep strong luminance contrast; red shared-PLU highlighting is also bold-weighted.
