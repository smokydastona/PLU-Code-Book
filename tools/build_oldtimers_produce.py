"""Generate a large-print, text-only produce list from the binder display rows."""
from collections import Counter

import pymupdf

from produce_order import ASIAN_SECTION

OUTPUT_NAME = "OldTimers_Produce.pdf"
SOURCE = "sources/binders/OldTimers_Produce_codes.pdf"
SOURCE_SHA256 = "6846210817a7f7bffe84bb5f30db75ef1276b203fda53c0256cb77cf037e0672"
FRUIT_FAMILIES = {"CANTALOUPE", "HONEYDEW", "MELON", "WATERMELON"}
PAGE_WIDTH, PAGE_HEIGHT = 612, 792
LEFT, RIGHT, TOP, BOTTOM = 36, 36, 78, 40
COLUMN_GAP = 12
COLUMNS = 2
COLUMN_WIDTH = (PAGE_WIDTH - LEFT - RIGHT - (COLUMNS - 1) * COLUMN_GAP) / COLUMNS
CODE_WIDTH = 64
FONT_SIZE = 12
LINE_HEIGHT = 16
HEADING_HEIGHT = 22
REGULAR_FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
BOLD_FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
REGULAR = pymupdf.Font(fontfile=REGULAR_FONT)
BOLD = pymupdf.Font(fontfile=BOLD_FONT)


def section(row):
    if row["display_section"] == ASIAN_SECTION:
        return ASIAN_SECTION
    if row["sort_product"] in FRUIT_FAMILIES:
        return "FRUITS"
    if row["book"] == "Fruit":
        return "FRUITS"
    if row["book"] == "Vegetable":
        return "VEGETABLES"
    raise ValueError(f"{row['source_id']}: invalid produce classification {row['book']!r}")


def wrap_name(text, width):
    words = text.split()
    if not words:
        raise ValueError("Produce name is blank")
    lines, line = [], ""
    for word in words:
        candidate = f"{line} {word}".strip()
        if REGULAR.text_length(candidate, fontsize=FONT_SIZE) <= width:
            line = candidate
        elif line:
            lines.append(line)
            line = word
        else:
            raise ValueError(f"Produce name contains a word too wide to print: {word!r}")
        if REGULAR.text_length(line, fontsize=FONT_SIZE) > width:
            raise ValueError(f"Produce name contains a word too wide to print: {word!r}")
    lines.append(line)
    return lines


def generate(path, rows):
    """Write all and only the displayed binder rows, in the legacy category split."""
    groups = {"FRUITS": {}, "VEGETABLES": {}, ASIAN_SECTION: {}}
    for row in rows:
        code = row["plu"]
        if not isinstance(code, str) or not code or not code.isdigit():
            raise ValueError(f"{row['source_id']}: PLU must be a non-empty digit string")
        if REGULAR.text_length(code, fontsize=FONT_SIZE) > CODE_WIDTH - 4:
            raise ValueError(f"{row['source_id']}: PLU is too wide to print")
        groups[section(row)].setdefault(row["sort_product"], []).append(row)
    if not rows:
        raise ValueError("No produce records for OldTimers list")

    document = pymupdf.open()
    placements = []
    try:
        for title, categories in groups.items():
            if not categories:
                continue
            page = None
            column = COLUMNS - 1
            y = 0

            def advance():
                nonlocal page, column, y
                column += 1
                if column == COLUMNS:
                    page = document.new_page(width=PAGE_WIDTH, height=PAGE_HEIGHT)
                    page.insert_font(fontname="regular", fontfile=REGULAR_FONT)
                    page.insert_font(fontname="bold", fontfile=BOLD_FONT)
                    page.insert_text((LEFT, 35), "PRODUCE DEPARTMENT CODES (DEPT #2)",
                                     fontname="bold", fontsize=13)
                    page.insert_text((LEFT, 58), title, fontname="bold", fontsize=12)
                    column = 0
                y = TOP

            def x_position():
                return LEFT + column * (COLUMN_WIDTH + COLUMN_GAP)

            for category, entries in sorted(categories.items()):
                entries.sort(key=lambda row: (row["sort_variety"], row["organic"],
                                              row["display_name"], row["plu"], row["source_id"]))
                show_heading = True
                for row in entries:
                    lines = wrap_name(row["display_name"], COLUMN_WIDTH - CODE_WIDTH)
                    height = len(lines) * LINE_HEIGHT
                    if page is None or y + height + (HEADING_HEIGHT if show_heading else 0) > PAGE_HEIGHT - BOTTOM:
                        advance()
                        show_heading = True
                    if y + height + (HEADING_HEIGHT if show_heading else 0) > PAGE_HEIGHT - BOTTOM:
                        raise ValueError(f"{row['source_id']}: produce entry is too tall to print")
                    x = x_position()
                    if show_heading:
                        heading = category + (" (CONT.)" if row is not entries[0] else "")
                        if BOLD.text_length(heading, fontsize=FONT_SIZE) > COLUMN_WIDTH:
                            raise ValueError(f"Produce heading is too wide to print: {heading!r}")
                        page.insert_text((x, y), heading, fontname="bold", fontsize=FONT_SIZE)
                        y += HEADING_HEIGHT
                        show_heading = False
                    page.insert_text((x, y), row["plu"], fontname="regular", fontsize=FONT_SIZE)
                    for index, line in enumerate(lines):
                        page.insert_text((x + CODE_WIDTH, y + index * LINE_HEIGHT), line,
                                         fontname="regular", fontsize=FONT_SIZE)
                    placements.append(dict(source_id=row["source_id"], plu=row["plu"],
                                           product=row["display_name"], section=title,
                                           page=page.number + 1, column=column + 1,
                                           x=x, y=y, lines=len(lines)))
                    y += height
                y += 6
        document.set_metadata({"title": "OldTimers Produce Codes"})
        document.save(path, garbage=4, deflate=True)
    finally:
        document.close()
    validate(path, rows, placements)
    return dict(file=OUTPUT_NAME, records=len(placements),
                fruit_records=sum(row["section"] == "FRUITS" for row in placements),
                vegetable_records=sum(row["section"] == "VEGETABLES" for row in placements),
                asian_records=sum(row["section"] == ASIAN_SECTION for row in placements),
                pages=max(row["page"] for row in placements), image_free=True,
                minimum_font_points=FONT_SIZE)


def validate(path, rows, placements):
    expected = Counter((row["source_id"], row["plu"], row["display_name"], section(row))
                       for row in rows)
    actual = Counter((row["source_id"], row["plu"], row["product"], row["section"])
                     for row in placements)
    if actual != expected:
        raise ValueError("OldTimers list does not match the binder produce entries")
    with pymupdf.open(path) as document:
        for page in document:
            if page.rect != pymupdf.Rect(0, 0, PAGE_WIDTH, PAGE_HEIGHT) or page.get_images():
                raise ValueError("OldTimers page must be image-free portrait US Letter")
            for block in page.get_text("dict")["blocks"]:
                for line in block.get("lines", []):
                    if any(span["size"] < 12 for span in line["spans"]):
                        raise ValueError("OldTimers text must be at least 12 points")
            for word in page.get_text("words"):
                if not pymupdf.Rect(LEFT - 1, 0, PAGE_WIDTH - RIGHT + 1, PAGE_HEIGHT - BOTTOM + 4).contains(
                        pymupdf.Rect(word[:4])):
                    raise ValueError("OldTimers text is outside the printable area")
        for row in placements:
            page = document[row["page"] - 1]
            x, y = row["x"], row["y"]
            words = [word for word in page.get_text("words")
                     if x - 1 <= word[0] < x + COLUMN_WIDTH and
                     y - 8 <= (word[1] + word[3]) / 2 < y + (row["lines"] - 1) * LINE_HEIGHT + 3]
            code = " ".join(word[4] for word in words if word[0] < x + CODE_WIDTH - 1)
            name = " ".join(word[4] for word in words if word[0] >= x + CODE_WIDTH - 1)
            if code != row["plu"] or name != " ".join(row["product"].split()):
                raise ValueError(f"OldTimers page {row['page']} has missing or altered produce text: "
                                 f"expected {row['plu']!r} {row['product']!r}, got {code!r} {name!r}")
