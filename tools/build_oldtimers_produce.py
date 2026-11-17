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
        return "VEGETABLES"
    if row["sort_product"] in FRUIT_FAMILIES:
        return "FRUITS"
    if row["book"] == "Fruit":
        return "FRUITS"
    if row["book"] == "Vegetable":
        return "VEGETABLES"
    raise ValueError(f"{row['source_id']}: invalid produce classification {row['book']!r}")


def group(row):
    return ASIAN_SECTION if row["display_section"] == ASIAN_SECTION else row["sort_product"]


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


def group_layout(entries, column, y):
    """Fit a whole group into the remaining columns of one page."""
    layout = []
    show_heading = True
    for row, lines in entries:
        height = len(lines) * LINE_HEIGHT
        if y + height + (HEADING_HEIGHT if show_heading else 0) > PAGE_HEIGHT - BOTTOM:
            column += 1
            y = TOP
            show_heading = True
        if column >= COLUMNS or y + height + (HEADING_HEIGHT if show_heading else 0) > PAGE_HEIGHT - BOTTOM:
            return None
        if show_heading:
            y += HEADING_HEIGHT
        layout.append((row, lines, column, y, show_heading))
        y += height
        show_heading = False
    return layout


def generate(path, rows):
    """Write all and only the displayed binder rows, in the legacy category split."""
    groups = {"FRUITS": {}, "VEGETABLES": {}}
    for row in rows:
        code = row["plu"]
        if not isinstance(code, str) or not code or not code.isdigit():
            raise ValueError(f"{row['source_id']}: PLU must be a non-empty digit string")
        if REGULAR.text_length(code, fontsize=FONT_SIZE) > CODE_WIDTH - 4:
            raise ValueError(f"{row['source_id']}: PLU is too wide to print")
        groups[section(row)].setdefault(group(row), []).append(row)
    if not rows:
        raise ValueError("No produce records for OldTimers list")

    document = pymupdf.open()
    placements = []
    try:
        for title, categories in groups.items():
            if not categories:
                continue
            page = None
            column = 0
            y = TOP

            def new_page():
                nonlocal page, column, y
                page = document.new_page(width=PAGE_WIDTH, height=PAGE_HEIGHT)
                page.insert_font(fontname="regular", fontfile=REGULAR_FONT)
                page.insert_font(fontname="bold", fontfile=BOLD_FONT)
                page.insert_text((LEFT, 35), "PRODUCE DEPARTMENT CODES (DEPT #2)",
                                 fontname="bold", fontsize=13)
                page.insert_text((LEFT, 58), title, fontname="bold", fontsize=12)
                column = 0
                y = TOP

            for category, entries in sorted(categories.items(),
                                            key=lambda item: (item[0] == ASIAN_SECTION, item[0])):
                entries.sort(key=lambda row: (row["sort_product"], row["sort_variety"], row["organic"],
                                              row["display_name"], row["plu"], row["source_id"]))
                prepared = [(row, wrap_name(row["display_name"], COLUMN_WIDTH - CODE_WIDTH))
                            for row in entries]
                if page is None:
                    new_page()
                layout = group_layout(prepared, column, y)
                if layout is None:
                    if column != 0 or y != TOP:
                        new_page()
                    layout = group_layout(prepared, column, y)
                if layout is None:
                    raise ValueError(f"Produce group {category!r} cannot fit on one page at 12 points")
                for row, lines, column, y, show_heading in layout:
                    x = LEFT + column * (COLUMN_WIDTH + COLUMN_GAP)
                    if show_heading:
                        heading = category + (" (CONT.)" if row is not entries[0] else "")
                        if BOLD.text_length(heading, fontsize=FONT_SIZE) > COLUMN_WIDTH:
                            raise ValueError(f"Produce heading is too wide to print: {heading!r}")
                        page.insert_text((x, y - HEADING_HEIGHT), heading, fontname="bold",
                                         fontsize=FONT_SIZE)
                    page.insert_text((x, y), row["plu"], fontname="regular", fontsize=FONT_SIZE)
                    for index, line in enumerate(lines):
                        page.insert_text((x + CODE_WIDTH, y + index * LINE_HEIGHT), line,
                                         fontname="regular", fontsize=FONT_SIZE)
                    placements.append(dict(source_id=row["source_id"], plu=row["plu"],
                                           product=row["display_name"], section=title, group=category,
                                           page=page.number + 1, column=column + 1,
                                           x=x, y=y, lines=len(lines)))
                    y += len(lines) * LINE_HEIGHT
                y += 6
        document.set_metadata({"title": "OldTimers Produce Codes"})
        document.save(path, garbage=4, deflate=True)
    finally:
        document.close()
    validate(path, rows, placements)
    return dict(file=OUTPUT_NAME, records=len(placements),
                fruit_records=sum(row["section"] == "FRUITS" for row in placements),
                vegetable_records=sum(row["section"] == "VEGETABLES" for row in placements),
                asian_records=sum(row["group"] == ASIAN_SECTION for row in placements),
                pages=max(row["page"] for row in placements), image_free=True,
                minimum_font_points=FONT_SIZE, groups_kept_on_one_page=True)


def validate(path, rows, placements):
    expected = Counter((row["source_id"], row["plu"], row["display_name"], section(row), group(row))
                       for row in rows)
    actual = Counter((row["source_id"], row["plu"], row["product"], row["section"], row["group"])
                     for row in placements)
    if actual != expected:
        raise ValueError("OldTimers list does not match the binder produce entries")
    group_pages = {}
    for row in placements:
        key = (row["section"], row["group"])
        if group_pages.setdefault(key, row["page"]) != row["page"]:
            raise ValueError(f"OldTimers produce group {key[1]!r} is split across pages")
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
