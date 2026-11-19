"""Build the standalone map using the owner's marked Sprinkler_Room_Map_Reference.pdf."""
import argparse
from pathlib import Path

import pymupdf

from build_till_binders import DIRECTORY_SOURCES, ROOT, themed_directory_image
from binder_theme import DARK_GREEN

OUTPUT_NAME = "Sprinkler Room Map.pdf"
OUTLINE = tuple(int(DARK_GREEN[i:i + 2], 16) / 255 for i in (0, 2, 4))
LABEL_FILL = tuple(int("DCE6DA"[i:i + 2], 16) / 255 for i in (0, 2, 4))
SPRINKLER_FILL = tuple(int("DDEBF7"[i:i + 2], 16) / 255 for i in (0, 2, 4))
SCALE = 744 / 1650
LEFT, TOP = 24, (612 - 1275 * SCALE) / 2


def point(x, y):
    return pymupdf.Point(LEFT + x * SCALE, TOP + y * SCALE)


def rectangle(box):
    return pymupdf.Rect(point(*box[:2]), point(*box[2:]))


def label(page, box, text, size, border=False, fill=LABEL_FILL):
    area = rectangle(box)
    if border:
        shape = page.new_shape()
        shape.draw_rect(area, radius=0.15)
        shape.finish(color=OUTLINE, fill=fill, width=0.8)
        shape.commit()
    if page.insert_textbox(area, text, fontsize=size * SCALE, fontname="hebo",
                           color=(0, 0, 0), align=pymupdf.TEXT_ALIGN_CENTER) < 0:
        raise ValueError(f"Sprinkler map label does not fit: {text}")


def generate(path, root=ROOT):
    # Start before the binder-only intercom panel is added, never erase map pixels.
    artwork = themed_directory_image((root / DIRECTORY_SOURCES[1]).read_bytes(), 1)
    with pymupdf.open() as document:
        page = document.new_page(width=792, height=612)
        page.insert_image(rectangle((0, 0, 1650, 1275)), stream=artwork)
        label(page, (20, 15, 1630, 78), "Sprinkler Room Map", 36)
        label(page, (590, 84, 750, 133), "Side A", 28, border=True)
        label(page, (960, 84, 1120, 133), "Side B", 28, border=True)
        page.draw_line(point(940, 190), point(940, 1150), color=OUTLINE,
                       width=3, dashes="[8 4] 0")

        page.draw_line(point(220, 1160), point(225, 1075), color=OUTLINE, width=1.2)
        page.draw_circle(point(225, 1075), 3, color=OUTLINE, width=1.2)
        label(page, (15, 1160, 320, 1216), "Sprinkler Room A", 32,
              border=True, fill=SPRINKLER_FILL)

        page.draw_line(point(1480, 191), point(1435, 298), color=OUTLINE, width=1.2)
        page.draw_circle(point(1435, 298), 3, color=OUTLINE, width=1.2)
        label(page, (1340, 85, 1640, 191), "Sprinkler Room B\n(Exterior of building)",
              29, border=True, fill=SPRINKLER_FILL)
        document.set_metadata({"title": "Sprinkler Room Map"})
        document.save(path, garbage=4, deflate=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path,
                        default=ROOT / "outputs/till-binder-review" / OUTPUT_NAME)
    args = parser.parse_args()
    generate(args.output)
    print(args.output)


if __name__ == "__main__":
    main()
