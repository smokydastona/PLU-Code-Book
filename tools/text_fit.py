"""Measure wrapped text so worksheet row heights always fit the rendered font."""
from functools import lru_cache

from PIL import ImageFont

FONT_FILES = {False: "DejaVuSans.ttf", True: "DejaVuSans-Bold.ttf"}
PADDING_POINTS = 8
DIGIT_PIXELS = 9.33


def column_points(width_characters):
    # Column widths count default-font digits. The print engine renders the workbook's
    # default font as DejaVu Sans 11 pt (9.33 px per digit), plus 5 px of padding.
    return (width_characters * DIGIT_PIXELS + 5) * 0.75 - PADDING_POINTS


@lru_cache(maxsize=None)
def _font(bold, points):
    # Measured at 10x for sub-point precision; DejaVu Sans is the binder's print font.
    return ImageFont.truetype(FONT_FILES[bold], round(points * 10))


def text_points(text, points, bold=False):
    return _font(bold, points).getlength(text) / 10


def line_count(text, points, width_characters, bold=False):
    """Greedy word-wrap line count for ``text`` in a column of the given Excel width."""
    limit = column_points(width_characters)
    lines, current = 1, ""
    for word in text.split():
        candidate = f"{current} {word}".strip()
        if current and text_points(candidate, points, bold) > limit:
            lines += 1
            current = word
        else:
            current = candidate
    return lines
