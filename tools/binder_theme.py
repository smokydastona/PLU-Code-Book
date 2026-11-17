"""Shared colors for the complete binder's editable tables."""
import re

from openpyxl.styles import PatternFill

DARK_GREEN = "305040"
PALE_GREEN = "DDF2DF"
WHITE = "FFFFFF"


def row_fill(index):
    return PatternFill("solid", fgColor=PALE_GREEN if index % 2 else WHITE)


def display_product_name(name):
    return re.sub(r"\borganic\b", "Org.", name, flags=re.IGNORECASE)
