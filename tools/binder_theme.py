"""Shared colors for the complete binder's editable tables."""
from openpyxl.styles import PatternFill

DARK_GREEN = "305040"
PALE_GREEN = "DDF2DF"
WHITE = "FFFFFF"


def row_fill(index):
    return PatternFill("solid", fgColor=PALE_GREEN if index % 2 else WHITE)
