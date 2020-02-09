""" Tool to produce an Excel workbook of registrees
for the MD410 2020 Convention

"""
__author__ = "Kim van Wyk"
__version__ = "0.0.1"

from collections import namedtuple
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

COLUMN = namedtuple("COLUMN", ("title", "width", "column_name"))

TITLE_HEIGHT = 1.5
NORMAL_HEIGHT = 1
ROWS_PER_PAGE = 22
COLUMNS = [
    COLUMN("Name", 7.5, "A"),
    COLUMN("Reg\nNumber", 1.5, "B"),
    COLUMN("Status", 4, "C"),
    COLUMN("Signature – registered\nand gift bag received\n(if applicable)", 4, "D"),
]

wb = Workbook()

dest_filename = "registrees_list.xlsx"
ws1 = wb.active
ws1.title = "By Name"

for column in COLUMNS:
    ws1.column_dimensions[column.column_name].width = column.width
row = 1
while row < 40:
    if not ((row - 1) % ROWS_PER_PAGE):
        for column in COLUMNS:
            ws1[f"{column.column_name}{row}"] = column.title
        ws1.row_dimensions[row] = TITLE_HEIGHT
    else:
        ws1[f"A{row}"] = row
        ws1.row_dimensions[row] = NORMAL_HEIGHT
    row += 1

wb.save(filename=dest_filename)

