""" Tool to produce an Excel workbook of registrees
for the MD410 2020 Convention

"""
__author__ = "Kim van Wyk"
__version__ = "0.0.1"

from collections import namedtuple
from openpyxl import Workbook
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment
from openpyxl.styles import Font

THIN_BORDER = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
COLUMN = namedtuple("COLUMN", ("title", "width", "column_name"))

TITLE_HEIGHT = 40
NORMAL_HEIGHT = 25
ROWS_PER_PAGE = 31
BOLD = Font(name="Arial", bold=True)
CENTRE = Alignment(horizontal="center", vertical="center")
COLUMNS = [
    COLUMN("Name", 35, "A"),
    COLUMN("Reg\nNumber", 10, "B"),
    COLUMN("Status", 15, "C"),
    COLUMN("Signature – registered\nand gift bag received\n(if applicable)", 35, "D"),
]

wb = Workbook()

dest_filename = "registrees_list.xlsx"
ws1 = wb.active
ws1.title = "By Name"
ws1.page_setup.paperSize = ws1.PAPERSIZE_A4
ws1.page_margins.top = 0.2
ws1.page_margins.bottom = 0.2
ws1.page_margins.left = 0.2
ws1.page_margins.right = 0.2

for column in COLUMNS:
    ws1.column_dimensions[column.column_name].width = column.width
row = 1
while row < 40:
    if not ((row - 1) % ROWS_PER_PAGE):
        for column in COLUMNS:
            ws1[f"{column.column_name}{row}"] = column.title
            ws1[f"{column.column_name}{row}"].font = BOLD
            ws1[f"{column.column_name}{row}"].alignment = CENTRE
            ws1[f"{column.column_name}{row}"].border = THIN_BORDER
        ws1.row_dimensions[row].height = TITLE_HEIGHT
    else:
        for column in COLUMNS:
            ws1[f"{column.column_name}{row}"].border = THIN_BORDER
        ws1.row_dimensions[row].height = NORMAL_HEIGHT
    row += 1

wb.save(filename=dest_filename)

