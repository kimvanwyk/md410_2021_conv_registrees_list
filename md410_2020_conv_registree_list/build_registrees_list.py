""" Tool to produce an Excel workbook of registrees
for the MD410 2020 Convention

"""
__author__ = "Kim van Wyk"
__version__ = "0.0.1"

from collections import namedtuple
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from md410_2020_conv_common import db


THIN_BORDER = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
COLUMN = namedtuple("COLUMN", ("title", "width", "column_name"))

TITLE_HEIGHT = 40
NORMAL_HEIGHT = 25
ROWS_PER_PAGE = 29
BOLD = Font(name="Arial", bold=True)
CENTRE = Alignment(horizontal="center", vertical="center")
COLUMNS = [
    COLUMN("Name", 35, "A"),
    COLUMN("Reg\nNumber", 10, "B"),
    COLUMN("Status", 15, "C"),
    COLUMN("Signature – registered\nand gift bag received\n(if applicable)", 35, "D"),
]
NOW = datetime.now()

wb = Workbook()
dest_filename = "registrees_list.xlsx"


def write_sheet(sheet, title, registrees):
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_margins.top = 0.8
    sheet.page_margins.bottom = 0.2
    sheet.page_margins.left = 0.2
    sheet.page_margins.right = 0.2

    sheet.oddHeader.center.text = f"Registrees {title} as of {NOW:%d/%m/%Y %H:%M}"
    sheet.oddHeader.center.size = 12
    sheet.oddHeader.center.font = "Arial,Bold"

    for column in COLUMNS:
        sheet.column_dimensions[column.column_name].width = column.width
    row = 1

    for registree in registrees:
        if not ((row - 1) % ROWS_PER_PAGE):
            for column in COLUMNS:
                sheet[f"{column.column_name}{row}"] = column.title
                sheet[f"{column.column_name}{row}"].font = BOLD
                sheet[f"{column.column_name}{row}"].alignment = CENTRE
                sheet[f"{column.column_name}{row}"].border = THIN_BORDER
            sheet.row_dimensions[row].height = TITLE_HEIGHT
            row += 1
        for column in COLUMNS:
            sheet[f"{column.column_name}{row}"].border = THIN_BORDER
        sheet.row_dimensions[row].height = NORMAL_HEIGHT
        sheet[f"A{row}"] = f"{registree.last_name}, {registree.first_names}"
        sheet[f"B{row}"] = f"{registree.reg_num:03}"
        row += 1


dbh = db.DB()
registrees = dbh.get_all_registrees()

sheet = wb.active
title = "By Name"
sheet.title = title
registrees.sort(key=lambda x: f"{x.last_name}, {x.first_names}")
write_sheet(sheet, title, registrees)

title = "By Reg Num"
sheet = wb.create_sheet(title)
registrees.sort(key=lambda x: x.reg_num)
write_sheet(sheet, title, registrees)

wb.save(filename=dest_filename)

