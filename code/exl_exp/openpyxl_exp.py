from openpyxl import Workbook
from openpyxl.drawing.image import Image


def basic():
    workbook = Workbook()
    active_sheet = workbook.active
    active_sheet.title = "Basic Sheet"

    data_bank = [
        [10, 11, 12, 13],
        [14, 15, 16, 17],
        [18, 19, 20, 21],
    ]

    for data in data_bank:
        active_sheet.append(data)

    workbook.save("ignore/basic-openpyxl.xlsx")


def add_image():
    workbook = Workbook()
    active_sheet = workbook.active
    active_sheet.title = "Image Sheet"

    logo = Image("assets/logo.png")
    logo.height = 150
    logo.width = 150

    active_sheet.add_image(logo, "A3")
    workbook.save("image-openpyxl.xlsx")


basic()
add_image()
