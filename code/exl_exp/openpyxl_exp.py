from openpyxl import Workbook


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

    workbook.save("basic-openpyxl.xlsx")


basic()
