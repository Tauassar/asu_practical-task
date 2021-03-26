from openpyxl import load_workbook

from utils import check_iin


def print_hi(name):
    url = 'static/xlsx/Исходный файл.xlsx'
    wb = load_workbook(url)
    sheets = wb.sheetnames
    work_sheet = wb[sheets[0]]
    # for value in work_sheet.iter_rows(values_only=True):
    # print(value)


if __name__ == '__main__':
    iin = '970113351179'
    print(check_iin(iin))

