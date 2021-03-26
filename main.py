"""
Description: python script for check and retrieving data from xlsx,
and render correct data to another xlsx
Author: Tauassar Tatiyev
"""
import log

from openpyxl import load_workbook

import settings
from utils import check_iin, check_bin

log.setup_logging('DEBUG' if settings.DEBUG else 'INFO')

def print_hi(name):
    wb = load_workbook(settings.inp_file)
    sheets = wb.sheetnames
    work_sheet = wb[sheets[0]]
    # for value in work_sheet.iter_rows(values_only=True):
    # print(value)


if __name__ == '__main__':
    iin = '019743351179'
    check_bin(iin)

