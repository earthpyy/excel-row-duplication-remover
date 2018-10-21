import argparse
from argparse import ArgumentParser, Namespace

from openpyxl.reader.excel import load_workbook


parser = ArgumentParser()
parser.add_argument('file_name', type=str, help='Excel file name')


def main(args: Namespace):
    def find_duplication(ws):
        for row in ws.rows:
            pass

    file_name = args.file_name
    wb = load_workbook(filename=file_name)
    ws = wb.active
    result_ws = find_duplication(ws)


if __name__ == '__main__':
    args = parser.parse_args()
    main(args)
