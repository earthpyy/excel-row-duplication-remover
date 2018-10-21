import argparse
from argparse import ArgumentParser, Namespace

from openpyxl.reader.excel import load_workbook


parser = ArgumentParser()
parser.add_argument('file_name', type=str, help='Excel file name')
parser.add_argument('--column', dest='column', default='A', help='Specify which column to check (default: A)')


def main(args: Namespace):
    def find_duplication(ws, col: int = 0, skip_header: bool = True):
        last_index = 1
        start_row = 3 if skip_header else 2
        rows = ws.iter_rows(min_row=start_row)
        for index, row in enumerate(rows):
            last_row = ws[last_index]
            if last_row is not None and last_row[col] != row[col]:
                # remove duplication
                ws.delete_rows(last_index, index + 1)
                # reset last_index
                last_index = index + 1

    file_name = args.file_name
    column = ord(args.column) - ord('A')
    wb = load_workbook(filename=file_name)
    ws = wb.active
    result_ws = find_duplication(ws, col=column)


if __name__ == '__main__':
    args = parser.parse_args()
    main(args)
