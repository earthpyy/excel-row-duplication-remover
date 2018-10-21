import argparse
from argparse import ArgumentParser, Namespace

from openpyxl.reader.excel import load_workbook


parser = ArgumentParser()
parser.add_argument('file_name', type=str, help='Excel file name')
parser.add_argument('--column', dest='column', default='A', help='Specify which column to check (default: A)')
parser.add_argument('--result', dest='result_name', default=None, help='Specify result file name')


def main(args: Namespace):
    def find_duplication(ws, col: int = 0, skip_row: int = 1):
        count = -1
        start_row = 1 + skip_row
        rows = ws.iter_rows(min_row=start_row)
        for index, row in enumerate(rows):
            last_row = ws[index + skip_row]
            print(last_row[col].value, row[col].value, count)

            if count == -1 or last_row[col].value == row[col].value:
                count += 1
            else:
                # remove duplication
                ws.delete_rows(index - count + skip_row, count)
                # reset count
                count = 0

            if index == 25:
                break

    # parse arguments
    file_name = args.file_name
    result_name = args.result_name
    if len(args.column) == 1:  # TODO: support for column AA-AZ...
        column = ord(args.column) - ord('A')
    else:
        print('Column name must length 1')
        return

    # laod worksheet
    print(f'Loading file {file_name}...')
    wb = load_workbook(filename=file_name)
    ws = wb.active
    print(f'Finding duplication in column {args.column}...')
    find_duplication(ws, col=column)
    print(f'Saving result file to {result_name}...')
    wb.save('new.xlsx')


if __name__ == '__main__':
    args = parser.parse_args()
    main(args)
