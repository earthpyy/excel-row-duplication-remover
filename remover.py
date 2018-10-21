import argparse

from openpyxl.reader.excel import load_workbook


parser = argparse.ArgumentParser()
parser.add_argument('file_name', type=str, help='Excel file name')


def main(args):
    wb = load_workbook()


if __name__ == '__main__':
    args = parser.parse_args()
    print(type(args))
    # main(args)
