import os
import re

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parent
parser_files_dir = os.listdir(f"{BASE_DIR}/files_to_parse")


def get_max_row(column):
    for x in range(len(column)):
        if column[x].value != None or x == 1:
            continue
        else:
            return x
    return int(x) + 1


class ExcelParser:
    """
    Чтение данны из файла с помощью openpyxl
    Обработка текста с помощью регуоярных выражений
    """
    def __init__(self):
        self.path_to_folder = f"{BASE_DIR}/files_to_parse/"

    def format_is_valid(self, file):
        _, format = file.split(".")
        return format in ["xlsx", "xlsx"]

    def file_path(self, file):
        return f"{self.path_to_folder}{file}"

    def write_new_file(self, data):
        wb = Workbook()
        ws = wb.active
        # add column headings. NB. these must be strings
        ws.append(["Патентообладатель", "Исполнитель", "Адрес исполнителя"])
        for row in data:
            for i in row:
                ws.append(i)
        tab = Table(displayName="Table1", ref="A1:B2")

        # Add a default style with striped rows and banded columns
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True,
                               showColumnStripes=True
                               )
        tab.tableStyleInfo = style
        ws.add_table(tab)
        time_created = datetime.now().strftime("%d-%m-%Y_%H:%M")
        wb.save(f"test_{time_created}.xlsx")

    def parse_files(self):
        for file in parser_files_dir:
            if not self.format_is_valid(file):
                continue
            file_path = self.file_path(file)
            data = self.read(file_path)
            self.write_new_file(data)
        return f"File Parse completed see result document: `{BASE_DIR}`"

    def rows_generator(self, start, finish):
        for i in range(start, finish):
            yield i + 1

    def read(self, path_to_file):
        data_frame = load_workbook(path_to_file)
        frame = data_frame.active
        finish_row = get_max_row(frame['A'])
        print('finish_row', finish_row)
        rows_gen = (i + 1 for i in range(2, finish_row))
        return self.get_data_row(rows_gen, frame)

    def get_data_row(self, rows, frame, data=[]):
        for row in rows:
            string_data = []
            for col in frame.iter_cols(9, 12):
                string_data.append(col[row].value)
            new_data = self.string_data_valid(string_data)
            data.append(new_data)
        return data

    def get_members(self, string):
        if string is None:
            return []
        return [i.strip() for i in re.split(r'\n', re.sub(r'[)(RU]', "", string).strip())]

    def string_data_valid(self, string_data: list):
        new_list = []
        authors = self.get_members(string_data[0])
        patent_owners = self.get_members(string_data[1])

        for a in authors:
            if a[0] and a[1][0] and a[2][0] not in patent_owners:
                owners = "\n".join(i for i in patent_owners)
                new_list.append((a, owners, string_data[3]))
        return new_list


if __name__ == "__main__":
    parser = ExcelParser()
    res = parser.parse_files()
    print("Parse completed:", res)
