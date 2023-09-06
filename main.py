import os
import re

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parent
parser_files_dir = os.listdir(f"{BASE_DIR}/files_to_parse")

WORDS_IGNORE = [
    "ограниченной",
    "ответственностью",
    "Федеральное Государственное Бюджетное Образовательное Учреждение Высшего Образования",
    "Федеральное Государственное Бюджетное Научное Учреждение",
    "Федеральное Государственное Автономное Образовательное Учреждение Высшего Образования",
    "Публичное",
    "Общество",
    "Акционерное"
]
EXCEPTION_LIST = [
    "НАУЧНО",
]

def address_author(address):
    fio = re.split(" ", address)[-3:]
    return f"{fio[0]} {fio[1][0]} {fio[2][0]}"


class ExcelAggregateData:
    """
    Запись данных в таблицу
    1) исполнитель
    2) количество оформленных  патентов
    """

    def make_dict_data(self, data):
        data_list = []
        for obj in data:
            for i in obj:
                data_list.append(i[0])
        return data_list

    def write_new_file(self, data):
        from collections import Counter
        wb = Workbook()
        ws = wb.active
        new_data = self.make_dict_data(data)
        res = Counter(new_data).most_common()
        #add column headings. NB. these must be strings
        ws.append(["Патентообладатель", "Количество патентов"])
        for i in res:
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
        wb.save(f"aggregator_{time_created}.xlsx")
        del new_data


class ExcelPatentData:
    """
    Запись данных в таблицу
    1) патентообладатель
    2) исполнитель
    3) адрес исполнителя

    """
    def write_new_file(self, data):
        print('worksheet data', data)
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


class ExcelParser:
    """
    Чтение данны из файла с помощью openpyxl
    # https://openpyxl.readthedocs.io/en/stable/
    Обработка текста с помощью регулярных выражений
    """
    def __init__(self):
        self.path_to_folder = f"{BASE_DIR}/files_to_parse/"

    def format_is_valid(self, file):
        #_, format = file.split(".")
        format = file[-4:]
        return format in ["xlsx", "xlsm", "xls", "xltx"]

    def file_path(self, file):
        return f"{self.path_to_folder}{file}"

    def parse_files(self):
        for file in parser_files_dir:
            if not self.format_is_valid(file):
                continue
            file_path = self.file_path(file)
            data = self.read(file_path)
            ExcelPatentData().write_new_file(data)
            #ExcelAggregateData().write_new_file(data)
            data
        return f"Files Parse completed see result in: `{BASE_DIR}`"

    def read(self, path_to_file):
        data_frame = load_workbook(path_to_file)
        frame = data_frame.active
        # rows_gen = (i for i in range(2, len(frame["A"])))
        rows_gen = (i for i in range(2, 50))
        return self.get_data_row(rows_gen, frame)

    def get_data_row(self, rows, frame, data=[]):
        for row in rows:
            string_data = []
            for col in frame.iter_cols(10, 12):
                print('row number', col[row])
                string_data.append(col[row].value)
            new_data = self.string_data_valid(string_data)
            data.append(new_data)
        return data

    def check_author_in_address(data):
        pass

    def get_members(self, string):
        """
        Return  FIO or Company name
        """
        print('get_members', string)
        if string is None:
            return []
        return [i.strip().upper() for i in re.split(r'\n', re.sub(r'[)(A-Z]', "", string).strip())]

    def update_owner_list(self, author, patent_owners):
        owner_list = patent_owners.remove(author) if author in patent_owners else patent_owners
        if not owner_list:
            return None
        return "\n".join(i for i in owner_list)

    def author_fio(self, string):
        """
        Получаем фио автора или название фирмы
        """
        #print("fio string", string)
        fio = re.split(" ", string)
        if len(fio) == 3:
            res = f"{fio[0]} {fio[1][0]} {fio[2][0]}"
        elif len(fio) == 2:
            res = f"{fio[0]} {fio[1][0]}"
        if "оглы" in fio:
            res = f"{fio[0]} {fio[1][0]} {fio[2][0]} {fio[3][0]}"

        else:
            res = self.get_company_name(string)
        return res
    
    def fio_converter(self, string):
        try:
            print('fio_converter string', string)
            fio_data = re.split(" ", string)
            if len(fio_data[0]) > 1:
                return f"{fio_data[0]} {fio_data[1][0]} {fio_data[2][0]}"
            else:
                return f"{fio_data[2]} {fio_data[0][0]} {fio_data[1][0]}"
        except:
            print("Except fio_converter string", string)
            return string
        
    def get_company_name(self, string):
        try:
            list_name = re.findall(r'"(.*?)"', string)
            if len(list_name) != 0:
                # title = re.split("-", list_name[0])[0].strip()
                # print("title", title)
                if list_name[0] in EXCEPTION_LIST:
                    print('EXCEPTION_LIST !')
                    return list_name[0]
                return list_name[0]
            fio = string
            if len(string) < 35:
                return f"{fio[0]} {fio[1][0]} {fio[2][0]}"
            else:
                return re.split(",", string)[-1].strip()
        except:
            raise Exception(f"Invalid data company {string}")
    
    def address_author_fio(self, address):
        fio = re.split(" ", address)[-3:]
        return f"{fio[0]} {fio[1][0]} {fio[2][0]}"

    def number_in_string(self, string):
        list_number = re.findall(r'[0-9]', string)
        return len(list_number) > 0
    def string_data_valid(self, string_data: list):
        new_list = []
        # Получаем список всех патентообладателей
        patent_owners = self.get_members(string_data[0])
        # Получаем список всех патентообладателей
        address = string_data[2]
        # Получаем значение адреса
        for owner in patent_owners:
            # Получаем патентообладателя(фио или название компании)
            name = self.author_fio(owner)
            print("name", name)
            if name not in address.upper():
                address_author_fio = self.address_author_fio(address)
                if name not in address_author_fio:
                    if not name.replace(" ", "") in address_author_fio.replace(" ", ""):
                        author_fio = self.get_company_name(address)
                        if not self.number_in_string(author_fio):
                            if self.fio_converter(name).upper() != self.fio_converter(author_fio).upper():
                                # Если нужно полное имя из патентообладателя вместо name.title() -> owner
                                new_list.append((name.title(), author_fio, address))

                else:
                    continue
            else:
                continue
        return new_list

    def get_address_from_string(string):
        z = []
        res = re.split(" ", string)
        for a in res:
            if not a.isalnum():
                z.append(a)
            else:
                z.append(a)
                return " ".join(i for i in z).strip(",").strip()
        return " ".join(i for i in z).strip(",").strip()

if __name__ == "__main__":
    parser = ExcelParser()
    res = parser.parse_files()
    print("Parse completed:", res)


# class ExcelParser:
#     """
#     Чтение данны из файла с помощью openpyxl
#     # https://openpyxl.readthedocs.io/en/stable/
#     Обработка текста с помощью регулярных выражений
#     """
#     def __init__(self):
#         self.path_to_folder = f"{BASE_DIR}/files_to_parse/"

#     def format_is_valid(self, file):
#         #_, format = file.split(".")
#         format = file[-4:]
#         return format in ["xlsx", "xlsm", "xls", "xltx"]

#     def file_path(self, file):
#         return f"{self.path_to_folder}{file}"

#     def parse_files(self):
#         for file in parser_files_dir:
#             if not self.format_is_valid(file):
#                 continue
#             file_path = self.file_path(file)
#             data = self.read(file_path)
#             ExcelPatentData().write_new_file(data)
#             #ExcelAggregateData().write_new_file(data)
#         return f"Files Parse completed see result in: `{BASE_DIR}`"

#     def read(self, path_to_file):
#         data_frame = load_workbook(path_to_file)
#         frame = data_frame.active
#         # rows_gen = (i for i in range(2, len(frame["A"])))
#         rows_gen = (i for i in range(2, len(frame["A"])))
#         return self.get_data_row(rows_gen, frame)

#     def get_data_row(self, rows, frame, data=[]):
#         for row in rows:
#             string_data = []
#             for col in frame.iter_cols(9, 12):
#                 print('row number', col[row])
#                 string_data.append(col[row].value)
#             new_data = self.string_data_valid(string_data)
#             data.append(new_data)
#         return data

#     def check_author_in_address(data):
#         pass

#     def get_members(self, string):
#         """
#         Return  FIO or Company name
#         """
#         print('get_members', string)
#         if string is None:
#             return []
#         return [i.strip().upper() for i in re.split(r'\n', re.sub(r'[)(RU]', "", string).strip())]

#     def update_owner_list(self, author, patent_owners):
#         owner_list = patent_owners.remove(author) if author in patent_owners else patent_owners
#         if not owner_list:
#             return None
#         return "\n".join(i for i in owner_list)

#     def author_fio(self, string):
#         """
#         [Ivanov Sergey Petrovich]
#         return Ivanov S P
#         """
#         fio = re.split("", string)
#         print('fio', type(fio))
#         return f"{fio[0]} {fio[1][0]} {fio[2][0]}"

#     def string_data_valid(self, string_data: list):
#         new_list = []
#         authors = self.get_members(string_data[0])
#         patent_owners = self.get_members(string_data[1])
#         for author in authors:
#             if self.author_fio(author) not in patent_owners:
#                 # owners = self.update_owner_list(author, patent_owners)
#                 if owners := self.update_owner_list(author, patent_owners):
#                     if author.upper() in string_data[3].upper():
#                         print('string_data[3]', string_data[3])
#                         continue
#                     new_list.append((author.title(), owners.title(), string_data[3]))
#                 else:
#                     continue
#         return new_list


# if __name__ == "__main__":
#     parser = ExcelParser()
#     res = parser.parse_files()
#     print("Parse completed:", res)
