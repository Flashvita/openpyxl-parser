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
WORDS_TO_CUT = [
    'для', 'оглы',
]



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
            
        return f"Files Parse completed see result in: `{BASE_DIR}`"

    def read(self, path_to_file):
        data_frame = load_workbook(path_to_file)
        frame = data_frame.active
        #rows_gen = (i for i in range(2, len(frame["A"])))
        rows_gen = (i for i in range(2, 400))
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

    

    def get_members(self, string):
        """
        Возвращает имя компании
          или ФИО физ лица
        """
        if string is None:
            return []
        return [i.strip() for i in re.split(r'\n', re.sub(r'[)(A-Z]', "", string).strip())]

    def update_owner_list(self, author, patent_owners):
        owner_list = patent_owners.remove(author) if author in patent_owners else patent_owners
        if not owner_list:
            return None
        return "\n".join(i for i in owner_list)


    def cut_exception_words(self, string):
        """
        Вырезаем слова исключения и стороки и точки
        при сокращени имени и отчества
        """
        string = string
        for word in WORDS_TO_CUT:
            print('pattern', word)
            string = string.replace(word, "")
        return string.replace(".", " ").strip()

    def author_fio_or_company_title(self, string):
        """
        Получаем фио автора или название фирмы
        """
        try:
            
            string = self.cut_exception_words(string)
            fio = re.split(" ", string)
            len_fio = len(fio)
            if len_fio in [3, 4]:
                print('fio  fio', fio)
                company_name = self.get_company_name(string)
                if len(company_name) != 0:
                    return company_name[0]
                res = f"{fio[0]} {fio[1][0]} {fio[2][0]}"
            elif len(fio) == 2:
                res = f"{fio[0]} {fio[1][0]}"

            else:
                res = self.legal_entity_or_individual(string)
            return res
        except:
            print("Exception in author_fio_or_company_title", string)

    def correct_family(self, family):
        if family[-1] in ["А", "У"]:
            return family[0:-1]
        if family[-2:] in ["ОЙ"]:
            return family[0:-2]
        else:
            return family

    def fio_corrector(self, string):
        """
        Дополнительная проверка фамилии на склонения
        """
        print('fio_corrector IN', string)
        try:
            fio_data = re.split(" ", string)
            if len(fio_data[0]) > 1:
                family = self.correct_family(fio_data[0])
                fio = f"{family} {fio_data[1][0]} {fio_data[2][0]}"
            else:
                family = self.correct_family(fio_data[2])
                fio = f"{family} {fio_data[0][0]} {fio_data[1][0]}"
            print('fio_corrector OUT', fio)
            return fio
        except:
            print("Except fio_converter string", string)
            return string
        
    def get_company_name(self, string):
        """
        Ищет имя компании которое находиться в ковычках
        """
        return re.findall(r'"(.*)"', string)

    def tc_exist(self, string):
        
        """
        Проверка строки адреса на БЦ(торговый центр)
        """
        trading_center = re.findall(r'\БЦ', string)
        return len(trading_center) > 0
    
    def legal_entity_or_individual(self, string):
        """
        Из адресса получаем название фирмы в ковычках
        или ФИО автора исключаем БЦ
        """
        try:
            # Проверяем если в адресе торговый центр
            if self.tc_exist(string):
                list_data = self.get_company_name(string)
                if len(list_data) > 1:
                    return list_data[-1]
                else:
                    list_data = re.split(",",  string)
                    if self.tc_exist(string):
                        # Если нет занятой и из строки не удалился БЦ
                        return " ".join(i for i in list_data[-1].split(" ")[-3:])
                    return self.cut_exception_words(list_data[-1])

            list_data = self.get_company_name(string)
            if len(list_data) != 0:
                # Если нашлось название фирымы в ковычках 
                
                return list_data[0]
            fio = string
            if len(string) < 35:
                # Если  название фирымы в ковычках не нашлось
                list_data = re.split(",", string)
                if len(list_data) != 0:
                    # Значит есть фамилия в конце строки
                    if not self.number_in_string(list_data[-2]):
                        fio = self.cut_exception_words(list_data[-2])
                    else:
                        fio = self.cut_exception_words(list_data[-1])
                else:
                    # Если в адресе есть фамилия
                    fio = f"{fio[0]} {fio[1][0]} {fio[2][0]}"
                print("fio in address string", fio)
                return fio
            else:
                # Если в фамилия в конце строки
                return re.split(",", string)[-1].strip()
        except:
            raise Exception(f"Invalid data company {string}")

    def address_author_fio(self, address):
        try:
            address = (address.replace(".", " ")).strip()
            fio = re.split(" ", address)[-3:]
            return f"{fio[0]} {fio[1][0]} {fio[2][0]}"
        except:
            print("Exception addres author fio", address)

    def number_in_string(self, string):
        list_number = re.findall(r'[0-9]', string)
        return len(list_number) > 0


    def words_firs_indexs(self, string):
        res = []
        i_list = string.split(" ")
        if len(i_list) <= 1:
            return string
        for i in i_list:
            if i:
                res.append(i)
            else:
                continue
        return "".join(i[0] for i in res)
       

    def string_data_valid(self, string_data: list):
        """
        Парсим строку из таблицы
        """
        new_list = []
        # Получаем список всех патентообладателей
        patent_owners = self.get_members(string_data[0])
        # Получаем значение адреса
        address = string_data[2]
        address_upper = address.upper()
        requester = self.legal_entity_or_individual(address)
        authors_list = []
        for owner in patent_owners:
            
            print('patent_owners', patent_owners)
            # Получаем патентообладателя(фио или название компании)
            author = self.author_fio_or_company_title(owner)
            if author is None:
                print("AUTHOR is NONE", author)
                continue
            print("Author patent", author)
            author_upper = author.upper()
            if author_upper not in address_upper:
                    # Проверяем есль ли 
                    # address_author_fio = self.address_author_fio(address)
                    # address_author_upper = address_author_fio.upper()
                    # if author_upper not in address_author_upper:
                    #     
                if requester.upper().replace(" ", "") not in author_upper.replace(" ", ""):
                    requester = requester.replace(".", " ")
                    if author_upper.replace(".", " ") != requester.upper():
                            # Проверяем есть ли цифры в имени заказчика если есть то это адресс
                                # (нет фирмы и имени физ лица)
                        if not self.number_in_string(requester):
                            av = self.fio_corrector(author_upper)
                            print('self.fio_corrector(author_upper)', av)
                            if not av in self.fio_corrector(requester.upper()):
                                    #if self.fio_converter(author).upper() != self.fio_converter(requester).upper():
                                        # Если нужно полное имя из патентообладателя вместо author -> owner
                                if not self.words_firs_indexs(author_upper) in requester.upper():
                                        
                                        authors_list.append(author)
                                        print("authors_list authors_list", authors_list)


        if not self.number_in_string(requester) and len(authors_list) > 0: 
            authors = "\n".join(i for i in authors_list)
            print("results list authors_list", authors)
            new_list.append((authors, requester, address))
            

        return new_list    
        
    # def string_data_valid(self, string_data: list):
    #     """
    #     Вычисляем подходит ли строка под наши условия:

    #     В адресе переписки содержится один из вариантов:
    #     1) адрес
    #     2) адрес, Фамилия Имя Отчество,
    #     3) адрес, Фамилия И.О.
    #     4) адрес, компания,
    #     5) адрес, компания, Фамилия Имя Отчество
    #     6) адрес, компания, Фамилия И.О.
        
    #     Как понять, что строка нас интересует. 
    #     Если патентообладатель и тот, кто делает патент - разные люди

    #     Соответственно существуют варианты (слева патентообладатель, справа - тот кто делает)
    #     1) Человек - человек
    #     2) человек - компания
    #     3) компания - человек
    #     4) компания - компания

    #     Каждый из вариантов анализируется по своему:

    #     1) если ФИО отличаются - наш клиент, в противном случае - отбрасываем
    #     2) берём всегда
    #     3) берём всегда
    #     4) самое сложное - надо понять одна ли компания справа и слева.
    #     Сложность в том, что пишут их по разному.
    #     Как вариант можно справа и слева разбить по словам и сравнить
    #     есть ли справа и слева одинаковые слова. При этом надо исключать слова менее 3 букв.
    #     Так же надо исключать общие слова типа Организация, Предприятие и т.д,
    #     то есть сформировать некоторый словарь игнорируемых слов.
    #     Если одинаковое слово нашлось, то наш клиент.

    #     Итого, что надо сделать:

    #     1) научится выделять человека, фирму и по возможность - адрес.
    #     2) научится принимать решение по 4 вариантам
    #     3) сформировать новую таблицу с выделенными записями
    #     4) подсчитать количество патентов для каждого конкурента

    #     Что надо учесть:
    #     1) в патентообладателе может быть несколько записей, причем вперемешку и люди и фирма.
    #     Значит анализ надо проводить по каждой записи.
    #     Бывает так, что в списке есть человек и он же содержится в адресе переписки.
    #     Это нам не интересно.
    #     2) человек может записываться в разных форматах:
    #     Иванов Иван Иванович или Иванов И.И. следовательно надо понимать, что это один человек
    #     3) фирмы записываются по разному, надо научится понимать, что это одна фирма
    #     4) бывает что название фирмы одно, а адреса разные (типа филиалы).
    #     При подсчёте количества патентов надо такие записи считать вместе.
    #     """
    #     new_list = []
    #     # Получаем список всех патентообладателей
    #     patent_owners = self.get_members(string_data[0])
    #     # Получаем значение адреса
    #     address = string_data[2]
    #     authors_list = []
    #     for owner in patent_owners:
            
    #         print('patent_owners', patent_owners)
    #         # Получаем патентообладателя(фио или название компании)
    #         author = self.author_fio_or_company_title(owner)
    #         if author is None:
    #             print("AUTHOR is NONE", author)
    #             continue
           
    #         print("Author patent", author)
    #         author_upper = author.upper()
    #         address_upper = address.upper()
    #         if author_upper not in address_upper:
    #                 # Проверяем есль ли 
    #                 # address_author_fio = self.address_author_fio(address)
    #                 # address_author_upper = address_author_fio.upper()
    #                 # if author_upper not in address_author_upper:
    #                 #     
    #             requester = self.legal_entity_or_individual(address)
    #             if requester.upper().replace(" ", "") not in author_upper.replace(" ", ""):
    #                 requester = requester.replace(".", " ")
    #                 if author_upper.replace(".", " ") != requester.upper():
    #                         # Проверяем есть ли цифры в имени заказчика если есть то это адресс
    #                             # (нет фирмы и имени физ лица)
    #                     if not self.number_in_string(requester):
    #                         av = self.fio_corrector(author_upper)
    #                         print('self.fio_corrector(author_upper)', av)
    #                         if not av in self.fio_corrector(requester.upper()):
    #                                 #if self.fio_converter(author).upper() != self.fio_converter(requester).upper():
    #                                     # Если нужно полное имя из патентообладателя вместо author -> owner
    #                             if not self.words_firs_indexs(author_upper) in requester.upper():
                                        
    #                                     authors_list.append(author)
    #                                     print("authors_list authors_list", authors_list)


    #     if not self.number_in_string(requester): 
    #         authors = "\n".join(i for i in authors_list)
    #         print("results list authors_list", authors)
    #         new_list.append((authors, requester, address))
            

    #     return new_list


if __name__ == "__main__":
    """
    На выходе - первый документ 
    отфильтрованный список с полями:
        1) патентообладатель
        2) исполнитель
        3) адрес исполнителя

        Второй документ:
        1) исполнитель
        2) количество оформленных  патентов
    """
    parser = ExcelParser()
    res = parser.parse_files()
    print("Parse completed:", res)

