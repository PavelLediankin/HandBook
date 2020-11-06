import inspect
import sys
import openpyxl as xl


class HandBook:
    def __init__(self, name: str):
        self.name = name
        self.path = "D:\Python и все связанное\Справочник\HandBook.xlsx"
        self.count = 7  # ------------------------------------------------------> hb.rows count
        self.initialise_file()  # Создание таблицы для справочника - возможно создавать новые справочники
                                # на листах в том же файле. Инициализация только один раз ? флаг

    def initialise_file(self):
        hb = self.get_hb()
        sheet = hb["1"]
        sheet["A1"] = "Фамилия"
        sheet["B1"] = "Имя"
        sheet["C1"] = "Телефон"
        sheet["D1"] = "Город"
        sheet["E1"] = "e-mail"
        hb.save(self.path)

    __dict = {"surname": "A", "name": "B", "phone_number": "C", "city": "D", "email": "E"}

    # region Main methods
    def make_record(self, surname: str, name: str, phone_number: str, city: str, email: str):
        i = self.count + 1
        # Check --------------------------
        # Find(email)---------------------
        hb = self.get_hb()
        sheet = hb["1"]
        sheet[f"A{i}"] = surname   # modification - Выделить метод? Может удастся сделать один метод?
        sheet[f"B{i}"] = name
        sheet[f"C{i}"] = phone_number
        sheet[f"D{i}"] = city
        sheet[f"E{i}"] = email
        hb.save(self.path)
        self.count = i
        return "Ok"

    def modify_record(self, surname="", name="", phone_number="", city="", email="", i=-1):
        hb = self.get_hb()
        sheet = hb["1"]
        if i == -1:
            i = self.get_index_for_change("modify", surname, name, phone_number, city, email)
            if not i.isdigit():
                return i
        # modification - Выделить метод?  -- исп словарь (иф - словарь + стр,)
        sheet[f"A{i}"] = surname if surname != "" else sheet[f"A{i}"]
        # ....
        hb.save(self.path)
        return "Ok"

    def delete_record(self, surname="", name="", phone_number="", city="", email="", i=-1):
        hb = self.get_hb()
        sheet = hb["1"]
        if i == -1:
            i = self.get_index_for_change("delete", surname, name, phone_number, city, email)
            if not i.isdigit():
                return i
        sheet.delete_rows(i, 1)
        self.count -= 1
        hb.save(self.path)
        return "Ok"

    def find(self, surname="", name="", phone_number="", city="", email=""):
        indexes = self.get_founded_indexes(surname, name, phone_number, city, email)
        result = ""
        hb = self.get_hb()
        sheet = hb["1"]
        for i in indexes:
            for e in sheet[i]:
                result += e.value + " "
            result += "\n"
        if result == "":
            return "Not Found!"
        return result

    @property
    def records_count(self):
        return self.count
    # endregion

    # region Inner methods

    def __find(self, key: str, column: str):
        result = []
        hb = self.get_hb()
        sheet = hb["1"]
        vals = [v[0].value for v in sheet[f"{column}2:{column}{self.count}"]]
        for i, v in enumerate(vals):
            if v == key:
                result.append(i + 2)
        return result

    def get_founded_indexes(self, surname="", name="", phone_number="", city="", email=""):
        args = inspect.getfullargspec(self.find)[0][1:]
        d = inspect.getargvalues(sys._getframe())[3]
        for arg in args:
            if d[arg] != "":
                return self.__find(d[arg], self.__dict[arg])

    def get_index_for_change(self, change: str, surname="", name="", phone_number="", city="", email=""):
        record = self.get_founded_indexes(surname, name, phone_number, city, email)
        i = record[0]
        if len(record) != 1:
            return f"You can`t {change} several records. Use another key"
        return str(i)

    def get_hb(self):
        return xl.load_workbook(self.path)

    # endregion
# /--------------------------------------------------------------------------------------------------------------------/

a = HandBook("Name")
hb = a.get_hb()
print(dir(hb))
s = hb["1"]

print(dir(hb["1"]))

# def menu():
#     while True:
#         Choose_operatinon

