import datetime
import logging
import pandas as pd
import pymongo
import sys
import functools
import typing
import re

from PyQt5.QtCore import QSize, Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, \
    QWidget, QListWidget, QListWidgetItem, QVBoxLayout, QLineEdit, QLabel, \
    QHBoxLayout, QAction, QFileDialog, QComboBox, QMessageBox

class DataBase:
    def __init__(
        self, 
        host: str, 
        port: int, 
        dbname: str, 
        user: str = None,
        password: str = None
    ):
        url = "mongodb://"
        
        if user and password:
            url += f"{user}:{password}@"
        
        url += f"{host}:{port}/{dbname}"
        self.__client = pymongo.MongoClient(url)
        self.__db = self.__client[dbname]

    def get_teachers(self, filters: dict = None):
        try:
            if filters:
                elmnts = self.__db.teachers.find(filters)
            else:
                elmnts = self.__db.teachers.find()
        except:
            elmnts = None

        return elmnts

    def set_teacher(self, data: dict) -> typing.Union[str, None]:
        ans = self.__db.teachers.find_one({
            "fac": data["fac"],
            "teacher": data["teacher"],
            "$or": [
                {"year": data["year"]},
                {"year": data["state_year"]},
                {"state_year": data["state_year"]},
                {"state_year": data["year"]},
            ]
        })

        if ans:
            raise Exception(
                "Помилка: викладач вже отримав нагороду у вказаному році"
            )

        try:
            self.__db.teachers.insert_one(data)
        except Exception as ex_:
            logging.error(ex_)

    def set_and_rm_teachers(self, data: typing.List[dict]):
        try:
            self.__db.teachers.drop()
            self.__db.teachers.insert_many(data)
        except Exception as ex_:
            logging.error(ex_)

    def get_facs(self):
        try:
            elmnts = self.__db.facs.find()
        except Exception as ex_:
            logging.error(ex_)
        else:
            return elmnts

    def get_kpi_awards(self):
        try:
            elmnts = self.__db.kpi_awards.find()
        except Exception as ex_:
            logging.error(ex_)
        else:
            return elmnts

    def get_state_awards(self):
        try:
            elmnts = self.__db.state_awards.find()
        except Exception as ex_:
            logging.error(ex_)
        else:
            return elmnts

    def get_kpi_award_next(self, data: dict):
        try:
            elmnt = self.__db.kpi_awards.find_one(data)["id"]
            name = self.__db.kpi_awards.find_one(
                {"id": int(elmnt) + 1}
            )["name"]
        except Exception as ex_:
            logging.error(ex_)
        else:
            return name

    def get_state_award_next(self, data: dict):
        try:
            elmnt = self.__db.state_awards.find_one(data)["id"]
            name = self.__db.kpi_awards.find_one(
                {"id": int(elmnt) + 1}
            )["name"]
        except Exception as ex_:
            logging.error(ex_)
        else:
            return name

    def check_facs(self, fac: str):
        try:
            self.__db.facs.find_one({"name": fac})
        except Exception as ex_:
            logging.info(ex_)
            return False
        
        return True

    def check_kpi_awards(self, award: str):
        try:
            self.__db.kpi_awards.find_one({"name": award})
        except:
            return False
        
        return True

    def check_state_awards(self, award: str):
        try:
            self.__db.state_awards.find_one({"name": award})
        except:
            return False
        
        return True
        
class InformationFromDB(QListWidgetItem):
    def __init__(self, data: str):
        super().__init__()

        self.setText(data)

class Table(QListWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.__temp_data = []

        self.setStyleSheet(
            "QListWidget {"
                "background-color: whitesmoke;"
                "border: 1px solid black;"
                "max-width: 2000px;"
                "min-width: 500px;"
                "max-height: 900px;"
                "margin: 5px;"
            "}"
            "QListWidget::item {"
                "background-color: #82ccdd;"
                "border: 1px solid grey;"
                "border-radius: 2px;"
                "margin: 2px;"
            "}"
        )

    @property
    def temp_data(self):
        return self.__temp_data.copy()

    def show_data(self, data):
        self.clear()
        self.__temp_data.clear()

        now = datetime.datetime.now()

        for el in data:
            self.__temp_data.append(el)

            gram, kpi = ((el["gram"], True )if len(el["gram"]) > 0 
                         and el["gram"] != "nan"
                         else (el["state_gram"], False))
            year = el["year"] if kpi else el["state_year"]
            year = int(year)
            prog = el["prog"]
            year_d = year

            if (year_d >= now.year - 1 and 
                prog and len(prog) > 0 and 
                prog != "nan"):
                prog_m = f", за прогнозом є можливість отримати {prog} у {year_d}"
            else:
                prog_m = ""

            self.addItem(
                f"{el['teacher']} отримав  у {year} році нагороду "
                f"{gram}{prog_m}"
            )

class Filters(QWidget):
    def __init__(self, parent):
        super().__init__(parent)

        lable_fac = QLabel(self)
        lable_fac.setText("Факультет/ННІ")

        self.__input_fac = QLineEdit(self)
        self.__input_fac.setObjectName("fac")

        layout_fac = QHBoxLayout()
        layout_fac.addWidget(lable_fac)
        layout_fac.addWidget(self.__input_fac)

        lable_teacher = QLabel(self)
        lable_teacher.setText("ПІБ")

        self.__input_teacher = QLineEdit(self)
        self.__input_teacher.setObjectName("teacher")

        layout_teacher = QHBoxLayout()
        layout_teacher.addWidget(lable_teacher)
        layout_teacher.addWidget(self.__input_teacher)

        lable_gram = QLabel(self)
        lable_gram.setText("Нагорода")

        self.__input_gram = QLineEdit(self)
        self.__input_gram.setObjectName("gram")

        layout_gram = QHBoxLayout()
        layout_gram.addWidget(lable_gram)
        layout_gram.addWidget(self.__input_gram)

        lable_state_gram = QLabel(self)
        lable_state_gram.setText("Державна Нагорода")

        self.__input_state_gram = QLineEdit(self)
        self.__input_state_gram.setObjectName("state_gram")

        layout_state_gram = QHBoxLayout()
        layout_state_gram.addWidget(lable_state_gram)
        layout_state_gram.addWidget(self.__input_state_gram)

        lable_num_prot = QLabel(self)
        lable_num_prot.setText("№ протоколу ВР")

        self.__input_num_prot = QLineEdit(self)
        self.__input_num_prot.setObjectName("num")

        layout_num_prot = QHBoxLayout()
        layout_num_prot.addWidget(lable_num_prot)
        layout_num_prot.addWidget(self.__input_num_prot)

        lable_year = QLabel(self)
        lable_year.setText("Рік відзначення КПІ")

        self.__input_year = QLineEdit(self)
        self.__input_year.setObjectName("year")

        layout_year = QHBoxLayout()
        layout_year.addWidget(lable_year)
        layout_year.addWidget(self.__input_year)

        lable_state_year = QLabel(self)
        lable_state_year.setText("Рік відзначення державою")

        self.__input_state_year = QLineEdit(self)
        self.__input_state_year.setObjectName("state_year")

        layout_state_year = QHBoxLayout()
        layout_state_year.addWidget(lable_state_year)
        layout_state_year.addWidget(self.__input_state_year)

        lable_prog = QLabel(self)
        lable_prog.setText("Прогнозування")

        self.__input_prog = QLineEdit(self)
        self.__input_prog.setObjectName("prog")

        layout_prog = QHBoxLayout()
        layout_prog.addWidget(lable_prog)
        layout_prog.addWidget(self.__input_prog)

        layout_filters = QVBoxLayout()
        layout_filters.addLayout(layout_fac)
        layout_filters.addLayout(layout_teacher)
        layout_filters.addLayout(layout_gram)
        layout_filters.addLayout(layout_state_gram)
        layout_filters.addLayout(layout_num_prot)
        layout_filters.addLayout(layout_year)
        layout_filters.addLayout(layout_state_year)
        layout_filters.addLayout(layout_prog)

        self.setStyleSheet(
            "QWidget {"
                "max-width: 450px;"
                "max-height: 900px;"
                "margin: 5px;"
            "}"
            "QLineEdit {"
                "max-width: 300px;"
                "max-height: 20px;"
                "min-width: 200px;"
                "min-height: 20px;"
                "margin: 2px;"
            "}"
            "QLabel {"
                "max-width: 300px;"
                "max-height: 20px;"
                "min-width: 200px;"
                "min-height: 20px;"
                "margin: 2px;"
            "}"
            "QPushButton {"
                "max-height: 20px;"
                "min-height: 20px;"
                "margin: 2px;"
            "}"
        )
        self.setLayout(layout_filters)

    def get_filters(self, filters: dict):
        data = {
            "fac": self.__input_fac.text().strip(),
            "teacher": self.__input_teacher.text().strip(),
            "gram": self.__input_gram.text().strip(),
            "state_gram": self.__input_state_gram.text().strip(),
            "num": self.__input_num_prot.text().strip(),
            "year": self.__input_year.text().strip(),
            "state_year": self.__input_state_year.text().strip(),
            "prog": self.__input_prog.text().strip()
        }

        for k, v in data.items():
            if v != "":
                filters.update({k: v})

        return filters

    def clear_filters(self):
        self.__input_fac.clear()
        self.__input_teacher.clear()
        self.__input_gram.clear()
        self.__input_state_gram.clear()
        self.__input_num_prot.clear()
        self.__input_year.clear()
        self.__input_state_year.clear()
        self.__input_prog.clear()

class WindowToInsertData(QWidget):
    def __init__(self, parent, facs, grams, state_grams):
        super().__init__(parent)

        self.__data = self.__set_data()
        self.__facs = facs
        self.__grams = grams
        self.__state_grams = state_grams

        lable_fac = QLabel(self)
        lable_fac.setText("Факультет/ННІ")

        self.__input_fac = QComboBox(self)
        self.__input_fac.setObjectName("fac")
        self.__input_fac.addItems([el["name"] for el in self.__facs])
        self.__input_fac.addItem("")
        self.__input_fac.activated.connect(
            functools.partial(self.text_changed, self.__input_fac, True)
        )

        layout_fac = QHBoxLayout()
        layout_fac.addWidget(lable_fac)
        layout_fac.addWidget(self.__input_fac)

        lable_teacher = QLabel(self)
        lable_teacher.setText("ПІБ")

        self.__input_teacher = QLineEdit(self)
        self.__input_teacher.setObjectName("teacher")
        self.__input_teacher.textChanged.connect(
            functools.partial(self.text_changed, self.__input_teacher, False)
        )

        layout_teacher = QHBoxLayout()
        layout_teacher.addWidget(lable_teacher)
        layout_teacher.addWidget(self.__input_teacher)

        lable_gram = QLabel(self)
        lable_gram.setText("Нагорода КПІ")

        self.__input_gram = QComboBox(self)
        self.__input_gram.setObjectName("gram")
        self.__input_gram.addItems([el["name"] for el in self.__grams])
        self.__input_gram.addItem("")
        self.__input_gram.activated.connect(
            functools.partial(self.text_changed, self.__input_gram, True)
        )

        layout_gram = QHBoxLayout()
        layout_gram.addWidget(lable_gram)
        layout_gram.addWidget(self.__input_gram)

        lable_state_gram = QLabel(self)
        lable_state_gram.setText("Державна Нагорода")

        self.__input_state_gram = QComboBox(self)
        self.__input_state_gram.setObjectName("state_gram")
        self.__input_state_gram.addItems(
            [el["name"] for el in self.__state_grams]
        )
        self.__input_state_gram.addItem("")
        self.__input_state_gram.activated.connect(
            functools.partial(
                self.text_changed, 
                self.__input_state_gram, 
                True
            )
        )

        layout_state_gram = QHBoxLayout()
        layout_state_gram.addWidget(lable_state_gram)
        layout_state_gram.addWidget(self.__input_state_gram)

        lable_num_prot = QLabel(self)
        lable_num_prot.setText("№ протоколу ВР")

        self.__input_num_prot = QLineEdit(self)
        self.__input_num_prot.setObjectName("num")
        self.__input_num_prot.textChanged.connect(
            functools.partial(self.text_changed, self.__input_num_prot, False)
        )

        layout_num_prot = QHBoxLayout()
        layout_num_prot.addWidget(lable_num_prot)
        layout_num_prot.addWidget(self.__input_num_prot)

        lable_year = QLabel(self)
        lable_year.setText("Рік відзначення КПІ")

        self.__input_year = QLineEdit(self)
        self.__input_year.setObjectName("year")
        self.__input_year.textChanged.connect(
            functools.partial(self.text_changed, self.__input_year, False)
        )

        layout_year = QHBoxLayout()
        layout_year.addWidget(lable_year)
        layout_year.addWidget(self.__input_year)

        lable_state_year = QLabel(self)
        lable_state_year.setText("Рік відзначення державою")

        self.__input_state_year = QLineEdit(self)
        self.__input_state_year.setObjectName("state_year")
        self.__input_state_year.textChanged.connect(
            functools.partial(
                self.text_changed, 
                self.__input_state_year, 
                False
            )
        )

        layout_state_year = QHBoxLayout()
        layout_state_year.addWidget(lable_state_year)
        layout_state_year.addWidget(self.__input_state_year)

        lable_prog = QLabel(self)
        lable_prog.setText("Прогнозування")

        self.__input_prog = QLineEdit(self)
        self.__input_prog.setObjectName("prog")
        self.__input_prog.textChanged.connect(
            functools.partial(self.text_changed, self.__input_prog, False)
        )

        layout_prog = QHBoxLayout()
        layout_prog.addWidget(lable_prog)
        layout_prog.addWidget(self.__input_prog)

        layout_input = QVBoxLayout()
        layout_input.addLayout(layout_fac)
        layout_input.addLayout(layout_teacher)
        layout_input.addLayout(layout_gram)
        layout_input.addLayout(layout_state_gram)
        layout_input.addLayout(layout_num_prot)
        layout_input.addLayout(layout_year)
        layout_input.addLayout(layout_state_year)
        layout_input.addLayout(layout_prog)

        self.__dataviewer = QListWidget(self)
        self.__dataviewer.setStyleSheet(
            "QListWidget {"
                "background-color: whitesmoke;"
                "border: 1px solid black;"
                "max-width: 600px;"
                "min-width: 300px;"
                "max-height: 900px;"
                "margin: 5px 5px 5px 50px;"
            "}"
        )

        datalabel = QLabel(self)
        datalabel.setText("Формат, у якому зберігаються значення в БД")

        self.__statusviewer = QListWidget(self)
        self.__statusviewer.setStyleSheet(
            "QListWidget {"
                "background-color: whitesmoke;"
                "border: 1px solid black;"
                "max-width: 600px;"
                "min-width: 300px;"
                "max-height: 100px;"
                "margin: 5px 5px 5px 50px;"
            "}"
        )

        statuslabel = QLabel(self)
        statuslabel.setText("Статус внесення в БД")

        layout_viewer = QVBoxLayout()
        layout_viewer.addWidget(
            datalabel, 
            alignment=(
                Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignCenter
            )
        )
        layout_viewer.addWidget(self.__dataviewer)
        layout_viewer.addWidget(
            statuslabel, 
            alignment=(
                Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignCenter
            )
        )
        layout_viewer.addWidget(self.__statusviewer)

        layout_container = QHBoxLayout()
        layout_container.addLayout(layout_input)
        layout_container.addLayout(layout_viewer)

        self.setStyleSheet(
            "QWidget {"
                "margin: 5px;"
            "}"
            "QLineEdit, QComboBox {"
                "max-width: 300px;"
                "max-height: 20px;"
                "min-width: 200px;"
                "min-height: 20px;"
                "margin: 2px 50px 2px 2px;"
            "}"
            "QLabel {"
                "max-width: 400px;"
                "max-height: 40px;"
                "min-width: 200px;"
                "min-height: 20px;"
                "margin: 2px;"
            "}"
            "QListWidget::item {"
                "background-color: #82ccdd;"
                "border: 1px solid grey;"
                "border-radius: 2px;"
                "margin: 2px;"
                "padding: 2px;"
            "}"
        )
        self.setLayout(layout_container)

    def __set_data(self):
        return {
            "fac": "",
            "teacher": "",
            "gram": "",
            "state_gram": "",
            "num": "",
            "year": "",
            "state_year": "",
            "prog": ""
        }

    def get_data(self):
        return self.__data

    def text_changed(
        self, 
        item: typing.Union[QLineEdit, QComboBox], 
        check: bool = False
    ):
        name = item.objectName()

        if check:
            self.__data[name] = item.currentText().strip()
        else:
            self.__data[name] = item.text().strip()

        output = str(self.__data)
        output = re.sub("'", '"', output)
        output = re.sub(",", ",\n", output)

        self.__dataviewer.clear()
        self.__dataviewer.addItem(output)

    def clear_data(self, all):
        if all:
            self.__dataviewer.clear()
            self.__statusviewer.clear()

        self.__input_fac.setCurrentText("")
        self.__input_teacher.clear()
        self.__input_gram.setCurrentText("")
        self.__input_state_gram.setCurrentText("")
        self.__input_num_prot.clear()
        self.__input_year.clear()
        self.__input_state_year.clear()
        self.__input_prog.clear()

    def show_status(self, msg: str, color: str):
        self.__statusviewer.setStyleSheet(
            "QListWidget::item {"
                f"background-color: {color};"
            "}")
        self.__statusviewer.addItem(msg)
        
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.__db = DataBase("localhost", 27017, "teacher_awards")

        self.setWindowTitle("База даних нагород та подяк")
        self.setMinimumSize(QSize(1000, 500))

        self.set_menubar()
        self.set_mainmenu()

    def set_mainmenu(self):
        self.__table = Table(self)
        self.__filters = Filters(self)
        maindata = QWidget(self)

        btn_set_filters = QPushButton("Почати пошук")
        btn_set_filters.setStyleSheet("margin: 0% 0% 5% 35%")
        btn_set_filters.clicked.connect(self.search_data)

        btn_clear_filters = QPushButton("Збросити фільтри")
        btn_clear_filters.setStyleSheet("margin: 0% 35% 5% 0%")
        btn_clear_filters.clicked.connect(self.__filters.clear_filters)

        btn_show_data = QPushButton("Вивести всі дані")
        btn_show_data.setStyleSheet("margin: 0% auto 5%")
        btn_show_data.clicked.connect(
            functools.partial(
                self.__table.show_data, 
                self.__db.get_teachers()
            )
        )

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(btn_set_filters)
        btn_layout.addWidget(btn_clear_filters)

        vlayout = QVBoxLayout()
        vlayout.addWidget(self.__filters)
        vlayout.addLayout(btn_layout)
        vlayout.addWidget(btn_show_data)

        hlayout = QHBoxLayout()
        hlayout.addWidget(self.__table)
        hlayout.addLayout(vlayout)


        layout = QVBoxLayout()
        layout.addLayout(hlayout)

        maindata.setLayout(layout)
        self.setCentralWidget(maindata)

    def set_menubar(self):
        menu = self.menuBar()
        imp = menu.addMenu("Імпортувати дані")
        exp = menu.addMenu("Експортувати дані")
        save = menu.addMenu("Вилучити виведені дані")

        imp_xlsx = QAction("*.xlsx", self)
        imp_xlsx.setData({"imp": "xlsx"})
        imp_xlsx.triggered.connect(
            functools.partial(self.imp_data, imp_xlsx)
        )

        imp_csv = QAction("*.csv", self)
        imp_csv.setData({"imp": "csv"})
        imp_csv.triggered.connect(
            functools.partial(self.imp_data, imp_csv)
        )

        imp_by_hand = QAction("Ввести вручну", self)
        imp_by_hand.triggered.connect(self.set_insertmenu)

        imp.addAction(imp_xlsx)
        imp.addSeparator()
        imp.addAction(imp_csv)
        imp.addSeparator()
        imp.addAction(imp_by_hand)

        exp_xlsx = QAction("*.xlsx", self)
        exp_xlsx.setData({"exp": "xlsx"})
        exp_xlsx.triggered.connect(
            functools.partial(self.exp_data, exp_xlsx)
        )

        exp_csv = QAction("*.csv", self)
        exp_csv.setData({"exp": "csv"})
        exp_csv.triggered.connect(
            functools.partial(self.exp_data, exp_csv)
        )

        exp.addAction(exp_xlsx)
        exp.addSeparator()
        exp.addAction(exp_csv)

        save_xlsx = QAction("*.xlsx", self)
        save_xlsx.setData({"exp": "xlsx"})
        save_xlsx.triggered.connect(
            functools.partial(self.save_data, save_xlsx)
        )
        save_csv = QAction("*.csv", self)
        save_csv.setData({"exp": "csv"})
        save_csv.triggered.connect(
            functools.partial(self.save_data, save_csv)
        )

        save.addAction(save_xlsx)
        save.addSeparator()
        save.addAction(save_csv)

    def imp_data(self, item: QAction):
        dialog = QFileDialog(self)
        url = dialog.getOpenFileName()[0]
        data = item.data()

        try:
            if data["imp"] == "xlsx":
                df = pd.read_excel(url)
            elif data["imp"] == "csv":
                df = pd.read_csv(url)
            else:
                raise ValueError("Can't parse file (no such format)")

            data = []

            for i in range(len(df)):
                teacher = str(df.loc[i][1]).strip()
                fac = str(df.loc[i][2]).strip()
                gram = str(df.loc[i][3]).strip()
                state_gram = str(df.loc[i][4]).strip()
                num_prot = str(df.loc[i][5]).strip()
                year = str(df.loc[i][6]).strip()
                state_year = str(df.loc[i][7])
                prog = str(df.loc[i][8]).strip()

                if not self.__db.check_facs(fac):
                    raise KeyError(
                        f"У КПІ не існує такого факультету/ННІ: {fac}"
                    )

                if (not self.__db.check_kpi_awards(gram) or 
                    len(gram) == 0):
                    raise KeyError(
                        f"У КПІ не існує такої нагороди: {gram}"
                    )

                if (not self.__db.check_state_awards(state_gram) or 
                    len(state_gram) == 0):
                    raise KeyError(
                        f"Не існує такої державної нагороди: {state_gram}"
                    )

                gram = re.sub("(`|'|\")", "'", gram)
                state_gram = re.sub("(`|'|\")", "'", state_gram)
                prog = (self.__db.get_kpi_award_next({"name": gram}) or
                        self.__db.get_state_award_next({"name": state_gram}))

                try:
                    year = str(int(float(year)))
                except:
                    pass

                try:
                    state_year = str(int(float(state_year)))
                except:
                    pass

                data.append({
                    "teacher": teacher,
                    "fac": fac,
                    "gram": gram,
                    "state_gram": state_gram,
                    "num": num_prot,
                    "year": year,
                    "state_year": state_year,
                    "prog": prog
                })

            self.__db.set_and_rm_teachers(data)
            
        except Exception as ex_:
            logging.error(ex_)

            if len(url.strip()) > 0:
                msg = QMessageBox(self)
                msg.setBaseSize(700, 500)
                msg.setWindowTitle("Помилка!")
                msg.setText(
                    "Неможливо розпарсити файл. "
                    "Оберіть інший формат файла або "
                    "спробуйте перетягти таблицю у позицію A1 та "
                    "сформувати поля таким чином: №, "
                    "Прізвище, ім'я, по-батькові співробітника, "
                    "Факультет/ННІ, "
                    "Нагорода (Почесне звання, відзнака та грамота), "
                    "№ Протоколу ВР КПІ ім. Ігоря "
                    "Сікорського про відзнічення, "
                    "Рік відзначення КПІ, "
                    "Рік призначення державою, "
                    "Прогнозування."
                )
                msg.show()

            return

        self.__table.show_data(self.__db.get_teachers())

    def exp_data(self, item: QAction, teachers: typing.List[dict] = None):
        dialog = QFileDialog(self)
        url = dialog.getSaveFileName()[0]
        data = item.data()

        try:
            teachers = [el for el in (teachers or self.__db.get_teachers())]
        except:
            msg = QMessageBox(self)
            msg.setBaseSize(700, 500)
            msg.setWindowTitle("Помилка!")
            msg.setText(
                "Нема даних, які можна було експортувати."
            )
            msg.show()
        
        for el in teachers:
            el["gram"] = re.sub("'", "`", el["gram"])
            el["state_gram"] = re.sub("'", "`", el["state_gram"])
            try:
                el.pop("_id")
            except:
                pass
        
        teachers = re.sub("'", '"', str(teachers))

        try:
            df = pd.read_json(teachers)
            df.rename(columns={
                "teacher": "Прізвище, ім'я, по-батькові співробітника",
                "fac": "Факультет/ННІ",
                "gram": "Нагорода (Почесне звання, відзнака та грамота)",
                "state_gram": "Державна нагорода",
                "num": ("№ Протоколу ВР КПІ ім. Ігоря "
                        "Сікорського про відзнічення"),
                "year": "Рік відзначення КПІ",
                "state_year": "Рік призначення державою",
                "prog": "Прогнозування"
            }, inplace=True)

            if data["exp"] == "xlsx":
                df.to_excel(url)
            elif data["exp"] == "csv":
                df.to_csv(url)
            else:
                raise ValueError("Can't parse file (no such format)")
        except Exception as ex_:
            logging.error(ex_)

            if len(url.strip()) > 0:
                msg = QMessageBox(self)
                msg.setBaseSize(700, 500)
                msg.setWindowTitle("Помилка!")
                msg.setText(
                    "Неможливо сформувати файл. "
                    "Спробуйте обрати інше розширення, формат або "
                    "внесіть дані, щоб їх імпортувати."
                )
                msg.show()

            return

    def search_data(self):
        filters = {}
        filters = self.__filters.get_filters(filters)

        data = self.__db.get_teachers(filters)

        self.__table.show_data(data)

    def save_data(self, item):
        data = self.__table.temp_data
        self.exp_data(item, data)

    def set_insertmenu(self):
        self.__toinsert = WindowToInsertData(
            self, 
            self.__db.get_facs(),
            self.__db.get_kpi_awards(),
            self.__db.get_state_awards()
        )

        insert = QPushButton("Внести дані")
        insert.setStyleSheet("min-width: 150px")
        insert.clicked.connect(self.insert_into_db)

        remove = QPushButton("Стерти")
        remove.setStyleSheet("min-width: 150px")
        remove.clicked.connect(
            functools.partial(self.__toinsert.clear_data, True)
        )

        exit = QPushButton("Повернутися")
        exit.setStyleSheet("min-width: 300px")
        exit.clicked.connect(self.return_to_mainmenu)

        menu = self.menuBar()
        menu.clear()

        layout_row_btn = QHBoxLayout()
        layout_row_btn.addWidget(insert)
        layout_row_btn.addWidget(remove)

        layout_btn = QVBoxLayout()
        layout_btn.addLayout(layout_row_btn)
        layout_btn.addWidget(exit)

        container = QVBoxLayout()
        container.addWidget(self.__toinsert)
        container.addLayout(layout_btn)

        data = QWidget(self)
        data.setLayout(container)

        self.setCentralWidget(data)

    def insert_into_db(self):
        data = self.__toinsert.get_data()

        if len(data["teacher"]) == 0:
            self.__toinsert.show_status(
                "Помилка: потрібно внести ПІБ викладача",
                "red"
            )
            return

        l_gram = len(data["gram"])
        l_sgram = len(data["state_gram"])

        if l_gram > 0 and l_sgram > 0:
            self.__toinsert.show_status(
                "Помилка: можна вносити лише одну нагороду",
                "red"
            )
            return
        elif l_gram <= 0 and l_sgram <= 0:
            self.__toinsert.show_status(
                "Помилка: потрібно внести нагороду",
                "red"
            )

        l_year = len(data["year"])
        l_syear = len(data["state_year"])

        if l_year > 0 and l_syear > 0:
            self.__toinsert.show_status(
                "Помилка: можна вносити лише один рік",
                "red"
            )
            return
        elif l_year <= 0 and l_syear <= 0:
            self.__toinsert.show_status(
                "Помилка: потрібно внести рік",
                "red"
            )
            return

        if ((l_gram > 0 and l_syear > 0) or
            (l_sgram > 0 and l_year > 0)):
            self.__toinsert.show_status(
                "Помилка: потрібно внести нагороду й рік її отримання"
                "або від КПІ, або від держави",
                "red"
            )
            return

        try:
            self.__db.set_teacher(data)
        except Exception as ex_:
            self.__toinsert.show_status(ex_.args, "red")
        else:
            self.__toinsert.show_status("Дані були внесені", "green")

        self.__toinsert.clear_data(False)

    def return_to_mainmenu(self):
        self.set_menubar()
        self.set_mainmenu()

if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())