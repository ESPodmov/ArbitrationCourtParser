import asyncio
import os
import re
import sys

from PyQt5 import QtGui
from PyQt5.QtGui import QIcon, QCloseEvent
from PyQt5.QtWidgets import QWidget, QLabel, QComboBox, QPushButton, QGridLayout, QDateEdit, \
    QFormLayout, QVBoxLayout, QAbstractSpinBox, QLineEdit, QMessageBox, QFileDialog, QTextEdit, QHBoxLayout, \
    QSizePolicy, QDesktopWidget
from PyQt5.QtCore import QDate, Qt, QFile, QTextStream, QMargins

import NewParser
import configparser
from docx_editor import create_or_edit_docx_from_list

court_list = ["", "АС Республики Татарстан", "АС города Москвы", "АС города Санкт-Петербурга и Ленинградской области",
              "Верховный Суд РФ", "Высший Арбитражный Суд РФ", "АС Волго-Вятского округа",
              "АС Восточно-Сибирского округа", "АС Дальневосточного округа", "АС Западно-Сибирского округа",
              "АС Московского округа", "АС Поволжского округа", "АС Северо-Западного округа",
              "АС Северо-Кавказского округа", "АС Уральского округа", "АС Центрального округа",
              "1 арбитражный апелляционный суд", "2 арбитражный апелляционный суд", "3 арбитражный апелляционный суд",
              "4 арбитражный апелляционный суд", "5 арбитражный апелляционный суд", "6 арбитражный апелляционный суд",
              "7 арбитражный апелляционный суд", "8 арбитражный апелляционный суд", "9 арбитражный апелляционный суд",
              "10 арбитражный апелляционный суд", "11 арбитражный апелляционный суд",
              "12 арбитражный апелляционный суд", "13 арбитражный апелляционный суд",
              "14 арбитражный апелляционный суд", "15 арбитражный апелляционный суд",
              "16 арбитражный апелляционный суд", "17 арбитражный апелляционный суд",
              "18 арбитражный апелляционный суд", "19 арбитражный апелляционный суд",
              "20 арбитражный апелляционный суд", "21 арбитражный апелляционный суд", "АС Алтайского края",
              "АС Амурской области", "АС Архангельской области", "АС Астраханской области", "АС Белгородской области",
              "АС Брянской области", "АС Владимирской области", "АС Волгоградской области", "АС Вологодской области",
              "АС Воронежской области", "АС города Севастополя", "АС Донецкой Народной Республики",
              "АС Еврейской автономной области", "АС Забайкальского края", "АС Запорожской области",
              "АС Ивановской области", "АС Иркутской области", "АС Кабардино-Балкарской Республики",
              "АС Калининградской области", "АС Калужской области", "АС Камчатского края",
              "АС Карачаево-Черкесской Республики", "АС Кемеровской области", "АС Кировской области",
              "АС Коми-Пермяцкого АО", "АС Костромской области", "АС Краснодарского края", "АС Красноярского края",
              "АС Курганской области", "АС Курской области", "АС Липецкой области", "АС Луганской Народной Республики",
              "АС Магаданской области", "АС Московской области", "АС Мурманской области", "АС Нижегородской области",
              "АС Новгородской области", "АС Новосибирской области", "АС Омской области", "АС Оренбургской области",
              "АС Орловской области", "АС Пензенской области", "АС Пермского края", "АС Приморского края",
              "АС Псковской области", "АС Республики Адыгея", "АС Республики Алтай", "АС Республики Башкортостан",
              "АС Республики Бурятия", "АС Республики Дагестан", "АС Республики Ингушетия", "АС Республики Калмыкия",
              "АС Республики Карелия", "АС Республики Коми", "АС Республики Крым", "АС Республики Марий Эл",
              "АС Республики Мордовия", "АС Республики Саха", "АС Республики Северная Осетия", "АС Республики Тыва",
              "АС Республики Хакасия", "АС Ростовской области", "АС Рязанской области", "АС Самарской области",
              "АС Саратовской области", "АС Сахалинской области", "АС Свердловской области", "АС Смоленской области",
              "АС Ставропольского края", "АС Тамбовской области", "АС Тверской области", "АС Томской области",
              "АС Тульской области", "АС Тюменской области", "АС Удмуртской Республики", "АС Ульяновской области",
              "АС Хабаровского края", "АС Ханты-Мансийского АО", "АС Херсонской области", "АС Челябинской области",
              "АС Чеченской Республики", "АС Чувашской Республики", "АС Чукотского АО", "АС Ямало-Ненецкого АО",
              "АС Ярославской области", "ПСП Арбитражного суда Пермского края",
              "ПСП Арбитражный суд Архангельской области", "Суд по интеллектуальным правам"]


def format_dictionary_to_string(dictionary: dict, format_line):
    case_dict = {1: "Административное", 2: "Гражданское", 3: "Банкротство"}
    case_date = dictionary["case_date"]
    case_type = case_dict[dictionary["case_type"]]
    case_num = dictionary["case_num"]
    case_link = dictionary["case_link"]
    case_judge = dictionary["case_judge"]
    case_court = dictionary["cas_court"]
    plaintiffs_text, respondents_text = get_plaintiffs_and_respondents_text(dictionary)

    return_line = format_line.format(case_num=case_num, case_link=case_link, case_date=case_date, case_type=case_type,
                                     plaintiffs=plaintiffs_text, respondents=respondents_text, case_court=case_court,
                                     case_judge=case_judge) + "\n"
    return return_line


def create_or_edit_docx_file(path, my_list, format_line):
    case_dict = {1: "Административное", 2: "Гражданское", 3: "Банкротство", 0: "Не определено"}
    link_dict = {"case_link": my_list[0]["case_link"]}

    for dictionary in my_list:
        dictionary["case_type"] = case_dict[dictionary["case_type"]]
        plaintiffs_text, respondents_text = get_plaintiffs_and_respondents_text(dictionary)
        dictionary["plaintiffs"] = plaintiffs_text
        dictionary["respondents"] = respondents_text

    create_or_edit_docx_from_list(path, my_list, format_line, link_dict)


def get_plaintiffs_and_respondents_text(dictionary):
    plaintiffs = dictionary["plaintiff"]
    respondents = dictionary["respondent"]
    if len(plaintiffs) == 0:
        plaintiffs_text = "Истец: нет информации"
    else:
        plaintiffs_text = ""
        for elem in plaintiffs:
            current_line = "Истец: "
            current_line += elem["plaintiff_name"] + "\n"
            current_line += elem["plaintiff_address"]
            plaintiffs_text += current_line
    if len(respondents) == 0:
        respondents_text = "Ответчик: нет информации"
    else:
        respondents_text = ""
        for elem in respondents:
            current_line = "Ответчик: "
            current_line += elem["respondent_name"] + "\n"
            current_line += elem["respondent_address"]
            respondents_text += current_line
    return plaintiffs_text, respondents_text


def format_list_to_text(current_list: list, format_line):
    return_line = ""
    for elem in current_list:
        return_line += format_dictionary_to_string(elem, format_line=format_line)
    return return_line


class MainWindow(QWidget):
    def __init__(self, cwd):
        super().__init__()
        self.cwd = cwd
        self.new_window = None
        center_point = QDesktopWidget().screenGeometry().center()
        self.setGeometry(center_point.x() - int(500 / 2), center_point.y() - int(475 / 2), 500, 475)
        styleFile = QFile(f"{self.cwd}/style.css")
        styleFile.open(QFile.ReadOnly | QFile.Text)
        stream = QTextStream(styleFile)
        self.setStyleSheet("MainWindow {background-color: #fff;}\n" + stream.readAll())

        self.setWindowTitle('Парсер')
        self.setWindowIcon(QIcon("icon.ico"))

        label1 = QLabel('Старт поиска', self)
        label2 = QLabel('Конец поиска', self)
        label3 = QLabel('Браузер', self)
        label4 = QLabel('Тип дел', self)
        label5 = QLabel('Суд', self)

        size_policy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        size_policy.setHorizontalStretch(1)

        self.dateEdit1 = QDateEdit(self)
        self.dateEdit1.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self.dateEdit1.setDisplayFormat("dd.MM.yyyy")
        self.dateEdit1.setSizePolicy(size_policy)

        # Устанавливаем текущую дату в первое поле ввода
        self.dateEdit2 = QDateEdit(self)
        self.dateEdit2.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self.dateEdit2.setDate(QDate.currentDate())
        self.dateEdit2.setDisplayFormat("dd.MM.yyyy")
        self.dateEdit2.setSizePolicy(size_policy)

        # Создаем выпадающее меню для выбора браузера
        self.browserComboBox = QComboBox(self)
        self.browserComboBox.addItem('Chrome')
        self.browserComboBox.addItem('Dolphin')
        self.browserComboBox.setSizePolicy(size_policy)

        self.actTypeComboBox = QComboBox(self)
        self.actTypeComboBox.addItem("")
        self.actTypeComboBox.addItem("Административные")
        self.actTypeComboBox.addItem("Гражданские")
        self.actTypeComboBox.addItem("Банкротные")
        self.actTypeComboBox.setSizePolicy(size_policy)

        self.courtComboBox = QComboBox(self)
        self.courtComboBox.addItems(court_list)
        self.courtComboBox.setSizePolicy(size_policy)

        self.button = QPushButton('Запустить', self)
        self.button.clicked.connect(self.runScript)

        form = QGridLayout()
        form.addWidget(label1, 0, 0)
        form.addWidget(self.dateEdit1, 0, 1)
        form.addWidget(label2, 1, 0)
        form.addWidget(self.dateEdit2, 1, 1)
        form.addWidget(label3, 2, 0)
        form.addWidget(self.browserComboBox, 2, 1)
        form.addWidget(label4, 3, 0)
        form.addWidget(self.actTypeComboBox, 3, 1)
        form.addWidget(label5, 4, 0)
        form.addWidget(self.courtComboBox, 4, 1)
        form.setContentsMargins(0, 7, 0, 0)

        file_choose_label = QLabel("Путь к файлу")
        config = configparser.ConfigParser()
        config.read(f"{self.cwd}/config.ini", encoding="utf-8")
        self.file_choose_line = QLineEdit()
        self.file_choose_line.setText(f"{config.get('global', 'directory')}/file.docx")
        self.file_choose_line.setSizePolicy(size_policy)
        self.file_choose_line.setContentsMargins(0, 0, 0, 0)
        self.file_choose_btn = QPushButton("Обзор")
        self.file_choose_btn.clicked.connect(self.open_select)
        g_lay = QGridLayout()
        g_lay.addWidget(file_choose_label, 0, 0)
        g_lay.addWidget(self.file_choose_line, 0, 1)
        g_lay.addWidget(self.file_choose_btn, 0, 2)
        g_lay.setContentsMargins(0, 7, 0, 0)

        m_l = QVBoxLayout()
        layout = QGridLayout()
        layout.addWidget(self.button, 0, 0, 1, 1, alignment=Qt.AlignCenter)

        change_link = QLabel("<a href='f'>Изменить данные</a>")
        change_link.setOpenExternalLinks(False)
        change_link.linkActivated.connect(self.edit_data)
        change_link.setStyleSheet("""
        font-size: 13px;
        """)
        layout.addWidget(change_link, 1, 0, 1, 1, alignment=Qt.AlignCenter)
        m_l.addLayout(layout)
        m_l.setAlignment(Qt.AlignBottom)

        main_layout = QVBoxLayout()
        main_layout.addLayout(form)
        main_layout.addLayout(g_lay)
        main_layout.addLayout(m_l)

        self.setLayout(main_layout)

    def runScript(self):
        file_path = self.file_choose_line.text()
        # if self.dateEdit1.text() > self.dateEdit2.text():
        #     alert = QMessageBox()
        #     alert.setIcon(QMessageBox.Critical)
        #     alert.setWindowTitle("Ошибка")
        #     alert.setText("Начальная дата больше конечной!")
        #     alert.exec_()
        #     return
        # el
        if file_path == "":
            alert = QMessageBox()
            alert.setIcon(QMessageBox.Critical)
            alert.setWindowTitle("Ошибка")
            alert.setText("Пустое имя файла")
            alert.exec_()
            return
        elif not ((sys.platform == "win32" or sys.platform == "win64") and
                  re.match(r'^([A-Za-z]:)?(\/[^\/\n]+)+\.docx$', file_path) or
                  re.match(r'^\/(?:[^\/\n]+\/)*[^\/\n]+\.docx$', file_path)):
            alert = QMessageBox()
            alert.setIcon(QMessageBox.Critical)
            alert.setWindowTitle("Ошибка")
            alert.setText("Неправильный путь к файлу\nУбедитесь что расширение файла .docx")
            alert.exec_()
            return
        elif not (os.access(file_path[:file_path.rfind("/")], os.F_OK) and os.access(file_path[:file_path.rfind("/")],
                                                                                     os.R_OK)):
            alert = QMessageBox()
            alert.setIcon(QMessageBox.Critical)
            alert.setWindowTitle("Ошибка")
            alert.setText(
                f"Какой-либо папки в пути к файлу не существует, создайте требуемые папки")
            alert.exec_()
            return
        start_date = self.dateEdit1.text()
        end_date = self.dateEdit2.text()
        start_date.replace(".", "")
        end_date.replace(".", "")
        browser = self.browserComboBox.currentText()
        act_type = self.get_act_type()
        court = self.courtComboBox.currentText()
        loop = asyncio.get_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(self.my_async(start_date, end_date, browser, act_type, court))

    async def my_async(self, start_date, end_date, browser, act_type, court):
        self.setDisabled(True)
        config = configparser.ConfigParser()
        config.read(f"{self.cwd}/config.ini", encoding="utf-8")
        parser = NewParser.Parser(browser, config.get("dolphin", "token"), float(config.get("global", "pause")),
                                  config.get("global", "chrome_path"))
        file_dir = self.file_choose_line.text()
        file_dir = file_dir[:file_dir.rfind("/")]
        config.set("global", "directory", file_dir)
        with open(f"{self.cwd}/config.ini", "w", encoding="utf-8") as file:
            config.write(file)
        main_list = []
        errored = False
        try:
            main_list = await parser.start_parse("https://kad.arbitr.ru/", start_date, end_date, act_type, court)
        except Exception as e:
            print(e)
            await asyncio.get_event_loop().shutdown_asyncgens()
            self.setDisabled(False)
            alert = QMessageBox()
            if len(main_list) == 0:
                alert.setWindowTitle("Ошибка")
                alert.setIcon(QMessageBox.Critical)
                alert.setText("Что-то пошло не так, перезапустите программу")
            else:
                alert.setWindowTitle("Внимание")
                alert.setIcon(QMessageBox.Warning)
                alert.setText("Что-то пошло не так, но файл с некотрыми данными сохранился")
            errored = True
            alert.exec_()
        self.setDisabled(False)
        # with open(self.file_choose_line.text(), mode="w", encoding='utf-8') as file:
        #     file.write(format_list_to_text(main_list, config.get("global", "current_text_template")))
        #     # file.write(main_string)
        create_or_edit_docx_file(self.file_choose_line.text(), main_list, config.get("global", "current_text_template"))
        if not errored:
            alert = QMessageBox()
            alert.setIcon(QMessageBox.Information)
            alert.setWindowTitle("Успех")
            alert.setText("Ваш файл сохранен")
            alert.exec_()

    def open_select(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_dir = self.file_choose_line.text()
        file_dir = file_dir[:file_dir.rfind("/")]
        try:
            file_name, _ = QFileDialog.getSaveFileName(None, "Save File", file_dir,
                                                       "All Files (*);;Word Files (*.docx)",
                                                       options=options)
        except Exception as e:
            print(e.with_traceback(e.__traceback__))
            file_name, _ = QFileDialog.getSaveFileName(None, "Save File", "", "All Files (*);;Word Files (*.docx)",
                                                       options=options)
        if file_name:
            self.file_choose_line.setText(file_name)

    def edit_data(self):
        if self.new_window is None:
            self.new_window = ChangeDataWindow(self.cwd, self.on_close)
            self.new_window.show()

    def on_close(self):
        self.new_window = None

    def get_act_type(self):
        act_name = self.actTypeComboBox.currentText()
        action_dict = {"": 0, "Административные": 1, "Гражданские": 2, "Банкротные": 3}
        return action_dict.get(act_name)


class ChangeDataWindow(QWidget):
    def __init__(self, cwd, close_this_func):
        super().__init__()
        self.close_this_func = close_this_func
        self.cwd = cwd
        self.config = configparser.ConfigParser()
        self.config.read(f"{self.cwd}/config.ini", encoding="utf-8")
        self.setWindowTitle("Данные")
        self.setWindowIcon(QIcon("icon.ico"))
        center_point = QDesktopWidget().screenGeometry().center()
        self.setGeometry(center_point.x() - int(925 / 2), center_point.y() - int(700 / 2), 925, 700)
        label = QLabel("Пауза перед стартом")
        label2 = QLabel("Путь к Chrome")
        label3 = QLabel("Токен")
        label4 = QLabel("Шаблон текста")

        size_policy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        size_policy.setHorizontalStretch(1)

        self.token_field = QLineEdit(self.config.get("dolphin", "token"))
        self.token_field.setSizePolicy(size_policy)

        self.pause_field = QLineEdit(self.config.get("global", "pause"))
        self.pause_field.setSizePolicy(size_policy)

        self.chrome_field = QLineEdit(self.config.get("global", "chrome_path"))
        self.chrome_field.setSizePolicy(size_policy)

        self.text_template_field = QTextEdit()
        self.text_template_field.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.text_template_field.insertPlainText(self.config.get("global", "current_text_template"))

        form = QGridLayout()
        form.setContentsMargins(0, 7, 0, 0)
        form.setSpacing(15)
        form.setHorizontalSpacing(15)
        form.addWidget(label3, 0, 0)
        form.addWidget(self.token_field, 0, 1)
        form.addWidget(label, 1, 0)
        form.addWidget(self.pause_field, 1, 1)
        form.addWidget(label2, 2, 0)
        form.addWidget(self.chrome_field, 2, 1)
        form.addWidget(label4, 3, 0)
        form.addWidget(self.text_template_field, 3, 1)

        styleFile = QFile(f"{self.cwd}/style.css")
        styleFile.open(QFile.ReadOnly | QFile.Text)
        stream = QTextStream(styleFile)
        self.setStyleSheet("ChangeDataWindow {background-color: #fff;}\n" + stream.readAll())

        self.confirm_button = QPushButton("Изменить")
        self.confirm_button.clicked.connect(self.change_data)
        self.cancel_button = QPushButton("Отмена")
        self.cancel_button.clicked.connect(self.close_this)
        self.reset_template_text_button = QPushButton("Сбросить текст")
        self.reset_template_text_button.clicked.connect(self.reset_template)

        m_l = QVBoxLayout()
        layout = QGridLayout()
        layout.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.confirm_button, 0, 0, 1, 1, alignment=Qt.AlignRight)
        layout.addWidget(self.reset_template_text_button, 0, 1, 1, 1, alignment=Qt.AlignHCenter)
        layout.addWidget(self.cancel_button, 0, 2, 1, 1, alignment=Qt.AlignLeft)
        m_l.addLayout(layout)
        m_l.setAlignment(Qt.AlignBottom)

        main_layout = QVBoxLayout()
        main_layout.addLayout(form)
        main_layout.addLayout(m_l)

        self.setLayout(main_layout)

    def closeEvent(self, a0: QtGui.QCloseEvent) -> None:
        self.close_this_func()

    def change_data(self):
        pause = self.pause_field.text()
        if pause.__contains__(","):
            pause = pause.replace(",", ".")
        if pause.count(".") > 1:
            alert = QMessageBox()
            alert.setIcon(QMessageBox.Critical)
            alert.setWindowTitle("Ошибка")
            alert.setText("Введите целое либо десятичное число")
            alert.exec_()
            return
        else:
            self.config.set("global", "pause", self.pause_field.text())
        self.config.set("dolphin", "token", self.token_field.text())
        self.config.set("global", "chrome_path", self.chrome_field.text())
        self.config.set("global", "current_text_template", self.text_template_field.toPlainText())
        self.close()
        with open(f"{self.cwd}/config.ini", "w", encoding="utf-8") as file:
            self.config.write(file)

    def close_this(self):
        self.close()

    def reset_template(self):
        self.text_template_field.setText(self.config.get("global", "default_text_template"))
