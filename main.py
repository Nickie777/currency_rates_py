import sys
import os
import requests
import xml.etree.ElementTree as ET
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QLineEdit, QTextEdit, QVBoxLayout, QFileDialog, QTimeEdit, QHBoxLayout, QDateEdit
from PyQt5.QtCore import QTime, QDate
import pandas as pd
import schedule
import time
import threading
from datetime import datetime, timedelta

class CurrencyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # Выбор папки
        self.folder_label = QLabel('Папка для сохранения:', self)
        self.folder_path = QLineEdit(self)
        self.folder_button = QPushButton('Выбрать папку', self)
        self.folder_button.clicked.connect(self.choose_folder)

        # Время для срабатывания
        self.time_label = QLabel('Время выполнения:', self)
        self.time_input = QTimeEdit(self)
        self.time_input.setTime(QTime.currentTime())

        # Даты для загрузки за период
        self.start_date_label = QLabel('Начало периода:', self)
        self.start_date_input = QDateEdit(self)
        self.start_date_input.setDate(QDate.currentDate())

        self.end_date_label = QLabel('Конец периода:', self)
        self.end_date_input = QDateEdit(self)
        self.end_date_input.setDate(QDate.currentDate())

        # Кнопка для загрузки за период
        self.load_period_button = QPushButton('Загрузить за период', self)
        self.load_period_button.clicked.connect(self.load_missing_for_period)

        # Лог
        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)

        # Кнопка "Запустить"
        self.run_button = QPushButton('Запустить', self)
        self.run_button.clicked.connect(self.start_scheduled_task)

        # Layouts
        vbox = QVBoxLayout()
        hbox_folder = QHBoxLayout()
        hbox_period = QHBoxLayout()

        hbox_folder.addWidget(self.folder_label)
        hbox_folder.addWidget(self.folder_path)
        hbox_folder.addWidget(self.folder_button)

        hbox_period.addWidget(self.start_date_label)
        hbox_period.addWidget(self.start_date_input)
        hbox_period.addWidget(self.end_date_label)
        hbox_period.addWidget(self.end_date_input)
        hbox_period.addWidget(self.load_period_button)

        vbox.addLayout(hbox_folder)
        vbox.addWidget(self.time_label)
        vbox.addWidget(self.time_input)
        vbox.addLayout(hbox_period)
        vbox.addWidget(self.run_button)
        vbox.addWidget(self.log_text)

        self.setLayout(vbox)
        self.setWindowTitle('Currency Fetcher')
        self.setGeometry(300, 300, 600, 400)

    def choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Выберите папку')
        if folder:
            self.folder_path.setText(folder)
            self.log('Папка выбрана: ' + folder)

    def start_scheduled_task(self):
        folder = self.folder_path.text()
        if not folder:
            self.log('Пожалуйста, выберите папку для сохранения.')
            return

        # Запланировать задачу
        selected_time = self.time_input.time().toString("HH:mm")
        schedule.every().day.at(selected_time).do(self.fetch_and_save_data, folder)

        self.log(f'Задача запланирована на {selected_time} каждый день.')

        # Запуск в отдельном потоке
        threading.Thread(target=self.run_scheduler, daemon=True).start()

    def run_scheduler(self):
        while True:
            schedule.run_pending()
            time.sleep(1)

    def fetch_and_save_data(self, folder, date_str=None):
        url = 'https://www.cbr.ru/scripts/XML_daily.asp'
        if not date_str:
            date_str = time.strftime('%d/%m/%Y')  # Для API используем формат дд/мм/гггг

        # Преобразуем дату для названия файла
        date_for_filename = datetime.strptime(date_str, '%d/%m/%Y').strftime('%Y-%m-%d')

        params = {'date_req': date_str}

        try:
            response = requests.get(url, params=params)
            response.raise_for_status()

            # Парсинг XML
            xml_root = ET.fromstring(response.content)

            data = []
            for valute in xml_root.findall('Valute'):
                char_code = valute.find('CharCode').text
                name = valute.find('Name').text
                value = valute.find('Value').text
                data.append([char_code, name, value])

            # Сохранение в Excel
            file_name = os.path.join(folder, f'currency_rates_{date_for_filename}.xlsx')
            df = pd.DataFrame(data, columns=['Код валюты', 'Название', 'Курс'])
            df.to_excel(file_name, index=False)

            self.log(f'Данные за {date_str} успешно сохранены в файл: {file_name}')

        except Exception as e:
            self.log(f'Ошибка при получении данных за {date_str}: {str(e)}')

    def load_missing_for_period(self):
        folder = self.folder_path.text()
        if not folder:
            self.log('Пожалуйста, выберите папку для сохранения.')
            return

        start_date = self.start_date_input.date().toPyDate()
        end_date = self.end_date_input.date().toPyDate()

        current_date = start_date
        while current_date <= end_date:
            file_name = os.path.join(folder, f'currency_rates_{current_date.strftime("%Y-%m-%d")}.xlsx')
            if not os.path.exists(file_name):
                self.log(f'Файл за {current_date.strftime("%Y-%m-%d")} отсутствует, загружаем...')
                self.fetch_and_save_data(folder, current_date.strftime('%d/%m/%Y'))
            else:
                self.log(f'Файл за {current_date.strftime("%Y-%m-%d")} уже существует.')

            current_date += timedelta(days=1)

    def log(self, message):
        self.log_text.append(message)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CurrencyApp()
    ex.show()
    sys.exit(app.exec_())
