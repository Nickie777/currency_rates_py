import sys
import os
import requests
import xml.etree.ElementTree as ET
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QLineEdit, QTextEdit, QVBoxLayout, QFileDialog, QTimeEdit, QHBoxLayout
from PyQt5.QtCore import QTime, QTimer
import pandas as pd
import schedule
import time
import threading

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

        # Лог
        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)

        # Кнопка "Запустить"
        self.run_button = QPushButton('Запустить', self)
        self.run_button.clicked.connect(self.start_scheduled_task)

        # Layouts
        vbox = QVBoxLayout()
        hbox = QHBoxLayout()

        hbox.addWidget(self.folder_label)
        hbox.addWidget(self.folder_path)
        hbox.addWidget(self.folder_button)

        vbox.addLayout(hbox)
        vbox.addWidget(self.time_label)
        vbox.addWidget(self.time_input)
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

    def fetch_and_save_data(self, folder):
        url = 'https://www.cbr.ru/scripts/XML_daily.asp'
        date_str = time.strftime('%d/%m/%Y')
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
            df = pd.DataFrame(data, columns=['Код валюты', 'Название', 'Курс'])
            file_name = os.path.join(folder, f'currency_rates_{time.strftime("%Y-%m-%d")}.xlsx')
            df.to_excel(file_name, index=False)

            self.log(f'Данные успешно сохранены в файл: {file_name}')

        except Exception as e:
            self.log(f'Ошибка при получении данных: {str(e)}')

    def log(self, message):
        self.log_text.append(message)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CurrencyApp()
    ex.show()
    sys.exit(app.exec_())
