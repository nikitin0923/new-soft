import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QTableView, QVBoxLayout, \
    QAbstractItemView, QWidget, QTabWidget
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem
import schedule
import logging
import io
from cryptography.fernet import Fernet
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from PyQt5.QtGui import QClipboard
from PyQt5.QtWidgets import QMessageBox, QInputDialog, QLineEdit, QDialog
from PyQt5.QtWidgets import QFontDialog, QColorDialog
from PyQt5.QtWidgets import QUndoStack, QUndoCommand
import threading
import unittest
import io
import pickle
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import seaborn as sns
import time
from pandas.api.types import is_string_dtype
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase
from cryptography.fernet import Fernet
import shutil
from pandas.plotting import scatter_matrix
import pdfkit


# Инициализация логгера
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Главное окно приложения
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.table_data = None  # Загруженные данные

        self.setWindowTitle("Программа")
        self.setGeometry(100, 100, 800, 600)

        # Создание основного виджета и компоновки
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Создание вкладок
        tab_widget = QTabWidget(self)
        layout.addWidget(tab_widget)

        # Вкладка загрузки таблиц
        load_tab = QWidget(self)
        tab_widget.addTab(load_tab, "Загрузка таблиц")
        load_layout = QVBoxLayout(load_tab)

        # Кнопка загрузки таблиц
        load_button = QPushButton("Загрузить таблицы", self)
        load_button.clicked.connect(self.load_tables)
        load_layout.addWidget(load_button)

        # Таблица для отображения данных
        self.table_view = QTableView(self)
        self.table_view.setEditTriggers(QAbstractItemView.DoubleClicked)
        load_layout.addWidget(self.table_view)

    # Загрузка таблиц
    def load_tables(self):
        file_names, _ = QFileDialog.getOpenFileNames(self, "Выберите файлы", "", "CSV (*.csv);;Excel (*.xlsx)")
        if file_names:
            try:
                for file_name in file_names:
                    if file_name.endswith(".csv"):
                        df = pd.read_csv(file_name)
                    elif file_name.endswith(".xlsx"):
                        df = pd.read_excel(file_name)
                    else:
                        raise ValueError("Неподдерживаемый формат файла")

                    # Добавление таблицы в QTableView
                    model = QStandardItemModel(df.shape[0], df.shape[1], self)
                    model.setHorizontalHeaderLabels(df.columns)
                    for row in range(df.shape[0]):
                        for col in range(df.shape[1]):
                            item = QStandardItem(str(df.iloc[row, col]))
                            item.setEditable(False)
                            model.setItem(row, col, item)
                    self.table_view.setModel(model)

                    # Сохранение загруженных данных
                    self.table_data = df

                logging.info("Таблицы успешно загружены")
                QMessageBox.information(self, "Успех", "Таблицы успешно загружены")
            except Exception as e:
                logging.error(f"Ошибка при загрузке таблиц: {str(e)}")
                QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке таблиц: {str(e)}")
# Редактирование таблиц
class CustomTableModel(QAbstractTableModel):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        self.data = data

    def rowCount(self, parent=None):
        return self.data.shape[0]

    def columnCount(self, parent=None):
        return self.data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            value = str(self.data.iloc[index.row(), index.column()])
            return value

    def setData(self, index, value, role=Qt.EditRole):
        if role == Qt.EditRole:
            self.data.iloc[index.row(), index.column()] = value
            return True
        return False

    def flags(self, index):
        return Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable

# Сравнение таблиц
def compare_tables(df1, df2, column):
    merged = pd.merge(df1, df2, on=column, how="outer", indicator=True)
    found = merged[merged["_merge"] == "both"].drop(columns="_merge")
    not_found = merged[merged["_merge"] != "both"].drop(columns="_merge")
    return found, not_found

# Создание вкладок по именам ячеек
def create_tabs_by_cell(df, cell_name):
    unique_values = df[cell_name].unique()
    tab_widget.clear()

    for value in unique_values:
        data = df[df[cell_name] == value].drop(columns=cell_name)
        model = CustomTableModel(data)
        view = QTableView()
        view.setModel(model)
        tab_widget.addTab(view, str(value))

# Создание отчета и статистики
def generate_report(dataframes, selected_columns):
    report_df = pd.DataFrame()
    for df in dataframes:
        if selected_columns:
            df = df[selected_columns]
        if not df.empty:
            statistics = df.describe()
            report_df = report_df.append(statistics)

    return report_df

# ... (весь предыдущий код остается без изменений) ...

# Добавление функциональности редактирования таблиц
model = CustomTableModel(self.table_data)
self.table_view.setModel(model)
self.table_view.doubleClicked.connect(self.edit_cell)

def edit_cell(self, index):
    if index.isValid():
        self.table_view.edit(index)

# Добавление функциональности сравнения таблиц
def compare_tables(self):
    if self.table_data is not None:
        column, ok = QInputDialog.getText(self, "Выберите столбец", "Введите имя столбца:")
        if ok:
            found, not_found = compare_tables(self.table_data, self.table_data, column)
            create_tabs_by_cell(found, column)
            QMessageBox.information(self, "Сравнение таблиц", "Таблицы успешно сравнены.")

# Добавление функциональности создания отчета и статистики
def generate_report(self):
    if self.table_data is not None:
        selected_columns, ok = QInputDialog.getText(self, "Выберите столбцы",
                                                    "Введите имена столбцов через запятую:")
        if ok:
            selected_columns = [col.strip() for col in selected_columns.split(",")]
            report_df = generate_report([self.table_data], selected_columns)
            save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить отчет", "", "CSV (*.csv);;Excel (*.xlsx)")
            if save_path:
                if save_path.endswith(".csv"):
                    report_df.to_csv(save_path, index=True)
                elif save_path.endswith(".xlsx"):
                    report_df.to_excel(save_path, index=True)
                QMessageBox.information(self, "Отчет сгенерирован", "Отчет успешно сгенерирован и сохранен.")

# ... (весь предыдущий код остается без изменений) ...

# Автоматизация
def run_scheduled_tasks():
    # Здесь добавьте вызовы функций, которые должны быть выполнены автоматически по расписанию
    pass

schedule.every().day.at("09:00").do(run_scheduled_tasks)
# Обработка ошибок и валидация данных
def validate_dataframe(df):
    if df is None or df.empty:
        raise ValueError("Нет доступных данных для обработки")

def validate_selected_columns(columns, dataframe_columns):
    for column in columns:
        if column not in dataframe_columns:
            raise ValueError(f"Столбец '{column}' не существует в таблице")

def validate_file_path(file_path):
    if not file_path:
        raise ValueError("Не выбран путь для сохранения файла")

# ... (весь предыдущий код остается без изменений) ...

# Создание отчета и статистики
def generate_report(self):
    try:
        validate_dataframe(self.table_data)

        selected_columns, ok = QInputDialog.getText(self, "Выберите столбцы", "Введите имена столбцов через запятую:")
        if ok:
            selected_columns = [col.strip() for col in selected_columns.split(",")]
            validate_selected_columns(selected_columns, self.table_data.columns)

            report_df = generate_report([self.table_data], selected_columns)

            save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить отчет", "", "CSV (*.csv);;Excel (*.xlsx)")
            validate_file_path(save_path)

            if save_path.endswith(".csv"):
                report_df.to_csv(save_path, index=True)
            elif save_path.endswith(".xlsx"):
                report_df.to_excel(save_path, index=True)

            QMessageBox.information(self, "Отчет сгенерирован", "Отчет успешно сгенерирован и сохранен.")
    except Exception as e:
        logging.error(f"Ошибка при генерации отчета: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при генерации отчета: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Запуск расписания
    while True:
        schedule.run_pending()
        app.processEvents()

    sys.exit(app.exec())
    # Добавление функциональности сохранения данных
def save_data(self):
    try:
        validate_dataframe(self.table_data)

        save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить данные", "", "CSV (*.csv);;Excel (*.xlsx)")
        validate_file_path(save_path)

        if save_path.endswith(".csv"):
            self.table_data.to_csv(save_path, index=False)
        elif save_path.endswith(".xlsx"):
            self.table_data.to_excel(save_path, index=False)

        QMessageBox.information(self, "Данные сохранены", "Данные успешно сохранены.")
    except Exception as e:
        logging.error(f"Ошибка при сохранении данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении данных: {str(e)}")

# Добавление функциональности загрузки данных
def load_data(self):
    try:
        file_name, _ = QFileDialog.getOpenFileName(self, "Выберите файл", "", "CSV (*.csv);;Excel (*.xlsx)")
        if file_name:
            if file_name.endswith(".csv"):
                self.table_data = pd.read_csv(file_name)
            elif file_name.endswith(".xlsx"):
                self.table_data = pd.read_excel(file_name)
            else:
                raise ValueError("Неподдерживаемый формат файла")

            model = CustomTableModel(self.table_data)
            self.table_view.setModel(model)

            logging.info("Данные успешно загружены")
            QMessageBox.information(self, "Успех", "Данные успешно загружены")
    except Exception as e:
        logging.error(f"Ошибка при загрузке данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке данных: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Обновление главного окна
def update_main_window(self):
    # Здесь добавьте логику обновления главного окна на основе ваших требований
    pass

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Запуск расписания
    while True:
        schedule.run_pending()
        app.processEvents()

    sys.exit(app.exec())
# Фильтрация данных
def filter_data(self):
    try:
        validate_dataframe(self.table_data)

        column, ok = QInputDialog.getText(self, "Выберите столбец", "Введите имя столбца:")
        if ok:
            filter_value, ok = QInputDialog.getText(self, "Введите значение", "Введите значение для фильтрации:")
            if ok:
                filtered_data = self.table_data[self.table_data[column] == filter_value]
                model = CustomTableModel(filtered_data)
                self.table_view.setModel(model)

                logging.info("Данные успешно отфильтрованы")
                QMessageBox.information(self, "Успех", "Данные успешно отфильтрованы")
    except Exception as e:
        logging.error(f"Ошибка при фильтрации данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при фильтрации данных: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Добавление функциональности поиска данных
def search_data(self):
    try:
        validate_dataframe(self.table_data)

        search_text, ok = QInputDialog.getText(self, "Введите текст", "Введите текст для поиска:")
        if ok:
            search_results = self.table_data.apply(lambda row: row.astype(str).str.contains(search_text, case=False).any(), axis=1)
            filtered_data = self.table_data[search_results]
            model = CustomTableModel(filtered_data)
            self.table_view.setModel(model)

            logging.info("Данные успешно найдены")
            QMessageBox.information(self, "Успех", "Данные успешно найдены")
    except Exception as e:
        logging.error(f"Ошибка при поиске данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при поиске данных: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Запуск расписания
    while True:
        schedule.run_pending()
        app.processEvents()

    sys.exit(app.exec())
# Сортировка данных
def sort_data(self):
    try:
        validate_dataframe(self.table_data)

        column, ok = QInputDialog.getText(self, "Выберите столбец", "Введите имя столбца:")
        if ok:
            sorted_data = self.table_data.sort_values(by=column)
            model = CustomTableModel(sorted_data)
            self.table_view.setModel(model)

            logging.info("Данные успешно отсортированы")
            QMessageBox.information(self, "Успех", "Данные успешно отсортированы")
    except Exception as e:
        logging.error(f"Ошибка при сортировке данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при сортировке данных: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Запуск расписания
    while True:
        schedule.run_pending()
        app.processEvents()

    sys.exit(app.exec())
    # Создание графиков
def create_plot(self):
    try:
        validate_dataframe(self.table_data)

        x_column, ok = QInputDialog.getText(self, "Выберите столбец X", "Введите имя столбца X:")
        if ok:
            y_column, ok = QInputDialog.getText(self, "Выберите столбец Y", "Введите имя столбца Y:")
            if ok:
                plt.plot(self.table_data[x_column], self.table_data[y_column])
                plt.xlabel(x_column)
                plt.ylabel(y_column)
                plt.title("График")
                plt.show()

                logging.info("График успешно создан")
                QMessageBox.information(self, "Успех", "График успешно создан")
    except Exception as e:
        logging.error(f"Ошибка при создании графика: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при создании графика: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Запуск расписания
    while True:
        schedule.run_pending()
        app.processEvents()

    sys.exit(app.exec())
    # Экспорт данных в PDF
def export_to_pdf(self):
    try:
        validate_dataframe(self.table_data)

        save_path, _ = QFileDialog.getSaveFileName(self, "Экспорт в PDF", "", "PDF (*.pdf)")
        validate_file_path(save_path)

        html = self.table_data.to_html(index=False)
        pdfkit.from_string(html, save_path)

        logging.info("Данные успешно экспортированы в PDF")
        QMessageBox.information(self, "Успех", "Данные успешно экспортированы в PDF")
    except Exception as e:
        logging.error(f"Ошибка при экспорте данных в PDF: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте данных в PDF: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Запуск расписания
    while True:
        schedule.run_pending()
        app.processEvents()

    sys.exit(app.exec())
    # Отправка данных по электронной почте
def send_email(self):
    try:
        validate_dataframe(self.table_data)

        email_dialog = QInputDialog()
        email_dialog.setWindowTitle("Отправка по электронной почте")
        email_dialog.setLabelText("Введите адрес электронной почты:")
        email_dialog.setInputMode(QInputDialog.TextInput)
        email_dialog.setTextEchoMode(QLineEdit.Normal)
        if email_dialog.exec_() == QInputDialog.Accepted:
            recipient_email = email_dialog.textValue()

            # Создание письма
            message = MIMEMultipart()
            message["From"] = self.sender_email
            message["To"] = recipient_email
            message["Subject"] = "Данные"

            # Текстовая версия письма
            text = self.table_data.to_string(index=False)
            text_part = MIMEText(text, "plain")
            message.attach(text_part)

            # Отправка письма
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.send_message(message)

            logging.info("Данные успешно отправлены по электронной почте")
            QMessageBox.information(self, "Успех", "Данные успешно отправлены по электронной почте")
    except Exception as e:
        logging.error(f"Ошибка при отправке данных по электронной почте: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при отправке данных по электронной почте: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Запуск расписания
    while True:
        schedule.run_pending()
        app.processEvents()

    sys.exit(app.exec())
    # Шифрование данных
def encrypt_data(self):
    try:
        validate_dataframe(self.table_data)

        # Генерация ключа шифрования
        key = Fernet.generate_key()

        # Создание объекта Fernet
        cipher = Fernet(key)

        # Преобразование данных в строку
        data_str = self.table_data.to_csv(index=False)

        # Шифрование данных
        encrypted_data = cipher.encrypt(data_str.encode())

        # Сохранение зашифрованных данных в файл
        save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить зашифрованные данные", "", "Encrypted File (*.enc)")
        validate_file_path(save_path)

        with open(save_path, "wb") as file:
            file.write(encrypted_data)

        logging.info("Данные успешно зашифрованы")
        QMessageBox.information(self, "Успех", "Данные успешно зашифрованы")
    except Exception as e:
        logging.error(f"Ошибка при шифровании данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при шифровании данных: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Запуск расписания
    while True:
        schedule.run_pending()
        app.processEvents()

    sys.exit(app.exec())
# Дешифрование данных
def decrypt_data(self):
    try:
        # Выбор зашифрованного файла
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите файл с зашифрованными данными", "", "Encrypted File (*.enc)")
        validate_file_path(file_path)

        # Чтение зашифрованных данных из файла
        with open(file_path, "rb") as file:
            encrypted_data = file.read()

        # Ввод ключа шифрования
        key, ok = QInputDialog.getText(self, "Введите ключ шифрования", "Введите ключ:")
        if ok:
            cipher = Fernet(key.encode())

            # Дешифрование данных
            decrypted_data = cipher.decrypt(encrypted_data)

            # Преобразование данных в DataFrame
            decrypted_df = pd.read_csv(io.BytesIO(decrypted_data))

            # Обновление модели представления данных
            model = CustomTableModel(decrypted_df)
            self.table_view.setModel(model)

            logging.info("Данные успешно дешифрованы")
            QMessageBox.information(self, "Успех", "Данные успешно дешифрованы")
    except Exception as e:
        logging.error(f"Ошибка при дешифровании данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при дешифровании данных: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Запуск расписания
    while True:
        schedule.run_pending()
        app.processEvents()

    sys.exit(app.exec())
    # Команда отмены и повтора действий
class ActionCommand(QUndoCommand):
    def __init__(self, action, stack):
        super().__init__()
        self.action = action
        self.stack = stack

    def undo(self):
        self.action.undo()

    def redo(self):
        self.action.redo()

# ... (весь предыдущий код остается без изменений) ...

# Инициализация стека отмены и повтора
self.undo_stack = QUndoStack(self)

# ... (весь предыдущий код остается без изменений) ...

# Функция отмены действия
def undo_action(self):
    self.undo_stack.undo()

# Функция повтора действия
def redo_action(self):
    self.undo_stack.redo()

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
# Сохранение состояния приложения
def save_state(self):
    try:
        state = {
            'table_data': self.table_data,
            # Добавьте здесь все остальные переменные состояния, которые вы хотите сохранить
        }

        save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить состояние", "", "State File (*.state)")
        validate_file_path(save_path)

        with open(save_path, "wb") as file:
            pickle.dump(state, file)

        logging.info("Состояние приложения успешно сохранено")
        QMessageBox.information(self, "Успех", "Состояние приложения успешно сохранено")
    except Exception as e:
        logging.error(f"Ошибка при сохранении состояния приложения: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении состояния приложения: {str(e)}")

# Загрузка состояния приложения
def load_state(self):
    try:
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите файл состояния", "", "State File (*.state)")
        validate_file_path(file_path)

        with open(file_path, "rb") as file:
            state = pickle.load(file)

        self.table_data = state['table_data']
        # Восстановление других переменных состояния

        model = CustomTableModel(self.table_data)
        self.table_view.setModel(model)

        logging.info("Состояние приложения успешно загружено")
        QMessageBox.information(self, "Успех", "Состояние приложения успешно загружено")
    except Exception as e:
        logging.error(f"Ошибка при загрузке состояния приложения: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке состояния приложения: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
    # Вставка данных из буфера обмена
def paste_data(self):
    try:
        clipboard = QApplication.clipboard()
        data = clipboard.mimeData()

        if data.hasText():
            # Получение текстовых данных из буфера обмена
            text = data.text()
            lines = text.split('\n')

            # Преобразование текстовых данных в DataFrame
            df = pd.read_csv(io.StringIO(text), sep='\t')

            # Добавление данных в таблицу
            self.table_data = pd.concat([self.table_data, df], ignore_index=True)

            # Обновление модели представления данных
            model = CustomTableModel(self.table_data)
            self.table_view.setModel(model)

            logging.info("Данные успешно вставлены из буфера обмена")
            QMessageBox.information(self, "Успех", "Данные успешно вставлены из буфера обмена")
    except Exception as e:
        logging.error(f"Ошибка при вставке данных из буфера обмена: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при вставке данных из буфера обмена: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
    # Создание резервной копии данных
def create_backup(self):
    try:
        backup_dir = "backup"
        os.makedirs(backup_dir, exist_ok=True)

        backup_path = os.path.join(backup_dir, "backup.csv")

        self.table_data.to_csv(backup_path, index=False)

        logging.info("Резервная копия данных успешно создана")
        QMessageBox.information(self, "Успех", "Резервная копия данных успешно создана")
    except Exception as e:
        logging.error(f"Ошибка при создании резервной копии данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при создании резервной копии данных: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
    # Фильтрация данных
def filter_data(self):
    try:
        filter_dialog = QInputDialog()
        filter_dialog.setWindowTitle("Фильтрация данных")
        filter_dialog.setLabelText("Введите условие фильтрации (например, 'column_name > value'):")
        filter_dialog.setInputMode(QInputDialog.TextInput)
        filter_dialog.setTextEchoMode(QLineEdit.Normal)
        if filter_dialog.exec_() == QInputDialog.Accepted:
            filter_condition = filter_dialog.textValue()

            filtered_data = self.table_data.query(filter_condition)

            model = CustomTableModel(filtered_data)
            self.table_view.setModel(model)

            logging.info("Данные успешно отфильтрованы")
            QMessageBox.information(self, "Успех", "Данные успешно отфильтрованы")
    except Exception as e:
        logging.error(f"Ошибка при фильтрации данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при фильтрации данных: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
    # Сортировка данных
def sort_data(self):
    try:
        sort_dialog = QInputDialog()
        sort_dialog.setWindowTitle("Сортировка данных")
        sort_dialog.setLabelText("Введите столбец для сортировки:")
        sort_dialog.setInputMode(QInputDialog.TextInput)
        sort_dialog.setTextEchoMode(QLineEdit.Normal)
        if sort_dialog.exec_() == QInputDialog.Accepted:
            sort_column = sort_dialog.textValue()

            sorted_data = self.table_data.sort_values(by=sort_column)

            model = CustomTableModel(sorted_data)
            self.table_view.setModel(model)

            logging.info("Данные успешно отсортированы")
            QMessageBox.information(self, "Успех", "Данные успешно отсортированы")
    except Exception as e:
        logging.error(f"Ошибка при сортировке данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при сортировке данных: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
# Расчет статистических показателей
def calculate_statistics(self):
    try:
        selected_columns = self.table_view.selectedIndexes()
        if not selected_columns:
            QMessageBox.warning(self, "Предупреждение", "Не выбраны столбцы для расчета статистики")
            return

        columns = set()
        for index in selected_columns:
            column = self.table_data.columns[index.column()]
            columns.add(column)

        statistics = self.table_data[list(columns)].describe()

        statistics_dialog = QDialog()
        statistics_dialog.setWindowTitle("Статистика")
        layout = QVBoxLayout()
        table_view = QTableView()
        table_view.setModel(CustomTableModel(statistics))
        layout.addWidget(table_view)
        statistics_dialog.setLayout(layout)
        statistics_dialog.exec_()
    except Exception as e:
        logging.error(f"Ошибка при расчете статистики: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при расчете статистики: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
# Графическое представление данных
def visualize_data(self):
    try:
        selected_columns = self.table_view.selectedIndexes()
        if not selected_columns:
            QMessageBox.warning(self, "Предупреждение", "Не выбраны столбцы для визуализации")
            return

        columns = set()
        for index in selected_columns:
            column = self.table_data.columns[index.column()]
            columns.add(column)

        data_to_visualize = self.table_data[list(columns)]

        fig, ax = plt.subplots(figsize=(10, 6))
        data_to_visualize.plot(ax=ax)
        ax.legend()
        plt.show()
    except Exception as e:
        logging.error(f"Ошибка при визуализации данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при визуализации данных: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())     
# Экспорт данных в файл
def export_data(self):
    try:
        export_dialog = QFileDialog()
        export_dialog.setWindowTitle("Экспорт данных")
        export_dialog.setAcceptMode(QFileDialog.AcceptSave)
        export_dialog.setNameFilter("CSV Files (*.csv);;Excel Files (*.xlsx)")

        if export_dialog.exec_() == QFileDialog.Accepted:
            file_path = export_dialog.selectedFiles()[0]

            if file_path.endswith(".csv"):
                self.table_data.to_csv(file_path, index=False)
                logging.info("Данные успешно экспортированы в файл CSV")
            elif file_path.endswith(".xlsx"):
                self.table_data.to_excel(file_path, index=False)
                logging.info("Данные успешно экспортированы в файл Excel")
            else:
                raise ValueError("Неподдерживаемый формат файла")

            QMessageBox.information(self, "Успех", "Данные успешно экспортированы в файл")
    except Exception as e:
        logging.error(f"Ошибка при экспорте данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте данных: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
# Настройка внешнего вида таблицы
def customize_table(self):
    try:
        font, font_accepted = QFontDialog.getFont(self.font())
        if font_accepted:
            self.table_view.setFont(font)

        background_color = QColorDialog.getColor(self.table_view.backgroundRole().color())
        if background_color.isValid():
            palette = self.table_view.palette()
            palette.setColor(QPalette.Base, background_color)
            self.table_view.setPalette(palette)

        text_color = QColorDialog.getColor(self.table_view.foregroundRole().color())
        if text_color.isValid():
            palette = self.table_view.palette()
            palette.setColor(QPalette.Text, text_color)
            self.table_view.setPalette(palette)

        logging.info("Настройки внешнего вида таблицы успешно применены")
        QMessageBox.information(self, "Успех", "Настройки внешнего вида таблицы успешно применены")
    except Exception as e:
        logging.error(f"Ошибка при настройке внешнего вида таблицы: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при настройке внешнего вида таблицы: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
# Автоматическое выполнение задач
def automate_tasks():
    try:
        # Здесь добавьте вызов необходимых функций для автоматического выполнения задач
        # Например:
        # load_data()
        # compare_tables()
        # create_report()

        logging.info("Задачи успешно выполнены в соответствии с расписанием")
    except Exception as e:
        logging.error(f"Ошибка при автоматическом выполнении задач: {str(e)}")

# Запуск автоматического выполнения задач по расписанию
def start_automation():
    schedule.every().day.at("12:00").do(automate_tasks)
    # Добавьте дополнительные расписания, если необходимо
    # schedule.every().monday.at("9:00").do(task1)
    # schedule.every().wednesday.at("14:00").do(task2)

    while True:
        schedule.run_pending()
        time.sleep(1)

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Запуск автоматического выполнения задач
    automation_thread = threading.Thread(target=start_automation)
    automation_thread.start()

    sys.exit(app.exec())
# Обработка ошибок и валидация данных
def validate_data(data):
    try:
        # Проверка типа данных (DataFrame)
        if not isinstance(data, pd.DataFrame):
            raise ValueError("Неверный тип данных. Ожидается DataFrame.")

        # Проверка наличия необходимых столбцов
        required_columns = ["column1", "column2", "column3"]
        missing_columns = set(required_columns) - set(data.columns)
        if missing_columns:
            raise ValueError(f"Отсутствуют необходимые столбцы: {', '.join(missing_columns)}")

        # Другие проверки и валидации данных
        # ...

        logging.info("Проверка и валидация данных успешно выполнены")
        return True
    except Exception as e:
        logging.error(f"Ошибка при проверке и валидации данных: {str(e)}")
        return False

# ... (весь предыдущий код остается без изменений) ...

# Загрузка данных
def load_data():
    try:
        # Код для загрузки данных
        # ...

        # Проверка и валидация данных
        if validate_data(data):
            # Добавление данных в таблицу
            self.table_data = pd.concat([self.table_data, data], ignore_index=True)

            # Обновление модели представления данных
            model = CustomTableModel(self.table_data)
            self.table_view.setModel(model)

            logging.info("Данные успешно загружены")
            QMessageBox.information(self, "Успех", "Данные успешно загружены")
    except Exception as e:
        logging.error(f"Ошибка при загрузке данных: {str(e)}")
        QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке данных: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
# Ведение лога в консоли
def setup_logging():
    logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    setup_logging()

    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
# Обработка ошибок и вывод сообщений об ошибке
def show_error_message(message, recommendation=None):
    error_message = QMessageBox()
    error_message.setIcon(QMessageBox.Critical)
    error_message.setWindowTitle("Ошибка")
    error_message.setText(message)
    if recommendation:
        error_message.setInformativeText(f"Рекомендация: {recommendation}")
    error_message.exec_()

# ... (весь предыдущий код остается без изменений) ...

# Обработка ошибок и валидация данных
def validate_data(data):
    try:
        # Проверка типа данных (DataFrame)
        if not isinstance(data, pd.DataFrame):
            raise ValueError("Неверный тип данных. Ожидается DataFrame.")

        # Проверка наличия необходимых столбцов
        required_columns = ["column1", "column2", "column3"]
        missing_columns = set(required_columns) - set(data.columns)
        if missing_columns:
            missing_columns_str = ", ".join(missing_columns)
            error_message = f"Отсутствуют необходимые столбцы: {missing_columns_str}"
            recommendation = "Проверьте структуру данных и убедитесь, что все необходимые столбцы присутствуют."
            show_error_message(error_message, recommendation)
            return False

        # Другие проверки и валидации данных
        # ...

        logging.info("Проверка и валидация данных успешно выполнены")
        return True
    except Exception as e:
        logging.error(f"Ошибка при проверке и валидации данных: {str(e)}")
        error_message = f"Ошибка при проверке и валидации данных: {str(e)}"
        show_error_message(error_message)
        return False

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
# Создание отчета и сохранение в файл
def create_report():
    try:
        # Выбор столбцов для анализа
        selected_columns = self.table_view.selectedIndexes()
        if not selected_columns:
            show_error_message("Не выбраны столбцы для анализа.")
            return

        # Получение данных для анализа
        data_to_analyze = self.table_data.iloc[:, [index.column() for index in selected_columns]]

        # Выполнение статистического анализа
        report_data = data_to_analyze.describe()

        # Сохранение отчета в файл
        save_dialog = QFileDialog()
        save_dialog.setWindowTitle("Сохранить отчет")
        save_dialog.setAcceptMode(QFileDialog.AcceptSave)
        save_dialog.setNameFilter("CSV Files (*.csv);;Excel Files (*.xlsx)")

        if save_dialog.exec_() == QFileDialog.Accepted:
            file_path = save_dialog.selectedFiles()[0]

            if file_path.endswith(".csv"):
                report_data.to_csv(file_path)
                logging.info("Отчет успешно сохранен в файл CSV")
            elif file_path.endswith(".xlsx"):
                report_data.to_excel(file_path)
                logging.info("Отчет успешно сохранен в файл Excel")
            else:
                raise ValueError("Неподдерживаемый формат файла")

            QMessageBox.information(self, "Успех", "Отчет успешно сохранен в файл")
    except Exception as e:
        logging.error(f"Ошибка при создании отчета и сохранении в файл: {str(e)}")
        show_error_message(f"Ошибка при создании отчета и сохранении в файл: {str(e)}")

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
# Создание пользовательского интерфейса для выбора параметров автоматического выполнения задач
def create_automation_ui():
    automation_dialog = QDialog()
    automation_dialog.setWindowTitle("Настройки автоматического выполнения задач")

    # ... (код для создания пользовательского интерфейса выбора параметров) ...

    return automation_dialog

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Создание пользовательского интерфейса для выбора параметров автоматического выполнения задач
    automation_ui = create_automation_ui()

    sys.exit(app.exec())
# Модульное тестирование
class TestApplication(unittest.TestCase):
    def setUp(self):
        # Подготовка данных или настройка окружения для каждого тестового случая
        pass

    def tearDown(self):
        # Зачистка после каждого тестового случая
        pass

    def test_load_data(self):
        # Тестирование загрузки данных
        # ...

    def test_compare_tables(self):
        # Тестирование сравнения таблиц
        # ...

    def test_create_report(self):
        # Тестирование создания отчета
        # ...

    # Добавьте дополнительные методы для других тестовых случаев

# Запуск модульного тестирования
if __name__ == "__main__":
    unittest.main()

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())              
    # Общее тестирование приложения
def test_application():
    # Тестирование загрузки данных
    test_load_data()

    # Тестирование сравнения таблиц
    test_compare_tables()

    # Тестирование создания отчета
    test_create_report()

    # ... (добавьте вызовы других тестовых методов, если необходимо)

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Общее тестирование приложения
    test_application()

    sys.exit(app.exec())
   # Автоматическое выполнение задач по расписанию
def run_scheduled_tasks():
    # Задайте расписание выполнения задач
    schedule.every().day.at("09:00").do(task1)
    schedule.every().monday.at("14:30").do(task2)
    schedule.every(5).minutes.do(task3)

    while True:
        schedule.run_pending()
        time.sleep(1)

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Создание пользовательского интерфейса для выбора параметров автоматического выполнения задач
    automation_ui = create_automation_ui()

    # Автоматическое выполнение задач по расписанию
    run_scheduled_tasks()

    sys.exit(app.exec()) 
    # Проверка входных параметров и валидация данных
def validate_input_parameters(data, columns):
    try:
        # Проверка типа данных (DataFrame)
        if not isinstance(data, pd.DataFrame):
            raise ValueError("Неверный тип данных. Ожидается DataFrame.")

        # Проверка наличия необходимых столбцов
        missing_columns = set(columns) - set(data.columns)
        if missing_columns:
            missing_columns_str = ", ".join(missing_columns)
            raise ValueError(f"Отсутствуют необходимые столбцы: {missing_columns_str}")

        # Другие проверки и валидации данных
        # ...

        logging.info("Проверка и валидация данных успешно выполнены")
        return True
    except Exception as e:
        logging.error(f"Ошибка при проверке и валидации данных: {str(e)}")
        show_error_message(f"Ошибка при проверке и валидации данных: {str(e)}")
        return False

# ... (весь предыдущий код остается без изменений) ...

# Главная функция
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    # Проверка входных параметров и валидация данных
    if validate_input_parameters(data, columns):
        sys.exit(app.exec())
    sys.exit(app.exec())
