import sys
import pandas as pd
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QMessageBox, QComboBox
)

class FileProcessor(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Обработка файлов без транслитерации")
        self.resize(400, 250)
        self.setup_ui()

    def setup_ui(self):
        self.first_file_path = None
        self.second_file_path = None
        self.save_dir = None
        self.df_first = None

        layout = QVBoxLayout()

        # Выбор файла для обработки
        self.btn_first = QPushButton("Выбрать файл для обработки")
        self.lbl_first = QLabel("Первый файл не выбран")
        layout.addWidget(self.btn_first)
        layout.addWidget(self.lbl_first)

        # Выбор файла-базы данных
        self.btn_second = QPushButton("Выбрать файл-базу данных")
        self.lbl_second = QLabel("Второй файл не выбран")
        layout.addWidget(self.btn_second)
        layout.addWidget(self.lbl_second)

        # Выбор папки для сохранения
        self.btn_save = QPushButton("Выбрать папку для сохранения")
        self.lbl_save = QLabel("Папка не выбрана")
        layout.addWidget(self.btn_save)
        layout.addWidget(self.lbl_save)

        # Выбор номера столбца (через выпадающий список)
        self.label_column = QLabel("Выберите номер столбца (начиная с 1):")
        self.combo_columns = QComboBox()
        self.combo_columns.addItem("Нет данных")  # изначально, обновится после загрузки файла
        layout.addWidget(self.label_column)
        layout.addWidget(self.combo_columns)

        # Кнопка запуска
        self.btn_start = QPushButton("Запустить обработку")
        self.btn_start.setEnabled(False)
        layout.addWidget(self.btn_start)

        self.setLayout(layout)

        # Связи
        self.btn_first.clicked.connect(self.select_first_file)
        self.btn_second.clicked.connect(self.select_second_file)
        self.btn_save.clicked.connect(self.select_save_directory)
        self.btn_start.clicked.connect(self.process_files)

    def select_first_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выбрать первый файл")
        if path:
            self.first_file_path = path
            self.lbl_first.setText(f"Первый файл: {path}")
            self.load_first_file()
            self.check_ready()

    def load_first_file(self):
        try:
            self.df_first = pd.read_excel(self.first_file_path, header=None, engine='openpyxl')
            num_columns = len(self.df_first.columns)
            self.combo_columns.clear()
            if num_columns > 0:
                for i in range(1, num_columns + 1):
                    self.combo_columns.addItem(str(i))
            else:
                self.combo_columns.addItem("Нет данных")
            self.combo_columns.setCurrentIndex(0)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка чтения файла", str(e))
            self.df_first = None

    def select_second_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выбрать второй файл")
        if path:
            self.second_file_path = path
            self.lbl_second.setText(f"Второй файл: {path}")
            self.check_ready()

    def select_save_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Выбрать папку для сохранения")
        if directory:
            self.save_dir = directory
            self.lbl_save.setText(f"Папка: {directory}")
            self.check_ready()

    def check_ready(self):
        if self.first_file_path and self.second_file_path and self.save_dir and self.df_first is not None:
            self.btn_start.setEnabled(True)

    def process_files(self):
        try:
            df_second = pd.read_excel(self.second_file_path, header=None, engine='openpyxl')

            # Создаем множество из второго файла
            set_of_strings = set()
            for text in df_second.iloc[:, 0]:
                if pd.isna(text):
                    continue
                s = str(text)
                if len(s) > 2:
                    s = s[:-2]
                s_clean = ''.join(s).replace(' ', '').upper()
                set_of_strings.add(s_clean)

            # Получаем выбранный номер столбца из ComboBox (строка)
            selected_index_str = self.combo_columns.currentText()
            if selected_index_str == "Нет данных" or not selected_index_str.isdigit():
                QMessageBox.warning(self, "Ошибка", "Выберите допустимый номер столбца.")
                return
            column_index = int(selected_index_str) - 1

            результаты = []

            for idx, row in self.df_first.iterrows():
                if column_index >= len(row):
                    continue
                cell = str(row[column_index])
                cell_cleaned = ''.join(cell).replace(' ', '').upper()
                if any(elem in cell_cleaned for elem in set_of_strings):
                    результаты.append(row)

            # Удаляем дубликаты
            результаты = [list(x) for x in {tuple(row) for row in результаты}]

            if результаты:
                df_result = pd.DataFrame(результаты)
                filename = f"Результат_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                save_path = f"{self.save_dir}/{filename}"
                df_result.to_excel(save_path, index=False, header=False)
                QMessageBox.information(self, "Готово", f"Результат сохранен: {save_path}")
            else:
                QMessageBox.information(self, "Результат", "Совпадений не найдено.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

if __name__ == "__main__":
    app = QApplication([])
    window = FileProcessor()
    window.show()
    sys.exit(app.exec())