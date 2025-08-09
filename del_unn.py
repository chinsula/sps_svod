import sys
import pandas as pd
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QMessageBox
)

class FilterRowsWithSetElements(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Фильтр: строки с элементами множества")
        self.setup_ui()

    def setup_ui(self):
        self.first_file_path = None
        self.second_file_path = None
        self.save_dir = None

        layout = QVBoxLayout()

        self.btn_first = QPushButton("Выбрать первый файл")
        self.lbl_first = QLabel("Первый файл не выбран")
        self.btn_second = QPushButton("Выбрать второй файл")
        self.lbl_second = QLabel("Второй файл не выбран")
        self.btn_save = QPushButton("Выбрать папку для сохранения")
        self.lbl_save = QLabel("Папка не выбрана")
        self.btn_start = QPushButton("Запустить")
        self.btn_start.setEnabled(False)

        layout.addWidget(self.btn_first)
        layout.addWidget(self.lbl_first)
        layout.addWidget(self.btn_second)
        layout.addWidget(self.lbl_second)
        layout.addWidget(self.btn_save)
        layout.addWidget(self.lbl_save)
        layout.addWidget(self.btn_start)

        self.setLayout(layout)

        self.btn_first.clicked.connect(self.select_first_file)
        self.btn_second.clicked.connect(self.select_second_file)
        self.btn_save.clicked.connect(self.select_save_directory)
        self.btn_start.clicked.connect(self.process_files)

    def select_first_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите первый файл")
        if path:
            self.first_file_path = path
            self.lbl_first.setText(f"Первый файл: {path}")
            self.check_ready()

    def select_second_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите второй файл")
        if path:
            self.second_file_path = path
            self.lbl_second.setText(f"Второй файл: {path}")
            self.check_ready()

    def select_save_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения")
        if directory:
            self.save_dir = directory
            self.lbl_save.setText(f"Папка: {directory}")
            self.check_ready()

    def check_ready(self):
        if self.first_file_path and self.second_file_path and self.save_dir:
            self.btn_start.setEnabled(True)

    def process_files(self):
        try:
            # Загружаем файлы
            df_first = pd.read_excel(self.first_file_path, header=None)
            df_second = pd.read_excel(self.second_file_path, header=None)

            # Создаем множество строк из 3-го столбца второго файла
            set_of_strings = set()
            for text in df_second.iloc[:, 2]:  # 3-й столбец
                if pd.isna(text):
                    continue
                cleaned = ''.join(str(text).split()).upper()
                set_of_strings.add(cleaned)

            print(f"Множество из 3-й колонки второго файла: {set_of_strings}")

            # Проверяем каждую строку первого файла
            result_rows = []

            for index, row in df_first.iterrows():
                row_str = ''
                # Проверяем 1-й и 3-й столбцы
                for col_idx in [0, 2]:
                    cell_value = str(row[col_idx])
                    cleaned_cell = ''.join(cell_value).replace(' ', '').upper()
                    print(f"Строка {index} в колонке {col_idx}: '{cell_value}' -> '{cleaned_cell}'")
                    # Проверяем наличие любого элемента множества в этой строке
                    if any(element in cleaned_cell for element in set_of_strings):
                        print(f"В строке {index} найден элемент множества")
                        result_rows.append(row.tolist())
                        break  # как только нашли совпадение, переходим к следующей строке

            # Удаляем дубликаты
            result_rows = [list(x) for x in {tuple(row) for row in result_rows}]

            # Сохраняем результат
            if result_rows:
                df_result = pd.DataFrame(result_rows)
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
    window = FilterRowsWithSetElements()
    window.show()
    sys.exit(app.exec())