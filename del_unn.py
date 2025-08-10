import sys
import pandas as pd
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QMessageBox
)

class FileProcessor(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Обработка файлов без транслитерации")
        self.resize(400, 200)
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
        self.btn_start = QPushButton("Запустить обработку")
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
        path, _ = QFileDialog.getOpenFileName(self, "Выбрать первый файл")
        if path:
            self.first_file_path = path
            self.lbl_first.setText(f"Первый файл: {path}")
            self.check_ready()

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
        if self.first_file_path and self.second_file_path and self.save_dir:
            self.btn_start.setEnabled(True)

    def process_files(self):
        try:
            # Загрузка файлов
            df_first = pd.read_excel(self.first_file_path, header=None)
            df_second = pd.read_excel(self.second_file_path, header=None)

            # Выводим все значения из второго файла, начиная с третьей строки
            print("Все исходные значения из второго файла:")
            for idx, text in enumerate(df_second.iloc[2:, 1]):
                print(f"Строка {idx + 2} (начиная с 0): {text}")

            # Собираем множество из второго файла, начиная с 3-й строки (индекс 2)
            set_of_strings = set()
            for text in df_second.iloc[2:, 1]:
                if pd.isna(text):
                    continue
                cleaned = ''.join(str(text).split()).upper()
                set_of_strings.add(cleaned)

            # Выводим сформированное множество
            print("Множество из второго файла (после сбора):")
            for item in set_of_strings:
                print(f"'{item}'")
            print("\n")  # разделитель

            # Проверяем строки из первого файла
            результаты = []

            for idx, row in df_first.iterrows():
                for col_idx in [0, 2]:
                    cell = str(row[col_idx])
                    cell_cleaned = ''.join(cell).replace(' ', '').upper()
                    if any(elem in cell_cleaned for elem in set_of_strings):
                        результаты.append(row)
                        break

            # Удаляем дубликаты
            результаты = [list(x) for x in {tuple(row) for row in результаты}]

            # Сохраняем результат
            if результаты:
                df_result = pd.DataFrame(результаты)
                filename = f"Результат_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                save_path = f"{self.save_dir}/{filename}"
                df_result.to_excel(save_path, index=False, header=False)

                # Вывод в консоль
                print("Множество из второго файла (после сбора):")
                for item in set_of_strings:
                    print(f"'{item}'")
                print("\nСодержимое второго файла (индексы строк и столбцы):")
                for row_idx, row in df_second.iterrows():
                    row_str = " | ".join([f"{col_idx}:{row[col_idx]}" for col_idx in range(len(row))])
                    print(f"Строка {row_idx}: {row_str}")

                QMessageBox.information(self, "Готово", f"Результат сохранен: {save_path}")
            else:
                print("Нет совпадений.")
                QMessageBox.information(self, "Результат", "Совпадений не найдено.")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))


if __name__ == "__main__":
    app = QApplication([])
    window = FileProcessor()
    window.show()
    sys.exit(app.exec())