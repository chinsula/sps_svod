import sys
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QFileDialog,
    QVBoxLayout, QLabel, QMessageBox
)
import pandas as pd

class ExcelComparer(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Сравнение и вывод совпадающих строк по всему столбцу")
        self.resize(600, 200)

        self.layout = QVBoxLayout()

        self.btn_select_file1 = QPushButton("Выбрать первый файл Excel")
        self.label_file1_path = QLabel("Первый файл не выбран")
        self.btn_select_file2 = QPushButton("Выбрать второй файл Excel")
        self.label_file2_path = QLabel("Второй файл не выбран")
        self.btn_save_location = QPushButton("Выбрать место сохранения")
        self.label_save_path = QLabel("Место сохранения не выбрано")
        self.btn_run = QPushButton("Сравнить и сохранить")

        self.layout.addWidget(self.btn_select_file1)
        self.layout.addWidget(self.label_file1_path)
        self.layout.addWidget(self.btn_select_file2)
        self.layout.addWidget(self.label_file2_path)
        self.layout.addWidget(self.btn_save_location)
        self.layout.addWidget(self.label_save_path)
        self.layout.addWidget(self.btn_run)

        self.setLayout(self.layout)

        self.file_path1 = None
        self.file_path2 = None
        self.save_path = None

        self.btn_select_file1.clicked.connect(self.select_file1)
        self.btn_select_file2.clicked.connect(self.select_file2)
        self.btn_save_location.clicked.connect(self.select_save_location)
        self.btn_run.clicked.connect(self.compare_and_save)

    def select_file1(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите первый файл Excel", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.file_path1 = path
            self.label_file1_path.setText(f"Первый файл: {path}")

    def select_file2(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите второй файл Excel", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.file_path2 = path
            self.label_file2_path.setText(f"Второй файл: {path}")

    def select_save_location(self):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"совпадения_{timestamp}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Выберите место сохранения", default_name, "Excel Files (*.xlsx)")
        if path:
            self.save_path = path
            self.label_save_path.setText(f"Место сохранения: {path}")

    def compare_and_save(self):
        if not all([self.file_path1, self.file_path2, self.save_path]):
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, выберите оба файла и место сохранения.")
            return
        try:
            df1 = pd.read_excel(self.file_path1)
            df2 = pd.read_excel(self.file_path2)

            # Проверка наличия хотя бы двух столбцов
            if len(df1.columns) < 2 or len(df2.columns) < 2:
                QMessageBox.warning(self, "Ошибка", "Один из файлов не содержит хотя бы двух столбцов.")
                return

            col1 = df1.iloc[:, 0]
            col2 = df2.iloc[:, 0]

            # Стандартизация (убираем пробелы, делаем нижний регистр)
            def standardize(series):
                return series.astype(str).str.replace(r'\s+', '', regex=True).str.lower()

            s1 = standardize(col1)
            s2 = standardize(col2)

            # Поиск совпадений по всему столбцу
            # Создаем множества уникальных значений для быстрого поиска
            set_s1 = set(s1)
            set_s2 = set(s2)

            # Общие совпадения
            common_values = set_s1.intersection(set_s2)

            # Индексы строк, где значения совпадают
            indices_s1 = s1[s1.isin(common_values)].index
            indices_s2 = s2[s2.isin(common_values)].index

            # Создаем DataFrame с соответствующими вторыми столбцами
            result_df = pd.DataFrame({
                'Второй столбец файла 1': col1.iloc[indices_s1].reset_index(drop=True),
                'Второй столбец файла 2': col2.iloc[indices_s2].reset_index(drop=True)
            })

            # Сохраняем результат
            result_df.to_excel(self.save_path, index=False)

            QMessageBox.information(self, "Успех", "Файл успешно сохранен.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelComparer()
    window.show()
    sys.exit(app.exec())