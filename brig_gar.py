import sys
import os
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QFileDialog,
    QVBoxLayout, QLabel, QLineEdit, QMessageBox
)
import pandas as pd

class ExcelComparer(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Сопоставление Excel файлов")
        self.resize(600, 400)

        self.layout = QVBoxLayout()

        # Кнопки и метки для выбора первого файла
        self.btn_select_file1 = QPushButton("Выбрать первый файл Excel")
        self.label_file1_path = QLabel("Первый файл не выбран")
        # Кнопки и метки для выбора второго файла
        self.btn_select_file2 = QPushButton("Выбрать второй файл Excel")
        self.label_file2_path = QLabel("Второй файл не выбран")
        # Кнопка и метка для выбора места сохранения
        self.btn_save_location = QPushButton("Выбрать место сохранения")
        self.label_save_path = QLabel("Место сохранения не выбрано")

        # Поля для ввода номеров столбцов
        self.label_col1 = QLabel("Номер столбца в первом файле (начинается с 0):")
        self.input_col1 = QLineEdit()
        self.label_col2 = QLabel("Номер столбца во втором файле (начинается с 0):")
        self.input_col2 = QLineEdit()

        # Кнопка для запуска сравнения
        self.btn_run = QPushButton("Запустить сравнение и сохранить результат")

        # Добавляем виджеты
        self.layout.addWidget(self.btn_select_file1)
        self.layout.addWidget(self.label_file1_path)
        self.layout.addWidget(self.btn_select_file2)
        self.layout.addWidget(self.label_file2_path)
        self.layout.addWidget(self.btn_save_location)
        self.layout.addWidget(self.label_save_path)
        self.layout.addWidget(self.label_col1)
        self.layout.addWidget(self.input_col1)
        self.layout.addWidget(self.label_col2)
        self.layout.addWidget(self.input_col2)
        self.layout.addWidget(self.btn_run)

        self.setLayout(self.layout)

        # Переменные для путей
        self.file_path1 = None
        self.file_path2 = None
        self.save_path = None

        # Связываем кнопки
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
        # Автоматически создаем имя файла с датой и временем
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"сравнение_{timestamp}.xlsx"
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

            # Получение номеров столбцов из ввода
            col_idx1 = int(self.input_col1.text())
            col_idx2 = int(self.input_col2.text())

            # Проверка наличия указанных столбцов
            if col_idx1 >= len(df1.columns) or col_idx2 >= len(df2.columns):
                QMessageBox.warning(self, "Ошибка", "Номер столбца превышает количество столбцов в файле.")
                return

            # Проверка, что выбранные столбцы не нулевые или пустые
            col_name1 = df1.columns[col_idx1]
            col_name2 = df2.columns[col_idx2]
            if str(col_name1).strip() in ('', '0') or str(col_name2).strip() in ('', '0'):
                QMessageBox.warning(self, "Ошибка", "Выбран нулевой или пустой столбец.")
                return

            # Извлекаем нужные столбцы
            col_series1 = df1.iloc[:, col_idx1]
            col_series2 = df2.iloc[:, col_idx2]

            # Приведение данных к единому виду
            def standardize(series):
                return series.astype(str).str.strip().str.lower()

            s1 = standardize(col_series1)
            s2 = standardize(col_series2)

            # Создаем результирующий DataFrame
            result_df = pd.DataFrame()

            # Вставляем первые столбцы каждого файла
            result_df['Первый файл'] = df1.iloc[:, 0]
            result_df['Второй файл'] = df2.iloc[:, 0]

            # Вставляем выбранные столбцы
            result_df['Выбранный столбец 1'] = col_series1
            result_df['Выбранный столбец 2'] = col_series2

            # Добавляем столбец с результатом сравнения
            result_df['Совпадает'] = s1.isin(s2)

            # Места для вставки дополнительных столбцов, например:
            # result_df['Дополнительный столбец'] = df1.iloc[:, номер]

            # Сохраняем с автоматическим именем
            result_df.to_excel(self.save_path, index=False)

            QMessageBox.information(self, "Успех", "Файл успешно сохранен.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelComparer()
    window.show()
    sys.exit(app.exec())