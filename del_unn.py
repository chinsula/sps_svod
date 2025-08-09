import sys
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QMessageBox
)
import pandas as pd
from datetime import datetime

class ExcelProcessor(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Обработка Excel файлов")
        self.layout = QVBoxLayout()

        self.btn_select_input_file = QPushButton("Выбрать исходный файл для обработки")
        self.lbl_input_file = QLabel("Файл для обработки не выбран")
        self.btn_select_data_file = QPushButton("Выбрать файл с данными для сравнения")
        self.lbl_data_file = QLabel("Файл с данными не выбран")
        self.btn_select_save_dir = QPushButton("Выбрать папку для сохранения результата")
        self.lbl_save_dir = QLabel("Папка для сохранения не выбрана")
        self.btn_start = QPushButton("Обработать и сохранить")
        self.btn_start.setEnabled(False)

        self.layout.addWidget(self.btn_select_input_file)
        self.layout.addWidget(self.lbl_input_file)
        self.layout.addWidget(self.btn_select_data_file)
        self.layout.addWidget(self.lbl_data_file)
        self.layout.addWidget(self.btn_select_save_dir)
        self.layout.addWidget(self.lbl_save_dir)
        self.layout.addWidget(self.btn_start)
        self.setLayout(self.layout)

        self.file_path_input = None
        self.file_path_data = None
        self.save_dir = None

        self.btn_select_input_file.clicked.connect(self.select_input_file)
        self.btn_select_data_file.clicked.connect(self.select_data_file)
        self.btn_select_save_dir.clicked.connect(self.select_save_directory)
        self.btn_start.clicked.connect(self.process_files)

        # Строка поиска
        self.search_str = "0747 УО"

    def select_input_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "Выберите файл для обработки", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.file_path_input = file
            self.lbl_input_file.setText(f"Файл для обработки: {file}")
            self.check_ready()

    def select_data_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "Выберите файл с данными для сравнения", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.file_path_data = file
            self.lbl_data_file.setText(f"Файл с данными: {file}")
            self.check_ready()

    def select_save_directory(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения результата")
        if dir_path:
            self.save_dir = dir_path
            self.lbl_save_dir.setText(f"Папка для сохранения: {dir_path}")
            self.check_ready()

    def check_ready(self):
        if self.file_path_input and self.file_path_data and self.save_dir:
            self.btn_start.setEnabled(True)

    def process_files(self):
        try:
            # Загружаем входной файл и преобразуем все ячейки в строки
            df_input = pd.read_excel(self.file_path_input, header=None)
            df_input = df_input.astype(object)

            # Загружаем файл с данными и тоже преобразуем все ячейки
            df_data = pd.read_excel(self.file_path_data, header=None)
            df_data = df_data.astype(object)

            # Получаем третью ячейку файла с данными
            if df_data.shape[1] < 3:
                QMessageBox.warning(self, "Ошибка", "Файл с данными должен содержать как минимум 3 столбца.")
                return
            third_cell_data = df_data.iat[0, 2]

            def clean_text(text):
                if pd.isna(text):
                    return ''
                return ' '.join(str(text).split()).lower()

            target_text = clean_text(third_cell_data)
            print(f"Ищу: '{target_text}'")  # для отладки

            result_rows = []
            count_matches = 0

            # Проходим по каждой ячейке входного файла
            for index, row in df_input.iterrows():
                for cell in row:
                    cell_text = clean_text(cell)
                    # для отладки выводим сравниваемые строки
                    print(f"Ячейка: '{cell_text}'")
                    if target_text == cell_text:
                        print("Совпадение найдено!")
                        result_rows.append(row.tolist())
                        count_matches += 1
                        break
                if count_matches >= 6:
                    break

            # Генерируем имя файла с датой и временем
            now = datetime.now()
            timestamp = now.strftime("%Y%m%d_%H%M%S")
            filename = f"Результат_{timestamp}.xlsx"
            save_path = f"{self.save_dir}/{filename}"

            # Сохраняем результат
            if result_rows:
                result_df = pd.DataFrame(result_rows)
                result_df.to_excel(save_path, index=False, header=False)
                QMessageBox.information(self, "Готово", f"Обработка завершена и сохранена: {save_path}")
            else:
                QMessageBox.information(self, "Результат", "Совпадений не найдено.")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelProcessor()
    window.show()
    sys.exit(app.exec())
#