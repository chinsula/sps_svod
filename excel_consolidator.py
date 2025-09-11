import sys

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QMessageBox, QComboBox
)
import os
import pandas as pd

class ExcelProcessor(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Обработка Excel файла")
        self.resize(400, 250)

        self.layout = QVBoxLayout()

        self.label_input = QLabel("Выберите входной файл Excel")
        self.btn_browse_input = QPushButton("Выбрать файл для обработки")
        self.btn_browse_input.clicked.connect(self.browse_input)

        self.label_output = QLabel("Выберите папку для сохранения")
        self.btn_browse_output = QPushButton("Выбрать папку")
        self.btn_browse_output.clicked.connect(self.browse_output)

        self.label_column = QLabel("Выберите столбец для консолидации (по сути, столбец, по которому происходит объединение)")
        self.combo_column = QComboBox()
        self.combo_column.setEnabled(False)  # Пока не выбран файл

        self.btn_process = QPushButton("Обработать файл")
        self.btn_process.clicked.connect(self.process_file)

        self.layout.addWidget(self.label_input)
        self.layout.addWidget(self.btn_browse_input)
        self.layout.addWidget(self.label_output)
        self.layout.addWidget(self.btn_browse_output)
        self.layout.addWidget(self.label_column)
        self.layout.addWidget(self.combo_column)
        self.layout.addWidget(self.btn_process)

        self.setLayout(self.layout)

        self.input_file = None
        self.output_folder = None
        self.selected_column = None

    def browse_input(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Выберите Excel файл", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.input_file = file_name
            self.label_input.setText(f"Входной файл: {file_name}")
            # После выбора файла пытаемся определить количество столбцов
            try:
                df = pd.read_excel(self.input_file, header=None)
                col_count = df.shape[1]
                self.combo_column.clear()
                for i in range(1, col_count + 1):
                    self.combo_column.addItem(str(i))
                self.combo_column.setEnabled(True)
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось прочитать файл: {e}")
                self.combo_column.setEnabled(False)

    def browse_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения")
        if folder:
            self.output_folder = folder
            self.label_output.setText(f"Папка для сохранения: {folder}")

    def process_file(self):
        if not self.input_file or not self.output_folder:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, выберите оба файла и папку.")
            return
        if not self.combo_column.isEnabled():
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, выберите файл и подгрузите список столбцов.")
            return

        # Получаем выбранный столбец (учитывая +1)
        self.selected_column = int(self.combo_column.currentText())

        try:
            df = pd.read_excel(self.input_file, header=None)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка чтения файла", str(e))
            return

        result_rows = []
        prev_row = None

        for index, row in df.iterrows():
            row = row.fillna('')
            second_cell = row.iloc[1]
            eleven_cell_value = row.iloc[10]  # 11-й столбец

            # Используем выбранный столбец для сравнения
            compare_cell = row.iloc[self.selected_column - 1]

            if prev_row is not None:
                prev_compare_cell = prev_row.iloc[self.selected_column - 1]
                prev_eleven_cell_value = prev_row.iloc[10] if len(prev_row) > 10 else None

                if compare_cell == prev_compare_cell:
                    if pd.isna(prev_eleven_cell_value) or prev_eleven_cell_value == '':
                        new_row_values = []
                        for i in range(1, 9):
                            prev_value = prev_row.iloc[i]
                            curr_value = row.iloc[i]
                            try:
                                prev_num = float(prev_value) if prev_value != '' else 0
                            except:
                                prev_num = 0
                            try:
                                curr_num = float(curr_value) if curr_value != '' else 0
                            except:
                                curr_num = 0
                            summed = prev_num + curr_num
                            if isinstance(prev_value, int) and isinstance(curr_value, int):
                                summed = int(summed)
                            new_row_values.append(summed)
                        consolidated_row = [
                            prev_row.iloc[0],  # первый столбец
                            compare_cell,
                            *new_row_values[1:8],
                            row.iloc[9],  # 10-й столбец из текущей строки
                            row.iloc[10]  # 11-й столбец из текущей строки
                        ]
                        prev_row = pd.Series(consolidated_row)

                        # Условие: если 11-й столбец пуст, текущая строка становится новой "предыдущей"
                        if pd.isna(row.iloc[10]) or row.iloc[10] == '':
                            prev_row = row
                        continue

            # Если не объединяем, просто добавляем предыдущую строку
            if prev_row is not None:
                result_rows.append(prev_row)
            prev_row = row

        # В конце добавляем последнюю строку
        if prev_row is not None:
            result_rows.append(prev_row)

        result_df = pd.DataFrame(result_rows)

        save_path = os.path.join(self.output_folder, f"processed_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        try:
            result_df.to_excel(save_path, index=False, header=None)
            QMessageBox.information(self, "Готово", f"Обработка завершена. Файл сохранён: {save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка сохранения", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelProcessor()
    window.show()
    sys.exit(app.exec())