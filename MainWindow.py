import sys
import pandas as pd
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import QFileDialog, QProgressBar, QMessageBox
from DatabaseHandler import DatabaseHandler
from Converter import ExcelConverter

class WorkerThread(QtCore.QThread):
    progress_updated = QtCore.pyqtSignal(int)
    conversion_finished = QtCore.pyqtSignal(bool)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        converter = ExcelConverter()

        try:
            df_excel = pd.read_excel(self.file_path, sheet_name=None, skiprows=4)
            db_handler = DatabaseHandler()

            total_sheets = len(df_excel)
            processed_sheets = 0

            for sheet_name, df in df_excel.items():
                df = df.dropna(axis=1, how='all')
                if df.empty:
                    continue

                try:
                    converted_data = converter.convert_to_dataframe(df)
                    db_handler.insert_data(sheet_name, converted_data)
                except Exception as e:
                    print(f"Error processing {sheet_name}: {str(e)}")

                processed_sheets += 1
                progress_value = int((processed_sheets / total_sheets) * 100)
                self.progress_updated.emit(progress_value)

            print("Conversion successful!")
            self.conversion_finished.emit(True)

        except Exception as e:
            print(f"Error reading or processing Excel file: {str(e)}")
            self.conversion_finished.emit(False)

class MainWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Excel to SQL Converter")

        self.file_path_line_edit = QtWidgets.QLineEdit(self)
        self.file_path_line_edit.setGeometry(10, 10, 300, 30)
        self.file_path_line_edit.setReadOnly(True)

        browse_button = QtWidgets.QPushButton('Browse', self)
        browse_button.setGeometry(320, 10, 80, 30)
        browse_button.clicked.connect(self.browse_file)

        self.convert_button = QtWidgets.QPushButton('Convert', self)
        self.convert_button.setGeometry(10, 50, 80, 30)
        self.convert_button.clicked.connect(self.start_conversion)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setGeometry(10, 90, 390, 20)

        self.worker_thread = None

    def browse_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Choose Excel File", "", "Excel Files (*.xlsx);;All Files (*)",
                                                   options=options)

        if file_name:
            self.file_path_line_edit.setText(file_name)
            self.reset_progress_bar()

    def start_conversion(self):
        file_path = self.file_path_line_edit.text()
        self.convert_button.setEnabled(False)

        self.worker_thread = WorkerThread(file_path)
        self.worker_thread.progress_updated.connect(self.update_progress)
        self.worker_thread.conversion_finished.connect(self.conversion_finished)
        self.worker_thread.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def conversion_finished(self, success):
        self.convert_button.setEnabled(True)

        if success:
            QMessageBox.information(self, "Success", "Conversion successful!")
        else:
            QMessageBox.critical(self, "Error", "Error reading or processing Excel file.")

    def reset_progress_bar(self):
        self.progress_bar.setValue(0)

def run_main_window():
    app = QtWidgets.QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
