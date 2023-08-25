#-*- coding: utf-8 -*-
import sys
import pandas as pd
from urllib.parse import urlencode
from urllib.request import urlopen
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QFileDialog, QVBoxLayout, QWidget, QProgressBar


def save_error_log(compound_name, error_message):
    with open('error_log.txt', 'a') as log_file:
        log_file.write(f"Error for compound '{compound_name}': {error_message}\n")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Compound Information Updater")
        self.setGeometry(100, 100, 400, 250)

        self.label = QLabel("Select an Excel file:")
        self.btn_select_file = QPushButton("Select File")
        self.btn_select_file.clicked.connect(self.select_file)
        self.btn_run = QPushButton("Run")
        self.btn_run.setEnabled(False)
        self.btn_run.clicked.connect(self.run_process)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.btn_select_file)
        layout.addWidget(self.btn_run)
        layout.addWidget(self.progress_bar)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def select_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)", options=options)
        if file_name:
            self.file_name = file_name
            self.btn_run.setEnabled(True)
            self.label.setText(f"Selected file: {self.file_name}")

    def run_process(self):
        df = pd.read_excel(self.file_name)
        apiurl = 'https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/SDF'
        compound_names = df["Compound Name"].tolist()
        
        total_compounds = len(compound_names)
        self.progress_bar.setMaximum(total_compounds)
        
        for index, compound_name in enumerate(compound_names):
            try:
                postdata = urlencode([('name', compound_name)]).encode('utf8')
                response = urlopen(apiurl, postdata).read().decode('utf-8')
                df.loc[df["Compound Name"] == compound_name, 'mol'] = response[:response.index('M  END')]
                self.label.setText(f"Complete {compound_name}")
            except Exception as e:
                error_message = str(e)
                save_error_log(compound_name, error_message)
            self.progress_bar.setValue(index + 1)

        updated_file_name = self.file_name.replace(".xlsx", "_updated.xlsx")
        df.to_excel(updated_file_name, index=False, engine='openpyxl')
        self.label.setText("작업이 완료되었습니다.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
