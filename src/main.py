from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QFileDialog, QLabel, QProgressBar, QMessageBox, QGridLayout, QDesktopWidget
)
from PyQt5.QtCore import pyqtSlot, QThread, pyqtSignal
import sys
import os
import shutil
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

# Azure credentials and paths
endpoint = "https://pos.cognitiveservices.azure.com/"
key = os.getenv("AZURE_API_KEY_POS")
credential = AzureKeyCredential(key)
main_path = "C:/Users/daniel.pace/Documents/Coding/PO Automation/Azure/Purchase Orders/"
paths = {'input': 'Input/', 'processed': 'Processed/', 'output': 'Output/', 'archive_input': 'Input/Archive/', 'archive_processed': 'Processed/Archive/'}

# Worker thread
class WorkerThread(QThread):
    finished = pyqtSignal()

    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs

    def run(self):
        self.func(*self.args, **self.kwargs)
        self.finished.emit()

# Main Application Window
class PDFProcessingApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.client = DocumentAnalysisClient(endpoint=endpoint, credential=credential)

    def initUI(self):
        self.setWindowTitle('PDF Processing App')
        self.setGeometry(300, 300, 600, 200)
        self.center()

        # Create a central widget to hold other widgets
        self.centralWidget = QWidget(self)
        self.setCentralWidget(self.centralWidget)

        # Create and set layout for the central widget
        layout = QVBoxLayout(self.centralWidget)

        self.progressBar = QProgressBar(self)
        layout.addWidget(self.progressBar)

        self.processButton = QPushButton('Process Files', self)
        self.processButton.clicked.connect(self.processFiles)
        layout.addWidget(self.processButton)

        self.statusLabel = QLabel('Status: Ready', self)
        layout.addWidget(self.statusLabel)


    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    @pyqtSlot()
    def processFiles(self):
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(self, "Choose a PDF file", "", "PDF Files (*.pdf)", options=options)
        if files:
            self.processButton.setEnabled(False)
            self.progressBar.setValue(0)
            self.statusLabel.setText('Status: Processing...')
            self.worker = WorkerThread(self.processPDFs, files)
            self.worker.finished.connect(self.onProcessingComplete)
            self.worker.start()

    def processPDFs(self, files):
        for file_path in files:
            shutil.copy(file_path, main_path + paths['input'])
        self.split_pdf(main_path + paths['input'])
        self.analyze_general_documents()
        self.move_to_archive(paths['input'], paths['archive_input'])
        self.move_to_archive(paths['processed'], paths['archive_processed'])
        self.progressBar.setValue(100)

    @pyqtSlot()
    def onProcessingComplete(self):
        self.processButton.setEnabled(True)
        self.statusLabel.setText('Status: Complete')
        QMessageBox.information(self, 'Complete', 'Files have been processed.')

    def split_pdf(self, path):
        for file in os.listdir(path):
            if file.endswith(".pdf"):
                pdf = PdfReader(path + file)
                for i, page in enumerate(pdf.pages):
                    pdf_writer = PdfWriter()
                    pdf_writer.add_page(page)
                    output_filename = f"{paths['processed']}{file[:-4]}_{i + 1}.pdf"
                    with open(main_path + output_filename, 'wb') as out:
                        pdf_writer.write(out)

    def extract_table_data(self, table):
        table_data = []
        row_data = []
        current_row_index = 0
        for cell in sorted(table.cells, key=lambda c: (c.row_index, c.column_index)):
            if cell.row_index != current_row_index:
                table_data.append(row_data)
                row_data = []
                current_row_index = cell.row_index
            row_data.append(cell.content)
        table_data.append(row_data)  # Append the last row
        return table_data

    def analyze_general_documents(self):
        for file in os.listdir(main_path + paths['processed']):
            if file.endswith(".pdf"):
                with open(main_path + paths['processed'] + file, "rb") as fd:
                    document = fd.read()
                poller = self.client.begin_analyze_document("prebuilt-document", document)
                result = poller.result()
                for i, table in enumerate(result.tables):
                    df = pd.DataFrame(self.extract_table_data(table))
                    df.columns = df.iloc[0]
                    df = df.drop(df.index[0])
                    df = self.replace_import_headers(df)
                    df.to_excel(f"{main_path}{paths['output']}{os.path.splitext(file)[0]}_table_{i}.xlsx", index=False)

    def move_to_archive(self, path, archive_path):
        for file in os.listdir(main_path + path):
            if os.path.isfile(main_path + path + file):
                shutil.move(main_path + path + file, main_path + archive_path + file)

    def replace_import_headers(self, df):
        replacements = {k: v for k, v in zip(
            ["order", "items", "quantity", "qty", "deacom#", "deacom", "deacom #", "deacom  #", "cost", "unit price", "price", "#"], 
            ["pu_quant"] * 4 + ["pr_codenum"] * 4 + ["pu_price"] * 4)}
        df.columns = df.columns.astype(str).str.lower()
        df.rename(columns={col: replacements.get(col, col) for col in df.columns}, inplace=True)
        return df

def main():
    app = QApplication(sys.argv)
    ex = PDFProcessingApp()
    ex.show()
    sys.exit(app.exec_())



if __name__ == "__main__":
    main()
