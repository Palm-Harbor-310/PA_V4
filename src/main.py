#main.py
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QFileDialog, QLabel, QProgressBar, QMessageBox, QGridLayout, QDesktopWidget, QComboBox
)
from PyQt5.QtCore import pyqtSlot, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
import sys
import os
import shutil
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
import re

# Azure credentials and paths

main_path = "C:/Users/daniel.pace/Documents/Coding/PO Automation/Azure/Purchase Orders/"
paths = {'input': 'Input/', 'processed': 'Processed/', 'output': 'Output/', 'archive_input': 'Input/Archive/', 'archive_processed': 'Processed/Archive/'}

# Worker thread
class WorkerThread(QThread):
    finished = pyqtSignal()
    progress = pyqtSignal(int)
    def __init__(self, func, *args, **kwargs):
        """
        Initializes a new instance of the class.

        Args:
            func (callable): The function to be called.
            *args: Variable length argument list.
            **kwargs: Arbitrary keyword arguments.

        Returns:
            None
        """
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
        """
        Initializes the class instance.
        """
        super().__init__()
        self.initUI()
        

    def initUI(self):
        """
        Initializes the user interface for the PDF processing app.

        This function sets the window title, geometry, and centers the window on the screen.
        It also creates a central widget to hold other widgets and sets its layout.
        The function adds a status label, a progress bar, and two buttons to the layout.

        Parameters:
        - self: The instance of the class.

        Returns:
        - None
        """
        self.setWindowIcon(QIcon('icons\\appIcon.png'))
        self.setWindowTitle('PDF Processing App')
        self.setGeometry(300, 300, 600, 200)
        self.center()

        # Create a central widget to hold other widgets
        self.centralWidget = QWidget(self)
        self.setCentralWidget(self.centralWidget)

        # Create and set layout for the central widget
        layout = QGridLayout(self.centralWidget)

        # Dropdown for selecting document type
        self.docTypeComboBox = QComboBox(self)
        self.docTypeComboBox.addItems(['Purchase Order','Invoice'])
        layout.addWidget(self.docTypeComboBox, 0, 0, 1, 2)  # Adjust grid position as needed

        self.statusLabel = QLabel('Status: Ready', self)
        self.statusLabel.setMaximumHeight(20)  # Set maximum height to 20 pixels
        layout.addWidget(self.statusLabel, 1, 0, 1, 2)

        self.progressBar = QProgressBar(self)
        layout.addWidget(self.progressBar, 2, 0, 1, 2)

        self.importButton = QPushButton('Import Files', self)
        self.importButton.clicked.connect(self.importFiles)
        self.importButton.setMaximumWidth(100)  # Set maximum width to 100 pixels
        layout.addWidget(self.importButton, 3, 0)

        self.processButton = QPushButton('Process Files', self)
        self.processButton.clicked.connect(self.processFiles)
        layout.addWidget(self.processButton, 3, 1)

        # Add a new button for pulling prices
        self.pullPricesButton = QPushButton('Pull Prices', self)
        self.pullPricesButton.clicked.connect(self.pullPrices)
        layout.addWidget(self.pullPricesButton, 4, 0)
        
        self.exitButton = QPushButton('Exit', self)
        self.exitButton.clicked.connect(QApplication.instance().quit)
        self.exitButton.setMaximumWidth(100)  # Set maximum width to 100 pixels, adjust as needed
        layout.addWidget(self.exitButton, 4, 1)  # Positioned in the right column


    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    @pyqtSlot()
    def importFiles(self):
        """
        Import files using a file dialog.

        This function opens a file dialog to allow the user to select one or more PDF files.
        The selected files are stored as an instance variable.

        Parameters:
        - None

        Returns:
        - None
        """
        # File import logic goes here
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(self, "Choose a PDF file", "", "PDF Files (*.pdf)", options=options)
        if files:
            self.files = files  # Store the files as an instance variable
            self.statusLabel.setText('Status: Files Imported')
    @pyqtSlot()
    def processFiles(self):
        # Set endpoint and key based on document type selection
        doc_type = self.docTypeComboBox.currentText()
        if doc_type == 'Invoice':
            endpoint = "https://phh-invoices.cognitiveservices.azure.com/"
            key = os.getenv("AZURE_API_KEY_PHH-INVOICES")
        else:  # Purchase Order
            endpoint = "https://pos.cognitiveservices.azure.com/"
            key = os.getenv("AZURE_API_KEY_POS")
        
        # Create the credential and client with the selected endpoint and key
        credential = AzureKeyCredential(key)
        self.client = DocumentAnalysisClient(endpoint=endpoint, credential=credential)

        if hasattr(self, 'files') and self.files:        
            self.processButton.setEnabled(False)
            self.progressBar.setValue(0)
            self.statusLabel.setText('Status: Processing...')
            self.worker = WorkerThread(self.processPDFs, self.files)
            self.worker.progress.connect(self.updateProgressBar)  # Connect progress signal to updateProgressBar slot
            self.worker.finished.connect(self.onProcessingComplete)  # Connect finished signal to onProcessingComplete slot
            self.worker.start()
        else:
            QMessageBox.warning(self, 'Warning', 'No files have been imported.')

    def processPDFs(self, files):
        """
        Process a list of PDF files.

        Args:
            files (list): A list of file paths to the PDF files.

        Returns:
            None
        """
        # Get the selected document type
        total_files = len(files)
        total_steps = 4  # Number of steps in the processing of each file
        progress_per_step = 100 / (total_files * total_steps)

        for index, file_path in enumerate(files):
            # Step 1: Copy file
            
            self.emitProgress(progress_per_step, 'Copying file')
            shutil.copy(file_path, main_path + paths['input'])

            # Step 2: Split PDF
            self.split_pdf(main_path + paths['input'])
            self.emitProgress(progress_per_step, 'Splitting PDF')

            # Step 3: Analyze documents
            self.emitProgress(progress_per_step, 'Analyzing documents')
            self.analyze_general_documents()

            # Step 4: Move to archive
            self.emitProgress(progress_per_step, 'Moving to archive')
            self.move_to_archive(paths['input'], paths['archive_input'])
            self.move_to_archive(paths['processed'], paths['archive_processed'])
            

        self.worker.finished.emit()  # Signal that processing is complete

    
    def emitProgress(self, progress_increment, status_message):
        """
        Emit the progress signal with the updated value and update the status label.

        Parameters:
            progress_increment (int): The amount by which the progress bar should be incremented.
            status_message (str): The new status message to be displayed.

        Returns:
            None
        """
        # Emit the progress signal with the updated value
        current_progress = self.progressBar.value()
        new_progress = min(current_progress + progress_increment, 100)
        self.worker.progress.emit(new_progress)
        # Update the status label
        self.statusLabel.setText(f'Status: {status_message}')

        
    @pyqtSlot(int)
    def updateProgressBar(self, value):
        """
        Update the progress bar with the given value.

        Parameters:
            value (int): The value to set the progress bar to.

        Returns:
            None
        """
        self.progressBar.setValue(value)

    @pyqtSlot()
    def onProcessingComplete(self):
        self.processButton.setEnabled(True)
        self.statusLabel.setText('Status: Complete')
        QMessageBox.information(self, 'Complete', 'Files have been processed.')

    def split_pdf(self, path):
        """
        Splits a PDF file into multiple pages.

        Parameters:
            path (str): The path to the directory containing the PDF files.

        Returns:
            None
        """
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
        files = [f for f in os.listdir(main_path + paths['processed']) if f.endswith(".pdf")]
        total_files = len(files)
        for i, file in enumerate(files):
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
            # Emit progress after each file is processed
            progress_increment = 100 / total_files
            self.emitProgress(progress_increment, 'Analyzing documents')
    
    def move_to_archive(self, path, archive_path):
        for file in os.listdir(main_path + path):
            if os.path.isfile(main_path + path + file):
                shutil.move(main_path + path + file, main_path + archive_path + file)

    def replace_import_headers(self, df):
        # Lowercase column names for uniformity
        df.columns = df.columns.astype(str).str.lower()

        # Initially handle 'amount' specifically
        if 'amount' in df.columns:
            index = df.columns.tolist().index('amount')
            if index == 0:
                df.rename(columns={'amount': 'quantity'}, inplace=True)
            elif index == len(df.columns) - 1:
                df.rename(columns={'amount': 'total'}, inplace=True)

        # Define replacements after handling 'amount'
        replacements = {k: v for k, v in zip(
            ["order", "items", "quantity", "qty", "cost", "unit price", "price", "#"], 
            ["pu_quant"] * 4 + ["pu_price"] * 4)}

        # Apply the replacements for other headers
        df.rename(columns={col: replacements.get(col, col) for col in df.columns}, inplace=True)

        # Part number identification and renaming
        part_number_column = None
        part_number_pattern = re.compile(r'P\d{2}-\d{3}-\d{3}')
        for column in df.columns:
            if df[column].astype(str).str.match(part_number_pattern).any():
                part_number_column = column
                break

        if part_number_column:
            df.rename(columns={part_number_column: 'pr_codenum'}, inplace=True)

        return df

    def extract_invoice_data(self, file_path):
        """Extract invoice data using Azure Form Recognizer"""
        
        errors = []
        invoice_data = None
        
        try:
            with open(file_path, "rb") as fd:
                document = fd.read()

            poller = self.client.begin_analyze_document("prebuilt-invoice", document)  
            result = poller.result()
            
            invoice_data = result.documents[0].fields

        except Exception as e:
            errors.append(f"Error processing {file_path}: {e}")
            
        if errors:
            self.log_errors(errors)
            
        return invoice_data

    @pyqtSlot()
    def pullPrices(self):
        """
        Slot method to handle the 'Pull Prices' button click.
        """
        try:
            self.pullPricesFromExcel()
            QMessageBox.information(self, 'Success', 'Prices have been pulled successfully.')
        except Exception as e:
            QMessageBox.warning(self, 'Error', f'An error occurred: {e}')



    def pullPricesFromExcel(self):
        """
        Method to pull prices from Excel files in a folder and process them.

        Args:
        folder_path (str): Path to the folder containing Excel files.
        """
        folder_path = "P:/Temp"
        # Get a list of all Excel files in the folder
        excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
        for file in excel_files:
            try:
                # Construct full path to the current file
                file_path = os.path.join(folder_path, file)

                # Read the current Excel file
                original_df = pd.read_excel(file_path, engine='openpyxl')

                # Load the LPC Excel file
                lpc_df = pd.read_excel(
                    r"C:\Users\daniel.pace\Documents\Coding\Purchasing Automation\PA_V4\Price Comparison\Last Part Cost 3.1.xlsx",
                    sheet_name='All')
                
                # Map PO Cost from LPC to original dataframe based on part numbers
                original_df = original_df.merge(lpc_df[['pr_codenum', 'PO Cost']], on='pr_codenum', how='left')

                # Save the processed data to a new file with the original filename
                original_df.to_excel(os.path.join(folder_path, f"processed_{file}"), index=False)
            except Exception as e:
                print(f"Error processing {file}: {e}")
                continue


def main():
    app = QApplication(sys.argv)
    ex = PDFProcessingApp()
    ex.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
