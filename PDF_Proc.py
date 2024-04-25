#main.py
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QFileDialog, QLabel, QMessageBox, QGridLayout, QDesktopWidget, QComboBox
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
import win32com.client
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
        self.setGeometry(300, 300, 400, 200)
        self.center()

        # Create a central widget to hold other widgets
        self.centralWidget = QWidget(self)
        self.setCentralWidget(self.centralWidget)

        # Create and set layout for the central widget
        layout = QGridLayout(self.centralWidget)

        # Dropdown for selecting document type
        self.docTypeComboBox = QComboBox(self)
        self.docTypeComboBox.addItems(['Purchase Order','Invoice'])
        layout.addWidget(self.docTypeComboBox, 0, 0, 1, 1)  # Adjust grid position as needed

        self.statusLabel = QLabel('Status: Ready', self)
        self.statusLabel.setMaximumHeight(20)  # Set maximum height to 20 pixels
        layout.addWidget(self.statusLabel, 0, 1)

        

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
        layout.addWidget(self.pullPricesButton, 4, 1)
        
        self.exitButton = QPushButton('Exit', self)
        self.exitButton.clicked.connect(QApplication.instance().quit)
        self.exitButton.setMaximumWidth(100)  # Set maximum width to 100 pixels, adjust as needed
        layout.addWidget(self.exitButton, 5, 1)  # Positioned in the right column

        self.runVbaButton = QPushButton('Batch Clean', self)
        self.runVbaButton.clicked.connect(self.runVbaMacro)
        layout.addWidget(self.runVbaButton, 4, 0)

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
            self.statusLabel.setText(f'Status: Files Imported{files}')
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
            
            self.statusLabel.setText('Status: Processing...')
            self.worker = WorkerThread(self.processPDFs, self.files)
            
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
        for file_path in files:
            # Extract filename for use in paths
            filename = os.path.basename(file_path)

            # Step 1: Copy file to the 'input' directory
            shutil.copy(file_path, os.path.join(main_path, paths['input'], filename))

            # Step 2: Split PDF
            # Note: split_pdf now takes a single file path, not a directory
            self.split_pdf(os.path.join(main_path, paths['input'], filename))

            # Step 3: Analyze documents
            # analyze_general_documents now processes files directly from the 'processed' directory
            self.analyze_general_documents()

            # Step 4: Move original file to 'archive_input'
            shutil.move(os.path.join(main_path, paths['input'], filename), os.path.join(main_path, paths['archive_input'], filename))

            # Move processed files to 'archive_processed'
            # This is done inside analyze_general_documents to ensure only processed files are moved
    
    @pyqtSlot()
    def onProcessingComplete(self):
        self.processButton.setEnabled(True)
        self.statusLabel.setText('Status: Complete')
        QMessageBox.information(self, 'Complete', 'Files have been processed.')

    def split_pdf(self, file_path):
        """        Splits a PDF file into multiple pages.

        Parameters:
            file_path (str): The path to the PDF file to be split.

        Returns:
            None
        """
        # Ensure the path points to a file
        if not file_path.endswith(".pdf"):
            print(f"Error: {file_path} is not a PDF file.")
            return

        # Initialize PdfReader with the provided file path
        pdf = PdfReader(file_path)
        for i, page in enumerate(pdf.pages):
            pdf_writer = PdfWriter()
            pdf_writer.add_page(page)

            # Construct output filename for each page
            output_filename = f"{os.path.splitext(os.path.basename(file_path))[0]}_{i + 1}.pdf"
            output_filepath = os.path.join(main_path, paths['processed'], output_filename)

            # Write out each page as a separate PDF
            with open(output_filepath, 'wb') as out:
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
                shutil.move(os.path.join(main_path, paths['processed'], file), os.path.join(main_path, paths['archive_processed'], file))
                # Move processed file to '/processed/Archive' after processing
        
            
    def move_to_archive(self, path, archive_path):
        for file in os.listdir(main_path + path):
            if os.path.isfile(main_path + path + file):
                shutil.move(main_path + path + file, main_path + archive_path + file)

    def replace_import_headers(self, df):
        try:
            
            df.columns = df.columns.astype(str).str.lower()

            # Check if any of the target words are present in the header row
            target_words = ["order", "items", "quantity", "qty", "cost", "unit price", "price", "#", "amount"]
            if not any(word in df.columns for word in target_words):
                return df  # Skip replacements if no target words are found

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
            part_number_column_index = None
            part_number_pattern = re.compile(r'P\d{2}-\d{3}-\d{3}')
            for index, column in enumerate(df.columns):
                # Check if any cell in the column matches the part number pattern
                if df[column].astype(str).str.match(part_number_pattern).any():
                    part_number_column_index = index
                    break

            if part_number_column_index is not None:
                # Directly set the column name in the DataFrame's columns attribute
                df.columns.values[part_number_column_index] = 'pr_codenum'

        except Exception as e:
            print(f"Error occurred while replacing import headers: {e}")
            # Optionally, log the error or handle it as per requirement
            return df  # Return the dataframe even if an error occurs to not interrupt the flow

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

    def runVbaMacro(self):
        excel = None  # Initialize excel variable to ensure it's in scope for the finally block
        try:
            excel = win32com.client.Dispatch('Excel.Application')  # Using late binding
            excel.Visible = True  # Set to False if you don't want Excel to be visible

            # Open PERSONAL workbook
            personal_wb_path = r"C:\Users\daniel.pace\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONALwb.XLSB"
            excel.Workbooks.Open(personal_wb_path)

            # Load the add-in if it's not already loaded
            addin_path = r"C:\Users\daniel.pace\AppData\Roaming\Microsoft\AddIns\CustomMacros.xlam"
            if addin_path not in [x.FullName for x in excel.AddIns]:
                excel.Workbooks.Open(addin_path)

            # Run the macro
            macro_name = "CustomMacros.xlam!batchClean"  # Make sure the macro name is correct
            excel.Application.Run(macro_name)

        except Exception as e:
            QMessageBox.warning(self, 'Error', f'An error occurred while running the macro: {e}')

        finally:
            # Clean up
            if excel is not None:  # Check if 'excel' is defined before attempting to quit
                excel.Quit()

def main():
    app = QApplication(sys.argv)
    ex = PDFProcessingApp()
    ex.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
