# Import the necessary libraries
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QComboBox, QFileDialog, QLabel,QDesktopWidget
from core_functions import send_pdf_to_azure
from core_functions import split_pdfs

# Create the main application window
app = QApplication([])
window = QMainWindow()

# Set the window title
window.setWindowTitle("PDF to Excel")

# Set the initial size of the window
window.resize(800, 600)  # Set the size as per your requirement

# Center the window on the screen
qtRectangle = window.frameGeometry()
centerPoint = QDesktopWidget().availableGeometry().center()
qtRectangle.moveCenter(centerPoint)
window.move(qtRectangle.topLeft())


# Create the instruction label
instruction_label = QLabel("Please select a PDF file:", window)
instruction_label.setGeometry(window.width() - 110, window.height() - 90, 100, 30)

# Create the select file button
select_file_button = QPushButton("Select File", window)
select_file_button.setGeometry(window.width() - 110, window.height() - 70, 100, 30)
convert_button = QPushButton("Convert", window)
convert_button.setGeometry(window.width() - 110, window.height() - 40, 100, 30)

# Create the instruction label for the dropdown list
dropdown_label = QLabel("Select the document type:", window)
dropdown_label.setGeometry(10, 130, 200, 30)

# Create the dropdown list
dropdown_list = QComboBox(window)
dropdown_list.setGeometry(10, 90, 100, 30)
dropdown_list.addItem("Invoice")
dropdown_list.addItem("PO")

# Define the function to handle the button click event
def select_file():
    file_dialog = QFileDialog()
    file_path = file_dialog.getOpenFileName(window, "Select File")[0]
    print("Selected file:", file_path)
    # Create the convert button
    
# Connect the button click event to the function
select_file_button.clicked.connect(select_file)

# Show the main application window
window.show()
app.exec_()
