import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QMenuBar, QMenu, 
                             QToolBar, QTabWidget, QWidget, QHBoxLayout, 
                             QVBoxLayout, QLineEdit, QPushButton, QComboBox,
                             QTableWidget, QAction, QFileDialog, QDialog, QLabel)

from PyQt5.QtGui import QIcon

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Invoice Processing")
        
        self.createMenuBar()
        
        self.createToolBar()

        self.createTabs()

        self.setCentralWidget(self.tabs)

    def createMenuBar(self):
      
        menuBar = self.menuBar()
        
        # File menu
        fileMenu = QMenu("File", self)
        menuBar.addMenu(fileMenu)
        
        openPdfAction = QAction("Open PDF", self) 
        openPdfAction.triggered.connect(self.openPdf)
        fileMenu.addAction(openPdfAction)

        # Settings menu
        settingsMenu = QMenu("Settings", self)
        menuBar.addMenu(settingsMenu)

        # Set input path action
        setInputPathAction = QAction("Set Input Path", self)
        setInputPathAction.triggered.connect(self.setInputPath)
        settingsMenu.addAction(setInputPathAction)

        # Set output path action
        setOutputPathAction = QAction("Set Output Path", self)
        setOutputPathAction.triggered.connect(self.setOutputPath) 
        settingsMenu.addAction(setOutputPathAction)

        # Azure Credentials action
        azureCredsAction = QAction("Set Azure Credentials", self)
        azureCredsAction.triggered.connect(self.setAzureCredentials)
        settingsMenu.addAction(azureCredsAction)
        
        # Help menu
        helpMenu = QMenu("Help", self)
        menuBar.addMenu(helpMenu)

    def createToolBar(self):
        
        toolBar = QToolBar()
        self.addToolBar(toolBar)

        browse_icon = QIcon("icons/browse.png")
        self.browseButton = QAction(QIcon(browse_icon),"Browse", self)
        self.browseButton.triggered.connect(self.browsePdf)
        toolBar.addAction(self.browseButton)
        
        convert_icon = QIcon("icons/convert.png")
        self.convertButton = QAction(QIcon(convert_icon),"Convert", self)
        toolBar.addAction(self.convertButton)

        save_excel_icon = QIcon("icons/excel.png")
        self.saveExcelButton = QAction(QIcon(save_excel_icon),"Save Excel", self)        
        toolBar.addAction(self.saveExcelButton)

    def createTabs(self):
        
        self.tabs = QTabWidget()
        self.pdfTab = QWidget()
        self.azureTab = QWidget()
       
        self.pdfTabLayout = QVBoxLayout()
        self.azureTabLayout = QVBoxLayout()

        self.pdfTab.setLayout(self.pdfTabLayout)
        self.azureTab.setLayout(self.azureTabLayout)

        self.tabs.addTab(self.pdfTab,"PDF")
        self.tabs.addTab(self.azureTab,"Azure")

    def openPdf(self):
      
        path = QFileDialog.getOpenFileName(self, 'Open PDF')[0]

        if path:
          print("Opened:", path)

    def browsePdf(self):
      
        path = QFileDialog.getOpenFileName(self, 'Browse PDF')[0]  

        if path:
          print("Browsed:", path)

    def setInputPath(self):
        dialog = SetPathDialog("Select Input Folder")
        if dialog.exec():
            self.inputPath = dialog.inputPath

    def setOutputPath(self):
        dialog = SetPathDialog("Select Output Folder")
        if dialog.exec():
            self.outputPath = dialog.inputPath
    
    def browse(self):

        path = QFileDialog.getExistingDirectory(self, "Select Output Folder")

        if path:
            self.pathInput.setText(path)

    def accept(self):
        self.inputPath = self.pathInput.text()
        super().accept()

    def setAzureCredentials(self):
        dialog = AzureCredentialsDialog()
        if dialog.exec():
            self.azureKey = dialog.azureKey
            self.azureEndpoint = dialog.azureEndpoint
            # Now you can use self.azureKey and self.azureEndpoint wherever you need them

class SetPathDialog(QDialog):

    def __init__(self, dialog_title):
        super().__init__()

        layout = QVBoxLayout()

        self.pathInput = QLineEdit()
        browseButton = QPushButton("Browse")
        browseButton.clicked.connect(self.browse)

        buttonsLayout = QHBoxLayout()
        okButton = QPushButton("OK")
        okButton.clicked.connect(self.accept)
        cancelButton = QPushButton("Cancel")
        cancelButton.clicked.connect(self.reject)

        buttonsLayout.addWidget(okButton)
        buttonsLayout.addWidget(cancelButton)

        layout.addWidget(self.pathInput)
        layout.addWidget(browseButton)
        layout.addLayout(buttonsLayout)

        self.setLayout(layout)
        self.dialog_title = dialog_title

    def browse(self):
        path = QFileDialog.getExistingDirectory(self, self.dialog_title)
        if path:
            self.pathInput.setText(path)

    def accept(self):
        self.inputPath = self.pathInput.text()
        super().accept()

class AzureCredentialsDialog(QDialog):

    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()

        # Azure API Key
        self.azureKeyInput = QLineEdit()
        azureKeyLayout = QHBoxLayout()
        azureKeyLayout.addWidget(QLabel("Azure API Key:"))
        azureKeyLayout.addWidget(self.azureKeyInput)
        layout.addLayout(azureKeyLayout)

        # Azure Endpoint
        self.azureEndpointInput = QLineEdit()
        azureEndpointLayout = QHBoxLayout()
        azureEndpointLayout.addWidget(QLabel("Azure Endpoint:"))
        azureEndpointLayout.addWidget(self.azureEndpointInput)
        layout.addLayout(azureEndpointLayout)

        # Buttons
        buttonsLayout = QHBoxLayout()
        okButton = QPushButton("OK")
        okButton.clicked.connect(self.accept)
        cancelButton = QPushButton("Cancel")
        cancelButton.clicked.connect(self.reject)
        buttonsLayout.addWidget(okButton)
        buttonsLayout.addWidget(cancelButton)
        layout.addLayout(buttonsLayout)

        self.setLayout(layout)

    def accept(self):
        self.azureKey = self.azureKeyInput.text()
        self.azureEndpoint = self.azureEndpointInput.text()
        super().accept()

# Main
app = QApplication(sys.argv)
window = MainWindow()
# Set the initial size of the window
window.resize(800, 600)
window.show() 
app.exec()