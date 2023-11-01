import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QMenuBar, QMenu, 
                            QToolBar, QTabWidget, QWidget, QHBoxLayout, 
                            QVBoxLayout, QLineEdit, QPushButton, QComboBox,
                            QTableWidget, QAction, QFileDialog, QDialog, QLabel,
                            QDialogButtonBox,QTreeView, QFileSystemModel, QSplitter,
                            QDockWidget)

from PyQt5.QtGui import QIcon
from PyQt5.QtCore import (Qt,QDir)

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Invoice Processing")
        
        # Initialize the label in the constructor
        self.fileIndicatorLabel = QLabel("No file loaded")
        self.fileIndicatorLabel.setMaximumHeight(30)

        self.createMenuBar()
        
        self.createToolBar()

        self.createTabs()

        self.createButtonBox()

        self.createNavigationPane()
        
        self.setupCentralLayout()

        self.setupButtonLayout()

        
        
    


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
    
    def createButtonBox(self):
        
        # Create the button box with "Next" and custom "Import" buttons
        self.buttonBox = QDialogButtonBox()
        nextButton = self.buttonBox.addButton("Next", QDialogButtonBox.AcceptRole)
        exitButton = self.buttonBox.addButton("Exit", QDialogButtonBox.RejectRole)

        # Connect the signals to slots
        nextButton.clicked.connect(self.onNext)
        exitButton.clicked.connect(self.close)
        # Create "Import" button
        buttonBoxLayout = QHBoxLayout()
        self.importButton = QPushButton("Import")
        self.importButton.clicked.connect(self.onImport)
        buttonBoxLayout.addWidget(self.importButton, alignment=Qt.AlignBottom | Qt.AlignLeft)

        # Add QDialogButtonBox to the layout
        buttonBoxLayout.addWidget(self.buttonBox)

    def setupCentralLayout(self):
    # Create a QWidget for the central widget
        centralWidget = QWidget(self)

        # Create a layout
        layout = QVBoxLayout()

        # Create a splitter
        splitter = QSplitter(Qt.Horizontal)

        # Add the navigation pane and the tabs to the splitter
        splitter.addWidget(self.fileExplorer)
        splitter.addWidget(self.tabs)
        splitter.setSizes([100, 400])

        
        layout.addWidget(self.fileIndicatorLabel)
        layout.addWidget(splitter)

        # Set the layout of the central widget
        centralWidget.setLayout(layout)

        # Set the central widget of the main window
        self.setCentralWidget(centralWidget)

    def setupButtonLayout(self):
        # Create a widget for the buttons
        buttonWidget = QWidget()

        # Create a layout for the buttons
        buttonBoxLayout = QHBoxLayout()

        # Add Import button and button box to the layout
        buttonBoxLayout.addWidget(self.importButton, alignment=Qt.AlignLeft)
        buttonBoxLayout.addWidget(self.buttonBox)

        # Set the layout of the button widget
        buttonWidget.setLayout(buttonBoxLayout)

        # Create a dock widget, set it to the button widget and add it to the main window
        dockWidget = QDockWidget()
        dockWidget.setWidget(buttonWidget)
        self.addDockWidget(Qt.BottomDockWidgetArea, dockWidget)
    
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

    # Slot for "Next" button
    def onNext(self):
        print("Next button clicked")

    # Slot for "Cancel" button
    def onExit(self):
        self.close()

    # Slot for "Import" button
    def onImport(self):
        path, _ = QFileDialog.getOpenFileName(self, 'Import File')
        if path:
            print(f"Imported: {path}")
            # Update the label text with the name of the imported file
            self.fileIndicatorLabel.setText(f"Loaded: {path}")

    def createNavigationPane(self):
        self.fileExplorer = FileExplorer()
        self.fileExplorer.tree.clicked.connect(self.onFileClicked)

    def onFileClicked(self, index):
        path = self.fileExplorer.model.fileInfo(index).absoluteFilePath()
        print(f"File clicked: {path}")

    

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

class FileExplorer(QWidget):
    def __init__(self):
        super().__init__()
        self.tree = QTreeView(self)
        self.model = QFileSystemModel()
        self.model.setRootPath(QDir.rootPath())
        self.tree.setModel(self.model)
        self.tree.setRootIndex(self.model.index(QDir.rootPath()))
        layout = QVBoxLayout(self)


        
        layout.addWidget(self.tree)



# Main
app = QApplication(sys.argv)
window = MainWindow()
# Set the initial size of the window
window.resize(800, 600)
window.show() 
app.exec()