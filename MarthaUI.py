import re
import sys

from PyQt6.QtCore import QRegularExpression
from PyQt6.QtWidgets import QApplication, QDialog, QLineEdit, QPushButton, QGridLayout, QLabel, QComboBox, QDateEdit
from PyQt6.QtGui import QRegularExpressionValidator, QPalette, QColor
from Util import Constants




class Form(QDialog):

    def __init__(self, parent=None):
        super(Form, self).__init__(parent)
        # Create widgets
        self.labelDropdown = QLabel("Kasualie")
        self.kasualieDropdown = QComboBox()
        self.kasualieDropdown.addItem("Taufe")
        self.kasualieDropdown.addItem("Segnung")



        self.button = QPushButton("Create Documents")
        # Create layout and add widgets
        layout = QGridLayout()
        layout.addWidget(self.kasualieDropdown, 0, 0)
        self.setUpTaufe(layout)
        layout.addWidget(self.button)



    # Set dialog layout
        self.setLayout(layout)
        # Add button signal to greetings slot
        self.button.clicked.connect(self.createDocuments)

        self.setWindowTitle("Martha.py")



    def setUpTaufe(self, layout):
        # Widgets für Taufe
        self.labelTäufling = QLabel(Constants.täufling)
        self.nameTäufling = QLineEdit()
       # rx = QRegularExpression("[A-Z][a-z]{100}")
       # validator = QRegularExpressionValidator(rx)
       # self.nameTäufling.setValidator(validator)

        self.nameTäufling.setPlaceholderText(Constants.nachname)
        self.firstNameTäufling = QLineEdit()
        self.firstNameTäufling.setPlaceholderText(Constants.vorname)
        self.labelDateOfBirth = QLabel(Constants.dateOfBirth)
        self.dateOfBirthTäufling = QDateEdit()
        self.placeOfBirthTäufingLabel = QLabel(Constants.placeOfBirth)
        self.placeOfBirthTäufling = QLineEdit()
        self.placeOfBirthTäufling.setPlaceholderText(Constants.placeOfBirth)

        layout.addWidget(self.labelTäufling, 1, 0)
        layout.addWidget(self.nameTäufling, 2, 0)
        layout.addWidget(self.firstNameTäufling, 2, 1)
        layout.addWidget(self.labelDateOfBirth)
        layout.addWidget(self.dateOfBirthTäufling)
        layout.addWidget(self.placeOfBirthTäufingLabel)
        layout.addWidget(self.placeOfBirthTäufling)









    def setUpKasualie2(self):
        # Widgets für andere Kasualie
        print("hallo")
    # Greets the user
    def createDocuments(self):
        #print(f"Hello {self.edit.text()}")
        data = self.placeOfBirthTäufling.text() + self.dateOfBirthTäufling.text()
        print(data)



if __name__ == '__main__':
    # Create the Qt Application
    app = QApplication(sys.argv)
    # Create and show the form
    form = Form()
    form.show()
    # Run the main Qt loop
    sys.exit(app.exec())

