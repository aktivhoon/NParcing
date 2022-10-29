from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIntValidator,QDoubleValidator,QFont
from PyQt5.QtCore import Qt
import sys
import pandas as pd
from io import StringIO
from parcing import generate_excel 


class MainWindow(QWidget):
    def __init__(self,parent=None):
        super().__init__(parent)
        self.e1 = QLineEdit()
        self.e1.setFixedWidth(300)
        self.e1.setFixedHeight(100)
        self.e2 = QLineEdit()
        self.e2.setFixedWidth(300)
        self.e2.setFixedHeight(100)

        self.b1 = QPushButton()
        self.b1.setText("Generate")
        self.b1.setShortcut('Enter')
        self.b1.clicked.connect(self.execute)

        flo = QFormLayout()
        flo.addRow("Admission",self.e1)
        flo.addRow("Discharge",self.e2)
        flo.addRow(self.b1)

        self.setLayout(flo)
        self.setWindowTitle("NParcing by. YHKim")

    def enterPress(self):
        print("Enter pressed")

    def execute(self):
        adm_text = StringIO(self.e1.text())
        dc_text = StringIO(self.e2.text())
        generate_excel(adm_text, dc_text)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
