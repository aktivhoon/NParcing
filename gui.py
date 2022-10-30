from PyQt5.QtWidgets import *
import sys
from io import StringIO
from parcing import generate_excel 
import os

class MainWindow(QWidget):
    def __init__(self,parent=None):
        super().__init__(parent)
        self.e1 = QLineEdit()
        self.e1.setFixedWidth(300)
        self.e1.setFixedHeight(100)
        self.e2 = QLineEdit()
        self.e2.setFixedWidth(300)
        self.e2.setFixedHeight(100)
        self._61_layout = QHBoxLayout()
        self._61_empty = QLineEdit()
        self._61_empty.setFixedWidth(30)
        self._61_empty.setFixedHeight(30)
        self._61_man = QLineEdit()
        self._61_man.setFixedWidth(30)
        self._61_man.setFixedHeight(30)
        self._61_woman = QLineEdit()
        self._61_woman.setFixedWidth(30)
        self._61_woman.setFixedHeight(30)
        self._61_layout.addWidget(self._61_empty)
        self._61_layout.addWidget(QLabel("남자 대기"))
        self._61_layout.addWidget(self._61_man)
        self._61_layout.addWidget(QLabel("여자 대기"))
        self._61_layout.addWidget(self._61_woman)
        self._62_layout = QHBoxLayout()
        self._62_empty = QLineEdit()
        self._62_empty.setFixedWidth(30)
        self._62_empty.setFixedHeight(30)
        self._62_man = QLineEdit()
        self._62_man.setFixedWidth(30)
        self._62_man.setFixedHeight(30)
        self._62_woman = QLineEdit()
        self._62_woman.setFixedWidth(30)
        self._62_woman.setFixedHeight(30)
        self._62_layout.addWidget(self._62_empty)
        self._62_layout.addWidget(QLabel("남자 대기"))
        self._62_layout.addWidget(self._62_man)
        self._62_layout.addWidget(QLabel("여자 대기"))
        self._62_layout.addWidget(self._62_woman)


        self.b1 = QPushButton()
        self.b1.setText("Generate")
        self.b1.setShortcut('Enter')
        self.b1.clicked.connect(self.execute)

        flo = QFormLayout()
        flo.addRow("Admission",self.e1)
        flo.addRow("Discharge",self.e2)
        flo.addRow("61병동 공실수",self._61_layout)
        flo.addRow("62병동 공실수",self._62_layout)
        flo.addRow(self.b1)

        self.setLayout(flo)
        self.setWindowTitle("NParcing by. YHKim")

    def execute(self):
        adm_text = StringIO(self.e1.text())
        dc_text = StringIO(self.e2.text())
        generate_excel(adm_text, dc_text, self._61_empty.text(),self._61_man.text(),self._61_woman.text(),self._62_empty.text(),self._62_man.text(),self._62_woman.text())
        QMessageBox.about(self,'작업 완료','당직표 파일이 생성되었습니다!\n엑셀이 실행됩니다.')
        os.startfile('dangjik.xlsx')
        self.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())