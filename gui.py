from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import sys
from io import StringIO
from parsing import generate_excel
import os

class MainWindow(QWidget):
    def __init__(self,parent=None):
        super().__init__(parent)
        self.admission_text = QLineEdit()
        self.admission_text.setFixedWidth(350)
        self.admission_text.setFixedHeight(100)
        self.discharge_text = QLineEdit()
        self.discharge_text.setFixedWidth(350)
        self.discharge_text.setFixedHeight(100)
        self._61_layout = QHBoxLayout()
        self._61_empty = QLineEdit()
        self._61_empty.setFixedWidth(30)
        self._61_empty.setFixedHeight(30)
        self._61_empty.setValidator(QIntValidator(0,100,self))
        self._61_man = QLineEdit()
        self._61_man.setFixedWidth(30)
        self._61_man.setFixedHeight(30)
        self._61_man.setValidator(QIntValidator(0,100,self))
        self._61_woman = QLineEdit()
        self._61_woman.setFixedWidth(30)
        self._61_woman.setFixedHeight(30)
        self._61_woman.setValidator(QIntValidator(0,100,self))
        self._61_layout.addWidget(QLabel("61병동 공실수"))
        self._61_layout.addWidget(self._61_empty)
        self._61_layout.addWidget(QLabel("남자 대기"))
        self._61_layout.addWidget(self._61_man)
        self._61_layout.addWidget(QLabel("여자 대기"))
        self._61_layout.addWidget(self._61_woman)
        self._62_layout = QHBoxLayout()
        self._62_empty = QLineEdit()
        self._62_empty.setFixedWidth(30)
        self._62_empty.setFixedHeight(30)
        self._62_empty.setValidator(QIntValidator(0,100,self))
        self._62_man = QLineEdit()
        self._62_man.setFixedWidth(30)
        self._62_man.setFixedHeight(30)
        self._62_man.setValidator(QIntValidator(0,100,self))
        self._62_woman = QLineEdit()
        self._62_woman.setFixedWidth(30)
        self._62_woman.setFixedHeight(30)
        self._62_woman.setValidator(QIntValidator(0,100,self))
        self._62_layout.addWidget(QLabel("62병동 공실수"))
        self._62_layout.addWidget(self._62_empty)
        self._62_layout.addWidget(QLabel("남자 대기"))
        self._62_layout.addWidget(self._62_man)
        self._62_layout.addWidget(QLabel("여자 대기"))
        self._62_layout.addWidget(self._62_woman)

        self.b1 = QPushButton()
        self.b1.setText("Generate")
        self.b1.setShortcut('Enter')
        self.b1.clicked.connect(self.execute)

        #Checkbox
        self.tg_admission = QCheckBox("입원 환자 없음")
        self.tg_admission.adjustSize()
        self.tg_admission.setChecked(False)
        self.tg_admission.toggled[bool].connect(self.no_admission)

        self.tg_discharge = QCheckBox("퇴원 환자 없음")
        self.tg_discharge.adjustSize()
        self.tg_discharge.setChecked(False)
        self.tg_discharge.toggled[bool].connect(self.no_discharge)

        flo = QFormLayout()
        flo.addRow("Admission",self.tg_admission)
        flo.addRow(self.admission_text)
        flo.addRow("Discharge",self.tg_discharge)
        flo.addRow(self.discharge_text)
        flo.addRow(self._61_layout)
        flo.addRow(self._62_layout)
        flo.addRow(self.b1)

        self.setLayout(flo)
        self.setWindowTitle("NParsing by. YHKim, JHChung")

    def execute(self):
        adm_text = StringIO(self.admission_text.text())
        dc_text = StringIO(self.discharge_text.text())
        generate_excel(adm_text, dc_text, self._61_empty.text(),self._61_man.text(),self._61_woman.text(),self._62_empty.text(),self._62_man.text(),self._62_woman.text())
        QMessageBox.about(self,'작업 완료','당직표 파일이 생성되었습니다!\n엑셀이 실행됩니다.')
        os.startfile('dangjik.xlsx')
        self.close()

    def no_admission(self, e) :
            if e:
                self.admission_text.setEnabled(False)
            else :
                self.admission_text.setEnabled(True)

    def no_discharge(self, e) :
            if e:
                self.discharge_text.setEnabled(False)
            else :
                self.discharge_text.setEnabled(True)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())