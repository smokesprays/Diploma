from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
import sys


class CheckList(QMainWindow):

    def __init__(self):
        super(CheckList, self).__init__()
        self.listWidget = QListWidget()
        self.resize(720, 480)
        '''был установлен listWidget'''

        self.listWidget.addItem("Конфигурация Test")
        self.listWidget.addItem("Конфигурация 2")
        self.listWidget.addItem("Конфигурация 3")
        self.listWidget.addItem("Конфигурация 4")

        self.listWidget.itemClicked.connect(self.Clicked)  # connect itemClicked to Clicked method

        self.setCentralWidget(self.listWidget)

    def Clicked(self, item):
        QMessageBox.information(self, "ListWidget", "You clicked: " + item.text())

    def keyPressEvent(self, e):

        if e.key() == Qt.Key_Escape:
            self.listWidget.hide()
            self.setWindowTitle("Drag and Drop")
            self.resize(100, 100)
            self.setAcceptDrops(True)
            self.show()
            self.listWidget.hide()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        for f in files:
            print(f)


def check():
    app = QApplication(sys.argv)
    w = CheckList()
    w.show()
    sys.exit(app.exec_())
