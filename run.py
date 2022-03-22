import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from callyourname import Ui_Form_homework
import xlrd
import random


class Call(QtWidgets.QWidget, Ui_Form_homework):
    def __init__(self):
        super(Call, self).__init__()
        self.setupUi(self)
        self.init()
        self.setWindowTitle("python机器学习课程随机点名")


    def init(self):
        #打开名单文件
        self.pushButton_path.clicked.connect(self.openfile)
        #随机点名
        self.pushButton_call.clicked.connect(self.call)

    #打开图像文件
    def openfile(self):
        openfile_name, openfile_type = QFileDialog.getOpenFileName(self, '选择文件', '.', "All Files (*);;excel Files (*.xlsx)")
        self.lineEdit_path.setText(openfile_name)

    def call(self):
        filename = self.lineEdit_path.text()
        data = xlrd.open_workbook(filename)
        table = data.sheet_by_index(0)
        rows = table.nrows
        num = random.randint(1, int(rows - 1))
        list_fuck = table.row_values(num, start_colx=0, end_colx=None)

        #你的excel中姓名位于第几列请在这里修改
        name = list_fuck[0]

        self.lineEdit_name.setText(name)

        title = table.row_values(0, start_colx=0, end_colx=None)

        # 列宽自动分配
        self.tableWidget_call.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 行高自动分配
        self.tableWidget_call.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.tableWidget_call.setColumnCount(len(list_fuck))
        self.tableWidget_call.setRowCount(1)

        self.tableWidget_call.setHorizontalHeaderLabels(title)
        font1 = self.tableWidget_call.horizontalHeader().font()
        font1.setBold(True)
        self.tableWidget_call.horizontalHeader().setFont(font1)

        for i in range(0, len(list_fuck)):
            self.item = QTableWidgetItem(list_fuck[i])
            self.item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            self.tableWidget_call.setItem(0, i, self.item)

if __name__ == '__main__':
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QtWidgets.QApplication(sys.argv)
    call = Call()
    call.show()
    sys.exit(app.exec_())
