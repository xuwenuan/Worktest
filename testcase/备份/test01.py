"""
#!/usr/bin/env python3.9
#-*- coding:utf-8 -*-
@Project:UItest
@File:test01.py
@Author:XU AO
@Time:2025/4/17 16:49
"""
import logging
import sys
import os
import traceback

import xlrd2
from PyQt5.QtWidgets import QMainWindow, QApplication, QHeaderView, QPushButton, QMessageBox, QFileDialog, \
    QTableWidgetItem

from Models.InterfaceTest import InterfaceTest
from UI2 import Ui_MainWindow



class UiMain(QMainWindow,Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        # 获取 stackedWidget 实例并检查其页面数量
        self.total_pages = self.stackedWidget.count()  # 检查 stackedWidget 的页面数量
        # 设置初始页面索引
        self.initial_page_no = 0
        self.stackedWidget.setCurrentIndex(self.initial_page_no)

        # 绑定侧边栏按钮信号
        self.pushButton_10.clicked.connect(self.click_pushButton_10)
        self.pushButton.clicked.connect(self.click_pushButton_1)
        self.pushButton_2.clicked.connect(self.click_pushButton_2)
        self.pushButton_3.clicked.connect(self.click_pushButton_3)
        self.pushButton_4.clicked.connect(self.click_pushButton_4)

        #接口测试
        self.Interface = InterfaceTest()
        self.interfacetest_settable()
        self.pushButton_17.clicked.connect(self.exceltoTable)
        self.pushButton_18.clicked.connect(self.add_insterface_row)
        self.pushButton_19.clicked.connect(self.Interface.clear_table)
        self.pushButton_20.clicked.connect(self.Interface.export_to_excel)
        self.pushButton_21.clicked.connect(lambda : self.Interface.readMesseage(self.filename, self.elfpath))

    #添加行
    def add_insterface_row(self):
        try:
            current_row_count = self.tableWidget_2.rowCount()
            self.tableWidget_2.insertRow(current_row_count)
            deleteButton = QPushButton("删除")
            deleteButton.clicked.connect(self.delete_clicked)
            if self.stackedWidget.currentIndex() == 2:
                self.tableWidget_2.setCellWidget(current_row_count, 10, deleteButton)
            elif self.stackedWidget.currentIndex() == 3:
                self.tableWidget_2.setCellWidget(current_row_count, 18, deleteButton)
            elif self.stackedWidget.currentIndex() == 4:
                self.tableWidget_2.setCellWidget(current_row_count, self.tableWidget_5.columnCount()-1, deleteButton)
        except Exception as e:
            logging.error(traceback.format_exc())
            QMessageBox.critical(self, "错误", f"添加行时发生错误: {str(e)}")#可以让详细错误以弹窗的形式弹出

    #删除行
    def delete_clicked(self):
        button = self.sender()
        if button:
            row = self.tableWidget_2.indexAt(button.pos()).row()
            self.tableWidget_2.removeRow(row)
            self.tableWidget_2.verticalScrollBar().setSliderPosition(row)

    def exceltoTable(self):
        try:
            self.openfile = self.exceltoTable or os.getcwd()
            path = QFileDialog.getOpenFileName(self, "选择文件", '/', "xlsx Files(*.xlsx)")
            sheet_name = '软件接口定义表'
            if not path[0]:
                QMessageBox.information(self, "温馨提示", "未选择文件！")
                return
            if path[0]:
                book = xlrd2.open_workbook(path[0])
                if sheet_name not in book.sheet_names():
                    QMessageBox.information(self, "温馨提示", "导入文件不正确！")
                    return
                sheet = book.sheet_by_name('软件接口定义表')
                row = self.tableWidget_2.rowCount()
                for i in range(1, sheet.nrows):
                    values = sheet.row_values(i)
                    self.add_insterface_row()
                    for i in range(1, sheet.nrows):
                        values = sheet.row_values(i)
                        self.add_insterface_row()
                        for j in range(len(values)):
                            try:
                                text = str(int(values[j]))
                            except:
                                text = str(values[j])
                                QMessageBox.critical(self, "错误", f"添加行时发生错误: {traceback.format_exc()}")
                            self.tableWidget_2.setItem(row + i - 1, j, QTableWidgetItem(text))
        except:
            logging.error(traceback.format_exc() + "\n")
            QMessageBox.critical(self, "错误", f"添加行时发生错误: {traceback.format_exc()}")


    # 以下为按钮点击事件的实现
    def click_pushButton_10 (self):
        self.stackedWidget.setCurrentIndex(0)  # 设置 page_0 为当前页面  配置页面
    def click_pushButton_1(self):
        self.stackedWidget.setCurrentIndex(1)  # 设置 page_1 为当前页面  通用测试用例编写
    def click_pushButton_2(self):
        self.stackedWidget.setCurrentIndex(2)  # 设置 page_2 为当前页面  接口测试用例编写
    def click_pushButton_3(self):
        self.stackedWidget.setCurrentIndex(3)  # 设置 page_3 为当前页面  id信号路由测试用例编写
    def click_pushButton_4(self):
        self.stackedWidget.setCurrentIndex(4)  # 设置 page_4 为当前页面  信号路由测试用例编写

    #接口测试表格初始化
    def interfacetest_settable(self):
        headnameList = ['接口类型', '接口名称（代码中的全局变量）', '信号描述（用例显示名称）', '信号归属模块（ARXML名称）',
                        'map地址', '信号方向', '   信号长度(bit)   ', '正向测试值   ', '关联信号',
                        '关联属性', '操作']
        self.tableWidget_2.horizontalHeader().setVisible(True)
        self.tableWidget_2.setColumnCount(len(headnameList))
        self.tableWidget_2.setRowCount(0)
        self.tableWidget_2.verticalHeader().setVisible(False)
        self.tableWidget_2.setHorizontalHeaderLabels(headnameList)
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(8, QHeaderView.ResizeToContents)
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(9, QHeaderView.ResizeToContents)
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(10, QHeaderView.ResizeToContents)




if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = UiMain()
    win.show()
    sys.exit(app.exec_())
