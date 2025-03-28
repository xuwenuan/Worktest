import configparser
import os
import re
import logging
import sys
import traceback
import xlrd2
import csv

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QCursor
from xlsxwriter import Workbook
from CanLinConfig import CanLinConfig
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QMessageBox, QApplication, QMainWindow, QFileDialog, QTableWidgetItem, QHeaderView, \
    QPushButton, QAction, QMenu, QComboBox
from xlsxwriter import Workbook
from CanLinConfig import CanLinConfig
import tkinter as tk
from tkinter import messagebox

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1300, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.gridLayout.addWidget(self.tableWidget, 0, 0, 1, 3)
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setObjectName("pushButton")
        self.gridLayout.addWidget(self.pushButton, 1, 1, 1, 1)
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setObjectName("pushButton_2")
        self.gridLayout.addWidget(self.pushButton_2, 1, 0, 1, 1)
        self.pushButton3 = QtWidgets.QPushButton(self.centralwidget)
        self.gridLayout.addWidget(self.pushButton3, 1, 2, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 23))
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(self.menubar)
        self.menu.setObjectName("menu")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "空调测试用例编写工具"))
        self.pushButton.setText(_translate("MainWindow", "导出文件"))
        self.pushButton_2.setText(_translate("MainWindow", "导入文件"))
        self.pushButton3.setText('添加行')
# class RoutingTest:

class SignalRoutingTest(QMainWindow, Ui_MainWindow):
    config = CanLinConfig()
    logging.basicConfig(
            filename="log.log",
            filemode="w",
            datefmt="%Y-%m-%d %H:%M:%S %p",
            format="%(asctime)s - %(name)s - %(levelname)s - %(module)s: %(message)s",
            level=logging.DEBUG
        )
    routingList = []
    CANList = []
    banqiaoList=[]
    def __init__(self, parent = None):
        super(SignalRoutingTest, self).__init__(parent)
        self.setupUi(self)
        self.pushButton_2.clicked.connect(self.exceltoTable)
        self.settable()
        # self.pushButton.clicked.connect(lambda :self.readMesseage(path2,path3,path4))
        self.pushButton3.clicked.connect(lambda: self.addrow(0))

    def exceltoTable(self):
        try:
            self.openfile = self.exceltoTable or os.getcwd()
            path = QFileDialog.getOpenFileName(self, "选择文件", '/', "csv Files(*.csv)")
            if path[0]:
                row = self.tableWidget.rowCount()
                with open(path[0], 'r') as file:
                    reader = csv.reader(file)
                    for i,val in enumerate(reader):
                        values = list(val)
                        if values!='':
                            self.addrow(0)
                        for j in range(len(values)):
                            try:
                                text = str(int(values[j]))
                            except:
                                text = str(values[j])
                            self.tableWidget.setItem(row + i - 1, j, QTableWidgetItem(text))
                self.delete_last_row()
        except:
            logging.error(traceback.format_exc() + "\n")

    def delete_last_row(self):
        row_count = self.tableWidget.rowCount()
        if row_count > 0:
            self.tableWidget.removeRow(row_count - 1)


    def addrow(self, row):
        try:
            if row == 0:
                rowPosition = self.tableWidget.rowCount()
            else:
                rowPosition = row
            self.tableWidget.insertRow(rowPosition)
            deleteButton = QPushButton("删除".format(rowPosition))
            deleteButton.clicked.connect(self.delete_clicked)
            self.tableWidget.setCellWidget(rowPosition, self.tableWidget.columnCount()-1, deleteButton)
        except:
            logging.error(traceback.format_exc() + "\n")

    def delete_clicked(self):
        button = self.sender()
        if button:
            row = self.tableWidget.indexAt(button.pos()).row()
            self.tableWidget.removeRow(row)
            self.tableWidget.verticalScrollBar().setSliderPosition(row)

    def settable(self):
        headnameList = ['   Name    ', 'Hyt', 'CANFD', 'ID', 'dlc', 'cycle', 'segment', 'startbit', 'length',
                        '预留','预留','预留','    Name    ','Hyt','CANFD','ID','dlc','cycle','segment','startbit',
                        'length','预留','操作']
        self.tableWidget.setColumnCount(len(headnameList))
        self.tableWidget.setRowCount(0)
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.setHorizontalHeaderLabels(headnameList)
        # self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.tableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tableWidget.horizontalHeader().setSectionResizeMode(12, QHeaderView.ResizeToContents)
        # self.tableWidget.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        # self.tableWidget.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        # self.tableWidget.horizontalHeader().setSectionResizeMode(9, QHeaderView.ResizeToContents)
        # self.tableWidget.horizontalHeader().setSectionResizeMode(11, QHeaderView.ResizeToContents)
        # self.tableWidget.horizontalHeader().setSectionResizeMode(18, QHeaderView.ResizeToContents)

    def show_success_dialog(self,state,message):
        # 创建一个简单的Tkinter窗口
        success_dialog = tk.Tk()
        success_dialog.withdraw()  # 隐藏主窗口

        # 显示成功消息
        messagebox.showinfo(state, message)



    #添加的复制粘贴模块
    def contextMenuEvent(self, event):
        try:
        # 创建右键菜单
            menu = QMenu(self)
            row = self.tableWidget.rowAt(event.y())

        # 添加菜单项：插入步骤、插入备注
            insert_row_action = menu.addAction('插入步骤')
            note_row_action = menu.addAction('插入备注')

        # 添加复制和粘贴功能
            copy_action = QAction("复制 (Ctrl+C)", self)
            paste_action = QAction("粘贴 (Ctrl+V)", self)

            copy_action.triggered.connect(self.copy_selected_cells)
            paste_action.triggered.connect(self.paste_cells)

            menu.addAction(copy_action)
            menu.addAction(paste_action)

        # 绑定菜单项触发事件
            insert_row_action.triggered.connect(lambda: self.addrow1(row + 1))
            note_row_action.triggered.connect(lambda: self.addnote1(row + 1))

        # 显示菜单
            menu.exec_(QCursor.pos())

        except Exception as e:
            pass
            # logging.error(f"Error in contextMenuEvent: {traceback.format_exc()}")

    def copy_selected_cells(self):
        """复制选中的单元格内容到剪贴板"""
        selected_ranges = self.tableWidget.selectedRanges()
        if not selected_ranges:
            # logging.info("No cells selected for copying.")
            return

        clipboard_text = ""
        for selected_range in selected_ranges:
            top_row, bottom_row = selected_range.topRow(), selected_range.bottomRow()
            left_col, right_col = selected_range.leftColumn(), selected_range.rightColumn()

        # 构造复制内容
            for row in range(top_row, bottom_row + 1):
                row_data = []
                for col in range(left_col, right_col + 1):
                    item = self.tableWidget.item(row, col)
                    if item:
                        row_data.append(item.text())
                    else:
                        widget = self.tableWidget.cellWidget(row, col)
                        if isinstance(widget, QComboBox):
                            row_data.append(widget.currentText())
                        else:
                            row_data.append("")
                clipboard_text += "\t".join(row_data) + "\n"

    # 将内容复制到剪贴板
        clipboard = QApplication.clipboard()
        clipboard.setText(clipboard_text.strip())
        # logging.info(f"Copied text to clipboard: {clipboard_text.strip()}")

    def paste_cells(self):
        """将剪贴板内容粘贴到选中的单元格"""
        clipboard = QApplication.clipboard()
        clipboard_text = clipboard.text()
        if not clipboard_text:
            # logging.info("No text in clipboard to paste.")
            return

    # 解析剪贴板内容为二维数组
        rows = clipboard_text.split("\n")
        data = [row.split("\t") for row in rows]

    # 获取目标起始位置
        selected_ranges = self.tableWidget.selectedRanges()
        if not selected_ranges:
            # logging.info("No cells selected for pasting.")
            return
        start_row = selected_ranges[0].topRow()
        start_col = selected_ranges[0].leftColumn()

    # 填充表格
        for i, row_data in enumerate(data):
            for j, cell_value in enumerate(row_data):
                target_row = start_row + i
                target_col = start_col + j

            # 动态扩展表格
                if target_row >= self.tableWidget.rowCount():
                    self.tableWidget.insertRow(target_row)
                if target_col >= self.tableWidget.columnCount():
                    self.tableWidget.insertColumn(target_col)

            # 设置单元格内容
                widget = self.tableWidget.cellWidget(target_row, target_col)
                if isinstance(widget, QComboBox):
                    widget.setCurrentText(cell_value)
                    # logging.info(f"Pasted text '{cell_value}' to QComboBox at ({target_row}, {target_col})")
                else:
                    item = QTableWidgetItem(cell_value)
                    self.tableWidget.setItem(target_row, target_col, item)
                    # logging.info(f"Pasted text '{cell_value}' to QTableWidgetItem at ({target_row}, {target_col})")






    def readMesseage(self,path2,path3,path4):
        try:
            filename = QFileDialog.getSaveFileName(self, '导出测试用例', '', 'Excel File(*.xlsx)')
            if filename[0]:
                path3 = filename[0]
                for i in range(0, self.tableWidget.rowCount()):
                    datalist = []
                    for j in range(0, self.tableWidget.columnCount() - 1):
                        data = self.tableWidget.item(i, j).text()
                        datalist.append(data)
                    self.routingList.append(datalist)
                book2 = xlrd2.open_workbook(path2)
                sheetbook2 = book2.sheet_by_name('硬件配置表')
                self.LINList=sheetbook2.col_values(3)[3:]
                self.CANList = sheetbook2.col_values(2)[3:]
                self.banqiaoList = sheetbook2.col_values(8)[3:]
                self.genautotable(path3,path4)
                self.show_success_dialog('Success', '生成成功！')
        except:
                self.show_success_dialog('FAIL', '生成失败，查看log！')
                logging.error(traceback.format_exc() + "\n")

    def genautotable(self,path3,path4):
        # self.routingList = self.routingList[1:]
        dataList=[]
        dataListhard = []
        for item in self.routingList:
            if item[1]!='' and item[0]:
                preview = []
                length = '1' * (int(item[8]))
                num = int(length, 2)
                if int(item[2]) == 2:
                    try:
                        string = self.config.getMessage(item[7], item[8], num, int(item[4]), 'intel')
                    except:
                        logging.info('\n此条数据可能存在异常：\nname:'+item[0]+'     startbit:'+item[7]+'     length:'+item[8]+'       value:'+str(num)+'     dlc:'+item[4]+'      encoding_order:intel\n')
                        # self.show_success_dialog('WARNING', '此条数据可能存在异常：\nstartbit:'+item[7]+'     length:'+item[8]+'       value:'+str(num)+'     dlc:'+item[4]+'      encoding_order:intel\n')
                    try:
                        data1 = '0' + str(self.LINList.index(item[1])) + '01'
                    except:
                        data1='0000'
                    ID = '4'+item[3].replace('0x', '').replace('0X', '').zfill(3)
                elif int(item[2]) == 1 or int(item[2]==0):
                    try:
                        string = self.config.getMessage(item[7], item[8], num, int(item[4]), 'Motorola')
                    except:
                        logging.info('\n此条数据可能存在异常：\nname:'+item[0]+'     startbit:'+item[7]+'     length:'+item[8]+'      value:'+str(num)+'      dlc:'+item[4]+'     encoding_order:Motorola\n')
                    # print('Motorola',item[0]+'_CAN')
                    try:
                        data1 = '0' + str(self.CANList.index(item[1]+'_CAN')) + '00'
                    except:
                        data1='0000'
                    ID = '8' + item[3].replace('0x', '').replace('0X', '').zfill(3)
                else:
                    try:
                        string = self.config.getMessage(item[7], item[8], num, int(item[4]), 'Motorola')
                    except:
                        logging.info('\n此条数据可能存在异常：\nname:'+item[0]+'     startbit:'+item[7]+'     length:'+item[8]+'      value:'+str(num)+'      dlc:'+item[4]+'     encoding_order:Motorola\n')
                    try:
                        data1 = '0' + str(self.CANList.index(item[1] + '_CAN')) + '00'
                    except:
                        data1 = '0000'
                    ID = item[3].replace('0x', '').replace('0X', '').zfill(4)
                if int(item[5]) < 20:
                    # data3 = '0014'
                    cycle = '20'
                else:
                    # data3 = str(hex(int(item[5]))).upper()[2:].zfill(4)
                    cycle = item[5]
                if int(item[2])==2:
                    name = item[1] + "_" + item[0]
                    jiange ='200'
                    newcycle='0014'
                    # type='Intel'
                else:
                    if item[0] != '/':
                        name = item[1] + '_CAN_' + item[3] + "_" + item[0]
                    else:
                        name = item[1] + '_CAN_' + item[3] + "_"
                    newcycle = str(hex(int(cycle) * 5)).upper()[2:].zfill(4)
                    jiange=int(cycle)*5
                    # type='Motorola'
                data2 = ID + '0014'
                data3= str(hex(int(item[4]))).upper()[2:].zfill(4)
                data4 = str(hex(int(item[7]))).upper()[2:].zfill(4)
                data5 = str(hex(int(item[8]))).upper()[2:].zfill(4)
                message1 = data1+data2+data3+data4+data5
                Mac = ' '.join(message1[i:i + 2] for i in range(0, len(message1), 2)).upper()
                data = ['event', name, Mac, '20', '3', '0x700', '', '', '--', 'Motorola']
                dataList.append(data)
                preview.append(data)
                seconditem = ['check', '单条测试用例处理状态', '0', '20', '100', '0x710', '16', '8', '--', 'Motorola']
                dataList.append(seconditem)
                preview.append(seconditem)
                # print(num)
                t = '00' * (int(item[4]) + 1)
                string1 = ' '.join(t[i:i + 2] for i in range(2, len(t), 2)).upper()
                thirditem = ['event', name, string, '20', '3', item[3], '', '', '--', 'Motorola']
                anthirditem = ['event', name, string1, '20', '3', item[3], '', '', '--', 'Motorola']
                dataList.append(thirditem)
                preview.append(anthirditem)
                if int(item[14]) == 2:
                    name = item[13] + "_" + item[12]
                    try:
                        channel = '0' + str(self.LINList.index(item[13])) + '01'
                    except:
                        channel='0000'
                    ID = '4'+item[15].replace('0x', '').replace('0X', '').zfill(3)
                elif int(item[14]) == 1:
                    if item[12] != '/':
                        name = item[13] + '_CAN_' + item[15] + "_" + item[12]
                    else:
                        name = item[13] + '_CAN_' + item[15] + "_"
                    try:
                        channel = '0' + str(self.CANList.index(item[13]+'_CAN')) + '00'
                    except:
                        channel='0000'
                    ID = '8' + item[15].replace('0x', '').replace('0X', '').zfill(3)
                else:
                    if item[12] != '/':
                        name = item[13] + '_CAN_' + item[15] + "_" + item[12]
                    else:
                        name = item[13] + '_CAN_' + item[15] + "_"
                    try:
                        channel = '0' + str(self.CANList.index(item[13]+'_CAN')) + '00'
                    except:
                        channel='0000'
                    ID = item[15].replace('0x', '').replace('0X', '').zfill(4)
                if int(item[17]) < 20:
                    data3 = '0014'
                    cycle2 = '20'
                else:
                    data3 = str(hex(int(item[17]))).upper()[2:].zfill(4)
                    cycle2 = item[17]
                if int(item[14])==2:
                    jiange = '100'
                    data2= '0014'
                else:
                    data2=str(hex(int(cycle2) * 5)).upper()[2:].zfill(4)
                    jiange=int(cycle2)
                data4 = str(hex(int(item[19]))).upper()[2:].zfill(4)
                data5 = str(hex(int(item[20]))).upper()[2:].zfill(4)
                message2 = channel+ID+data2+data3+data4+data5
                Mac2 = ' '.join(message2[i:i + 2] for i in range(0, len(message2), 2)).upper()
                forthitem = ['event', name, Mac2, '20', '3', '0x701', '', '', '--', 'Motorola']
                dataList.append(forthitem)
                preview.append(forthitem)
                fifthitem = ['check', '单条测试用例处理状态', '1', jiange, '100', '0x710', '16', '8', '--', 'Motorola']
                dataList.append(fifthitem)
                preview.append(fifthitem)
                relength = '1' * (int(item[20]))
                renum = int(relength, 2)
                sixth = ['check', name, renum, '20', '100', item[15], item[19], item[20], '--', 'Motorola']
                ansixth = ['check', name, '0', '20', '100', item[15], item[19], item[20], '--', 'Motorola']
                dataList.append(sixth)
                preview.append(ansixth)
                dataList = dataList+preview
                text=item[0]+' '+item[1]+' '+item[3]+'      '+item[12]+' '+item[13]+' '+item[15]
                note = ['测试用例说明', text, '12', '', '', '', '', '', '', '']
                dataList.append(note)
            elif item[1]=='' and item[0]!='':
                name = item[0]
                data1 = "0" + str(self.banqiaoList.index(item[0])) + '01'
                if item[9] == '1':
                    data2 = '0001'
                    check = '0'
                else:
                    data2 = '0000'
                    check='1'
                data3 = 'FFFF0000'
                message1 = data1 + data2 + data3
                Mac = ' '.join(message1[i:i + 2] for i in range(0, len(message1), 2)).upper()
                data = ['event', name, Mac, '100', '1', '0x705', '', '', '--', 'Motorola']
                dataListhard.append(data)
                seconditem = ['check', '单条测试用例处理状态', '5', '100', '100', '0x710', '16', '8', '--', 'Motorola']
                dataListhard.append(seconditem)
                if int(item[14]) == 2:
                    name = item[13] + "_" + item[12]
                    try:
                        channel = '0' + str(self.LINList.index(item[13])) + '01'
                    except:
                        channel = '0000'
                    ID = '4' + item[15].replace('0x', '').replace('0X', '').zfill(3)
                elif int(item[14]) == 1:
                    if item[12] != '/':
                        name = item[13] + '_CAN_' + item[15] + "_" + item[12]
                    else:
                        name = item[13] + '_CAN_' + item[15] + "_"
                    try:
                        channel = '0' + str(self.CANList.index(item[13] + '_CAN')) + '00'
                    except:
                        channel = '0000'
                    ID = '8' + item[15].replace('0x', '').replace('0X', '').zfill(3)
                else:
                    if item[12] != '/':
                        name = item[13] + '_CAN_' + item[15] + "_" + item[12]
                    else:
                        name = item[13] + '_CAN_' + item[15] + "_"
                    try:
                        channel = '0' + str(self.CANList.index(item[13] + '_CAN')) + '00'
                    except:
                        channel = '0000'
                    ID = item[15].replace('0x', '').replace('0X', '').zfill(4)
                if int(item[17]) < 20:
                    data3 = '0014'
                    cycle2 = '20'
                else:
                    data3 = str(hex(int(item[17]))).upper()[2:].zfill(4)
                    cycle2 = item[17]
                if int(item[14]) == 2:
                    jiange = '100'
                    data2 = '0014'
                else:
                    data2 = str(hex(int(cycle2) * 5)).upper()[2:].zfill(4)
                    jiange = int(cycle2)
                data4 = str(hex(int(item[19]))).upper()[2:].zfill(4)
                data5 = str(hex(int(item[20]))).upper()[2:].zfill(4)
                message2 = channel + ID + data2 + data3 + data4 + data5
                Mac2 = ' '.join(message2[i:i + 2] for i in range(0, len(message2), 2)).upper()
                forthitem = ['event', name, Mac2, '20', '1', '0x701', '', '', '--', 'Motorola']
                dataListhard.append(forthitem)
                fifthitem = ['check', '单条测试用例处理状态', '1', jiange, '100', '0x710', '16', '8', '--', 'Motorola']
                dataListhard.append(fifthitem)
                ansixth = ['check', name, check, '20', '100', item[15], item[19], item[20], '--', 'Motorola']
                dataListhard.append(ansixth)
                data1 = "0" + str(self.banqiaoList.index(item[0])) + '01'
                if item[9] == '0':
                    data2 = '0001'
                    check='0'
                else:
                    data2 = '0000'
                    check='1'
                data3 = 'FFFF0000'
                message1 = data1 + data2 + data3
                Mac = ' '.join(message1[i:i + 2] for i in range(0, len(message1), 2)).upper()
                sevenitem = ['event', item[0], Mac, '100', '1', '0x705', '', '', '--', 'Motorola']
                sixth = ['check', name, check, '20', '100', item[15], item[19], item[20], '--', 'Motorola']
                dataListhard.append(sevenitem)
                dataListhard.append(seconditem)
                dataListhard.append(forthitem)
                dataListhard.append(fifthitem)
                dataListhard.append(sixth)
                text = item[0] + ' ' + item[1] + ' ' + item[3] + '      ' + item[12] + ' ' + item[13] + ' ' + item[15]
                note = ['测试用例说明', text, '10', '', '', '', '', '', '', '']
                dataListhard.append(note)
        # print(dataList)
        if dataList!=[]:
            headnameList = ['序号', '操作类型', '操作名称', '操作值', '间隔（ms）', 'Cycle（ms）', 'canID', 'Start', 'Length',
                            'flag', 'format']
            workbook = Workbook(path3)
            sheet1 = workbook.add_worksheet('CAN')
            sheet1.set_column('C:D', 50)
            sheet1.set_column('B:B', 15)
            sheet1.set_column('E:K', 10)
            font = workbook.add_format({'font_name': '等线', 'font_size': 12, 'align': 'center'})
            for k in range(len(headnameList)):
                sheet1.write(0, k, headnameList[k], font)
            row = 1
            for l in range(len(dataList)):
                sheet1.write(row, 0, str(row), font)
                for j in range(len(dataList[l])):
                    sheet1.write(row, j+1, dataList[l][j], font)
                row = row+1
            workbook.close()
        if dataListhard!=[]:
            headnameList = ['序号', '操作类型', '操作名称', '操作值', '间隔（ms）', 'Cycle（ms）', 'canID', 'Start',
                            'Length',
                            'flag', 'format']
            workbook = Workbook(path4)
            sheet1 = workbook.add_worksheet('CAN')
            sheet1.set_column('C:D', 50)
            sheet1.set_column('B:B', 15)
            sheet1.set_column('E:K', 10)
            font = workbook.add_format({'font_name': '等线', 'font_size': 12, 'align': 'center'})
            for k in range(len(headnameList)):
                sheet1.write(0, k, headnameList[k], font)
            row = 1
            for l in range(len(dataListhard)):
                sheet1.write(row, 0, str(row), font)
                for j in range(len(dataListhard[l])):
                    sheet1.write(row, j + 1, dataListhard[l][j], font)
                row = row + 1
            workbook.close()


if __name__ == '__main__':
    # path1= r"C:\Users\pc\Downloads\SinaRoutingTable.CSV"
    # path2 = r"C:\Users\pc\Downloads\自动化测试盒协议_BDM_KQC2_R.xlsx"
    # path3 = './信号路由自动测试.xlsx'
    # path4='./硬件测试.xlsx'
    app = QApplication(sys.argv)
    # 初始化
    mainwindow = QMainWindow()
    ui_components = Ui_MainWindow()
    ui_components.setupUi(mainwindow)
    aa = SignalRoutingTest()
    aa.show()
    sys.exit(app.exec_())
    # aa = RoutingTest()
    # aa.readMesseage()



















