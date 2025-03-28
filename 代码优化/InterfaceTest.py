import configparser
import os
import json
import logging
import sys
import traceback

import openpyxl
import xlrd2
import csv

from PyQt5.QtGui import QKeySequence, QCursor, QFont
from PyQt5.QtWidgets import QShortcut, QComboBox, QAction, QMenu

from PyQt5 import QtWidgets, QtCore
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import QMessageBox, QApplication, QMainWindow, QFileDialog, QTableWidgetItem, QHeaderView, \
    QPushButton
from xlsxwriter import Workbook
from CanLinConfig import CanLinConfig
import tkinter as tk
from tkinter import messagebox
from ELFAnalysis import ELFAnalysis

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

# class InterfaceTest():

class InterfaceTest(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(InterfaceTest, self).__init__(parent)
        self.setupUi(self)


class InterfaceTest(QMainWindow,Ui_MainWindow):
    config = CanLinConfig()
    elfAnalysis = ELFAnalysis()
    # elfAnalysis=None
    InterfaceList=[]
    hardwaredic={}
    CANdic= {}
    LINdic={}
    interfacedic ={}
    logging.basicConfig(
            filename="log.log",
            filemode="w",
            datefmt="%Y-%m-%d %H:%M:%S %p",
            format="%(asctime)s - %(name)s - %(levelname)s - %(module)s: %(message)s",
            level=logging.DEBUG
        )

    def __init__(self, parent = None):
        super(InterfaceTest, self).__init__(parent)
        self.setupUi(self)
        self.pushButton_2.clicked.connect(self.exceltoTable)
        self.settable()
        # self.pushButton.clicked.connect(lambda :self.readMesseage(path1,path2,path3))
        self.pushButton3.clicked.connect(lambda :self.addrow(0))



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




    def exceltoTable(self):
        try:
            self.openfile = self.exceltoTable or os.getcwd()
            path = QFileDialog.getOpenFileName(self, "选择文件", '/', "xlsx Files(*.xlsx)")
            if path[0]:
                book = xlrd2.open_workbook(path[0])
                sheet = book.sheet_by_name('软件接口定义表')
                row = self.tableWidget.rowCount()
                for i in range(1, sheet.nrows):
                    values = sheet.row_values(i)
                    self.addrow(0)
                    for j in range(len(values)):
                        try:
                            text = str(int(values[j]))
                        except:
                            text = str(values[j])
                        self.tableWidget.setItem(row + i - 1, j, QTableWidgetItem(text))
        except:
            logging.error(traceback.format_exc() + "\n")


    def addrow(self, row):
        try:
            if row == 0:
                rowPosition = self.tableWidget.rowCount()
            else:
                rowPosition = row
            self.tableWidget.insertRow(rowPosition)
            deleteButton = QPushButton("删除".format(rowPosition))
            deleteButton.clicked.connect(self.delete_clicked)
            self.tableWidget.setCellWidget(rowPosition, 10, deleteButton)
        except:
            logging.error(traceback.format_exc() + "\n")

    def delete_clicked(self):
        button = self.sender()
        if button:
            row = self.tableWidget.indexAt(button.pos()).row()
            self.tableWidget.removeRow(row)
            self.tableWidget.verticalScrollBar().setSliderPosition(row)

    def settable(self):
        headnameList = ['接口类型', '接口名称（代码中的全局变量）', '信号描述（用例显示名称）', '信号归属模块（ARXML名称）', 'map地址', '信号方向', '   信号长度(bit)   ', '正向测试值   ', '关联信号',
                        '关联属性','操作']
        self.tableWidget.setColumnCount(len(headnameList))
        self.tableWidget.setRowCount(0)
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.setHorizontalHeaderLabels(headnameList)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        # self.tableWidget.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.tableWidget.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.tableWidget.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.tableWidget.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.tableWidget.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        self.tableWidget.horizontalHeader().setSectionResizeMode(8, QHeaderView.ResizeToContents)
        self.tableWidget.horizontalHeader().setSectionResizeMode(9, QHeaderView.ResizeToContents)
        self.tableWidget.horizontalHeader().setSectionResizeMode(10, QHeaderView.ResizeToContents)
    def show_success_dialog(self,state,message):
        # 创建一个简单的Tkinter窗口
        success_dialog = tk.Tk()
        # 隐藏主窗口
        success_dialog.withdraw()
        # 显示成功消息
        messagebox.showinfo(state, message)

    def readMesseage(self, path1):
        try:
            filename = QFileDialog.getSaveFileName(self, '导出测试用例', '', 'Excel File(*.xlsx)')
            if filename[0]:
                path3 = filename[0]
                for i in range(0, self.tableWidget.rowCount()):
                    datalist = []
                    for j in range(0, self.tableWidget.columnCount() - 1):
                        data = self.tableWidget.item(i, j).text()
                        datalist.append(data)
                    self.InterfaceList.append(datalist)
                # self.elfAnalysis.loadPath(path2)
                if path1:
                    book1 = xlrd2.open_workbook(path1)
                    sheetbook1 = book1.sheet_by_name('软件接口定义表')
                    for i in range(1, sheetbook1.nrows):
                        values = sheetbook1.row_values(i)
                        # self.InterfaceList.append(values)
                        self.interfacedic[values[1]] = values[2]
                    sheetbook2 = book1.sheet_by_name('硬件配置表')
                    for i in range(2, 12):
                        values = sheetbook2.col_values(i)
                        CANlist = [str(values[2])] + values[4:]
                        CANlist = [item for item in CANlist if item and item.strip()]
                        self.hardwaredic[values[3]] = CANlist
                    sheetbook3 = book1.sheet_by_name('LIN用例解析矩阵')
                    dataList = []
                    for i in range(sheetbook3.nrows - 1, 0, -1):
                        values = sheetbook3.row_values(i)
                        if values[3] == '':
                            # 信号描述、起始位、信号长度
                            data = [values[8], values[11], values[13]]
                            dataList.append(data)
                        elif values[3] != [] and dataList != []:
                            self.LINdic[values[3]] = dataList
                            dataList = []
                    dataList = []
                    sheetbook4 = book1.sheet_by_name('CAN用例解析矩阵')
                    for i in range(sheetbook4.nrows - 1, 0, -1):
                        values = sheetbook4.row_values(i)
                        if values[3] == '':
                            # 信号描述、起始位、信号长度、周期
                            data = [values[8], values[11], values[13]]
                            dataList.append(data)
                        elif values[3] != [] and dataList != []:
                            self.CANdic[values[3] + '/' + str(values[5])] = dataList
                            dataList = []
                    self.genautotable(path3)
                    self.show_success_dialog('Success', '生成成功！')
        except:
            self.show_success_dialog('FAIL', '生成失败，查看log！')
            logging.error(traceback.format_exc() + "\n")

    def genautotable(self,path3):
        dataList = []
        for item in self.InterfaceList:
            channel = '00'
            type = ''
            CANID = ''
            for key, values in self.hardwaredic.items():
                for value in values:
                    if value in item[8]:
                        CANID = str(values[0])
                        channel = values.index(value)
                        type = key
            # 对首行的CANID处理
            if CANID == '700/701':
                CANID = '0x700'
            elif '.' in CANID:
                CANID = '0x' + CANID.split('.')[0]
            else:
                CANID = '0x000'
            var_name = item[1].strip().upper()  # 去除空格并转换为大写
            try:
                inaddress = self.elfAnalysis.getAddressWithName(var_name).split('x')[1]
            except KeyError:
                logging.error(f"Variable not found in ELF file: {var_name}")
                #inaddress = "00000000"  # 使用默认值或跳过该条目

            # inaddress = '00000000'
            if type == 'LIN':
                thirddata1 = [key for key, values in self.LINdic.items() for val in values if item[8] in val]
                thirdstartbit = [val[1] for key, values in self.LINdic.items() for val in values if item[8] in val]
                thirdlength = [val[2] for key, values in self.LINdic.items() for val in values if item[8] in val][0]
                startbit = str(hex(int(thirdstartbit[0]))).upper()[2:].zfill(4)
                shuxing = '01'
                cycle = '000A'
            elif type == 'CAN':
                thirddata1 = [key for key, values in self.CANdic.items() for val in values if item[8] in val]
                thirdstartbit = [val[1] for key, values in self.CANdic.items() for val in values if item[8] in val]
                thirdlength = [val[2] for key, values in self.CANdic.items() for val in values if item[8] in val][0]
                startbit = str(hex(int(thirdstartbit[0]))).upper()[2:].zfill(4)
                shuxing = '00'
                cycle = str(hex(int(thirddata1[0].split('/')[1]))).upper()[2:].zfill(4)
            else:
                thirddata1 = '00'
                thirdstartbit = '0'
                startbit = '0000'
                shuxing = '00'
                cycle = '0000'
                thirdlength = '0'
            if item[5] =='IN':
                preview = []
                if item[9]!='HWA':
                    # 找关联信号ID
                    length = str(hex(int(item[6]))).upper()[2:].zfill(4)
                    # print(thirddata1[0])
                    data1 = thirddata1[0].split('x')[1].split('/')[0].zfill(3)
                    text = str(channel).zfill(2) + shuxing + '4' + data1 + '0064' + cycle + startbit + length
                    text = ' '.join(text[i:i + 2] for i in range(0, len(text), 2)).upper()
                    firsttext=['event',item[8],text,'21','1',CANID,'','','--', 'Motorola']
                    secondtext = ['check', '单条测试用例处理状态', CANID[-1], '22', '100', '0x710', '16', '8', '--','Motorola']
                else:
                    if item[7]=='1':
                        data2='0001'
                    else:
                        data2 = '0000'
                    #暂时按半桥处理
                    shuxing ='01'
                    message1 = str(channel).zfill(2)+shuxing + data2 + 'FFFF0000'
                    message1 = ' '.join(message1[i:i + 2] for i in range(0, len(message1), 2)).upper()
                    firsttext =['event', item[8], message1, '20', '1', '0x705', '', '', '--', 'Motorola']
                    secondtext = ['check', '单条测试用例处理状态', CANID[-1], '21', '100', '0x710', '16', '8', '--','Motorola']
                dataList.append(firsttext)
                preview.append(firsttext)
                dataList.append(secondtext)
                preview.append(secondtext)
                if item[9]!='HWA':
                    string = self.config.getMessage(thirdstartbit[0],item[6],item[7],8,'')
                    string11= self.config.getMessage(thirdstartbit[0],item[6],'0',8,'')
                    thirdtext=['event',item[8],string,'100','1',thirddata1[0],'','','--', 'Motorola']
                    dataList.append(thirdtext)
                    threetext=['event',item[8],string11,'100','1',thirddata1[0],'','','--', 'Motorola']
                    preview.append(threetext)
                if item[0]=='UInt8' or item[0]=='Boolean':
                    num = '01'
                elif item[0]=='SInt8':
                    num='81'
                elif item[0]=='UInt16':
                    num='02'
                elif item[0]=='SInt16':
                    num='82'
                else:
                    num ='00'
                # inaddress='00000000'
                message = '2EF6E9' + inaddress+num+'00000000'
                message = ' '.join(message[i:i + 2] for i in range(0, len(message), 2)).upper()
                if item[9]!='HWA':
                    forthtext =['event','Rte_Read_'+item[2],message,'24','1','0x7F0','','','--', 'Motorola']
                    fifthtext = ['check', '单条测试用例处理状态', '240', '25', '100', '0x710', '16', '8', '--','Motorola']
                    sixthtext = ['event', 'Rte_Read_' + item[2], '22 F6 E9', '26', '1', '0x7F0', '', '', '--','Motorola']
                    seventhtext = ['check', '单条测试用例处理状态', '240', '27', '100', '0x710', '16', '8', '--','Motorola']
                else:
                    forthtext = ['event', 'Rte_Read_' + item[2], message, '22', '1', '0x7F0', '', '', '--', 'Motorola']
                    fifthtext = ['check', '单条测试用例处理状态', '240', '23', '100', '0x710', '16', '8', '--','Motorola']
                    sixthtext = ['event', 'Rte_Read_' + item[2], '22 F6 E9', '24', '1', '0x7F0', '', '', '--','Motorola']
                    seventhtext = ['check', '单条测试用例处理状态', '240', '25', '100', '0x710', '16', '8', '--','Motorola']
                dataList.append(forthtext)
                preview.append(forthtext)
                dataList.append(fifthtext)
                preview.append(fifthtext)
                dataList.append(sixthtext)
                preview.append(sixthtext)
                dataList.append(seventhtext)
                preview.append(seventhtext)
                zhengxiang = str(hex(int(item[7]))).upper()[2:].zfill(8)
                if zhengxiang!='00000000':
                    fanxiang = '00000000'
                else:
                    fanxiang = str(hex(int(item[7]))).upper()[2:].zfill(8)
                    if fanxiang== '00000000':
                        fanxiang = '00000001'
                message8 ='62F6E9'+inaddress+num+zhengxiang
                message8 = ' '.join(message8[i:i + 2] for i in range(0, len(message8), 2)).upper()
                message16= '62F6E9'+inaddress+num+fanxiang
                message16 = ' '.join(message16[i:i + 2] for i in range(0, len(message16), 2)).upper()
                if item[9]!='HWA':
                    eighthtext = ['routercheck', 'Rte_Read_' + item[2], message8, '28', '100', '0x7F1', '', '', '--',
                                  'Motorola']
                    eighttext= ['routercheck','Rte_Read_'+item[2],message16,'28','100','0x7F1','','','--','Motorola']
                else:
                    eighthtext = ['routercheck', 'Rte_Read_' + item[2], message8, '26', '100', '0x7F1', '', '', '--',
                                  'Motorola']
                    eighttext = ['routercheck', 'Rte_Read_' + item[2], message16, '26', '100', '0x7F1', '', '', '--',
                                 'Motorola']
                dataList.append(eighthtext)
                preview.append(eighttext)
                dataList=dataList+preview
                note = ['测试用例说明','由'+item[8]+'   测：Rte_Read_'+item[2]]
                dataList.append(note)
            elif item[5]=='OUT':
                preview=[]
                zhengxiang = str(hex(int(item[7]))).upper()[2:].zfill(8)
                if zhengxiang != '00000000':
                    fanxiang = '00000000'
                    caozuozhi = str(hex(int(item[7]))).upper()[2:]
                    fanxiangnum = '0'
                else:
                    fanxiang = str(hex(int(item[7]))).upper()[2:].zfill(8)
                    caozuozhi = '0'
                    fanxiangnum=str(hex(int(item[7]))).upper()[2:]
                if item[0]=='UInt8' or item[0]=='Boolean':
                    num = '11'
                elif item[0]=='SInt8':
                    num='91'
                elif item[0]=='UInt16':
                    num='12'
                elif item[0]=='SInt16':
                    num='92'
                else:
                    num ='00'
                text = '2EF6E9'+inaddress+num+zhengxiang
                text = ' '.join(text[i:i + 2] for i in range(0, len(text), 2)).upper()
                text1 = '2EF6E9'+inaddress+num+fanxiang
                text1 = ' '.join(text1[i:i + 2] for i in range(0, len(text1), 2)).upper()
                firsttext = ['event','Rte_Write_'+item[2],text,'20','1','0x7F0','','','--','Motorola']
                dataList.append(firsttext)
                firsttext1 = ['event','Rte_Write_'+item[2],text1,'20','1','0x7F0','','','--','Motorola']
                preview.append(firsttext1)
                secondtext= ['check','单条测试用例处理状态','240','21','100','0x710','16','8','--','Motorola']
                dataList.append(secondtext)
                preview.append(secondtext)
                if '开关输入保留' in item[8]:
                    CANID = '0x706'
                if item[9]!='HWA' and item[9]!='APP':
                    # print(thirdstartbit, thirdlength)
                    data=thirddata1[0].split('/')[0][2:].zfill(4)
                    message3 = str(channel).zfill(2) + shuxing + data + '0064' + cycle + self.config.getStartandLengthHex(thirdstartbit[0],thirdlength).replace(' ','')
                    message3 = ' '.join(message3[i:i + 2] for i in range(0, len(message3), 2)).upper()
                    thirdtext = ['event', item[8],message3,'22','1','0x701','','','--','Motorola']
                    forthtext = ['check','单条测试用例处理状态','1','100','100','0x710','16','8','--','Motorola']
                    dataList.append(thirdtext)
                    preview.append(thirdtext)
                    dataList.append(forthtext)
                    preview.append(forthtext)
                    fifthtext = ['check', item[8], caozuozhi, '23', '100', CANID, '', '', '--', 'Motorola']
                    dataList.append(fifthtext)
                    fivetext = ['check', item[8], fanxiangnum, '23', '100', CANID, '', '', '--', 'Motorola']
                    preview.append(fivetext)
                    note = ['测试用例说明', '由Rte_Write_' + item[2] + '   测：' + item[8]]
                elif item[9]=='HWA':
                    fifthtext = ['check', item[8], caozuozhi, '22', '100', CANID, '', '', '--', 'Motorola']
                    dataList.append(fifthtext)
                    fivetext = ['check', item[8], fanxiangnum, '22', '100', CANID, '', '', '--', 'Motorola']
                    preview.append(fivetext)
                    note = ['测试用例说明', '由Rte_Write_' + item[2] + '   测：' + item[8]]
                else:
                    if item[0] == 'UInt8' or item[0] == 'Boolean':
                        num = '01'
                    elif item[0] == 'SInt8':
                        num = '81'
                    elif item[0] == 'UInt16':
                        num = '02'
                    elif item[0] == 'SInt16':
                        num = '82'
                    else:
                        num = '00'
                    try:
                        name = [key for key, values in self.interfacedic.items() if item[8] == values][0]
                    except:
                        self.show_success_dialog('FAIL', 'APP属性的关联信号无法找到Rte接口！')
                        sys.exit(1)
                    # outaddress='00000000'
                    outaddress = self.elfAnalysis.getAddressWithName(name).split('x')[1]
                    data = '2EF6E9'+outaddress+num+fanxiang
                    data =' '.join(data[i:i + 2] for i in range(0, len(data), 2)).upper()
                    thirdtext = ['event','Rte_Read_'+item[8],data,'22','1','0x7F0', '', '', '--', 'Motorola']
                    dataList.append(thirdtext)
                    preview.append(thirdtext)
                    forthtext = ['check','单条测试用例处理状态','240','23','100','0x710','16','8','--','Motorola']
                    dataList.append(forthtext)
                    preview.append(forthtext)
                    fifthtext = ['event','Rte_Read_'+item[8],'22 F6 E9','24','1','0x7F0', '', '', '--', 'Motorola']
                    dataList.append(fifthtext)
                    preview.append(fifthtext)
                    sixthtext = ['check','单条测试用例处理状态','240','25','100','0x710','16','8','--','Motorola']
                    dataList.append(sixthtext)
                    preview.append(sixthtext)
                    message7='62F6E9'+outaddress+num+zhengxiang
                    message7 =' '.join(message7[i:i + 2] for i in range(0, len(message7), 2)).upper()
                    message14 ='62F6E9'+outaddress+num+fanxiang
                    message14=' '.join(message14[i:i + 2] for i in range(0, len(message14), 2)).upper()
                    seventhtext = ['routercheck', 'Rte_Read_' + item[8], message7, '26', '100', '0x7F1', '', '', '--',
                                  'Motorola']
                    seventext =['routercheck', 'Rte_Read_' + item[8], message14, '26', '100', '0x7F1', '', '', '--',
                                  'Motorola']
                    dataList.append(seventhtext)
                    preview.append(seventext)
                    note = ['测试用例说明', '由Rte_Write_' + item[2] + '   测：Rte_Read_' + item[8]]
                dataList=dataList+preview
                dataList.append(note)
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


if __name__ == '__main__':
    # path1 = r"C:\Users\pc\Downloads\自动化测试盒协议_BDM_KQC2_R.xlsx"
    # path2 = r"C:\Users\pc\Downloads\CYT4BF_M7_Master.elf"
    # path3= './接口测试.xlsx'
    app =QApplication(sys.argv)
    # 初始化
    mainwindow = QMainWindow()
    ui_components = Ui_MainWindow()
    ui_components.setupUi(mainwindow)
    aa = InterfaceTest()
    aa.show()
    sys.exit(app.exec_())
    # path3= './接口测试.xlsx'
    # aa = InterfaceTest()
    # aa.readMesseage(path1,path2,path3)



















