"""
#!/usr/bin/env python3.9
#-*- coding:utf-8 -*-
@Project:testcase
@File:Testcase.py
@Author:XU AO
@Time:2025/4/22 08:51
"""
import logging
import re
import sys
import os
import threading
import time
import traceback
import icons_rc

import xlrd2
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QUrl
from PyQt5.QtGui import QDesktopServices, QCursor, QBrush, QColor, QIcon, QKeySequence
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox, QFileDialog, QProgressDialog, QTableWidgetItem, \
    QComboBox, QAction, QMenu, qApp, QPushButton, QLineEdit, QWidget, QHBoxLayout

from Function.CanLinConfig import CanLinConfig
from Function.ELFAnalysis import ELFAnalysis
from Function.GetFileData import GetFileData

from Models.GeneralTest import GeneralTest
from Models.InterFace import InterFace
from Models.RoutingTest import RoutingTest
from Models.SignaRoutingTest import SignalRoutingTest
from TestcaseUI import Ui_MainWindow


class UndoRedoStack:
    def __init__(self, max_operations=6):
        self.undo_stack = []
        self.redo_stack = []
        self.max_operations = max_operations

    def push(self, operation):
        self.undo_stack.append(operation)
        self.redo_stack.clear()
        if len(self.undo_stack) > self.max_operations:
            self.undo_stack.pop(0)

    def undo(self):
        if self.can_undo():
            operation = self.undo_stack.pop()
            self.redo_stack.append(operation)
            return self.undo_stack[-1]  # 不再额外检查 can_undo
        return None

    def redo(self):
        if self.can_redo():
            operation = self.redo_stack.pop()
            self.undo_stack.append(operation)
            return operation
        return None

    def can_undo(self):
        return len(self.undo_stack) > 1

    def can_redo(self):
        return len(self.redo_stack) > 0


class NewWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__()
        self.parent = parent
        self.setWindowTitle("搜索")


        # 初始化搜索框和按钮
        self.lineEdit_1 = QLineEdit()
        self.pushButton_1 = QPushButton("搜索")
        self.pushButton_1.clicked.connect(self.on_search_button_clicked)

        # 创建布局并添加控件
        layout = QHBoxLayout()
        layout.addWidget(self.lineEdit_1)
        layout.addWidget(self.pushButton_1)

        # 设置中心窗口
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setLayout(layout)

        self.show()

    def on_search_button_clicked(self):
        keyword = self.lineEdit_1.text().strip()
        if not keyword:
            QMessageBox.warning(None, "警告", "请输入搜索关键字")
            return
        # 调用父窗口的搜索功能
        if hasattr(self.parent, "search_in_table"):
            try:
                self.parent.search_in_table(keyword)
            except Exception as e:
                logging.error(f"搜索过程中发生错误: {str(e)}")
                QMessageBox.critical(None, "错误", f"搜索过程中发生错误: {str(e)}")
        else:
            QMessageBox.warning(None, "警告", "当前窗口未绑定搜索功能")

class UiMain(QMainWindow, Ui_MainWindow):
    diclist, lindic = [], []
    diclist1, lindic1 = [], []
    datalist, linlist = [], []
    signalD, linD = [], []
    signalD1, linD1 = [], []
    CAN = []
    LIN = []
    openclose = []
    temp = []
    CV = []
    PWM = []
    comboBoxList3 = []
    comboBoxList1 = ['前提', '触发', '期望', '配置']
    interfaceList = []
    interfaceDic = []
    structDict = []
    RTlist = []
    ID = []
    banqiao = []
    Canname = ''
    Linname = ''
    filename = ''
    resolutionnum = ''
    num = 0
    offset = 0
    resolution = 0
    chapterList = []
    chapterLevel1 = []
    chapterLevel2 = []
    chapterLevel3 = []
    chapterLevel4 = []
    config = CanLinConfig()
    getfile = GetFileData()
    elfAnalysis = ELFAnalysis()
    first_open_file_call = True

    logging.basicConfig(
        filename="log.log",
        filemode="w",
        datefmt="%Y-%m-%d %H:%M:%S %p",
        format="%(asctime)s - %(name)s - %(levelname)s - %(module)s: %(message)s",
        level=40
    )

    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.merge_column_config = {
            1: 9,  # 页面 1 合并 9 列
            2: 10,  # 页面 2 合并 6 列
            3: 18,  # 页面 3 合并 8 列
            4: 22  # 页面 4 合并 7 列
        }

        self.undo_redo_stack = UndoRedoStack()
        self.undo_redo_stack.push(self.get_table_data())  # 保存初始状态
        self.record_enabled = True  # 控制是否允许记录修改
        self.tableWidget_2.cellChanged.connect(self.record_operation)
        self.tableWidget_3.cellChanged.connect(self.record_operation)
        self.tableWidget_4.cellChanged.connect(self.record_operation)
        self.tableWidget_5.cellChanged.connect(self.record_operation)

        # 创建撤销快捷键 Ctrl+Z
        undo_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Z"), self)
        undo_shortcut.activated.connect(self.undo)

        # 创建重做快捷键 Ctrl+Y
        redo_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Y"), self)
        redo_shortcut.activated.connect(self.redo)


        # 初始化 stackedWidget 页面
        self.total_pages = self.stackedWidget.count()
        self.initial_page_no = 0
        self.stackedWidget.setCurrentIndex(self.initial_page_no)

        # 侧边栏按钮绑定
        self.pushButton_10.clicked.connect(self.click_pushButton_10)
        self.pushButton.clicked.connect(self.click_pushButton_1)
        self.pushButton_2.clicked.connect(self.click_pushButton_2)
        self.pushButton_3.clicked.connect(self.click_pushButton_3)
        self.pushButton_4.clicked.connect(self.click_pushButton_4)

        #菜单栏按钮绑定
        self.action_ELF_OUT.triggered.connect(self.getELFAnalysis)
        self.action.triggered.connect(self.show_message1)
        self.action_2.triggered.connect(self.show_message2)
        self.action111.triggered.connect(self.open_file)
        self.action_4.triggered.connect(self.open_window)

        # 功能点检
        #初始化 GeneralTest,传入主界面中的 tableWidget_5
        self.GeneralTest = GeneralTest(self.tableWidget_5)
        self.GeneralTest.settable()
        self.pushButton_32.clicked.connect(self.GeneralTest.exceltotable)
        self.pushButton_33.clicked.connect(self.GeneralTest.add_rows)
        self.pushButton_34.clicked.connect(self.GeneralTest.clear_table)
        self.pushButton_35.clicked.connect(self.GeneralTest.export_to_excel)
        self.pushButton_36.clicked.connect(self.GeneralTest.genAutoExcel)
        self.pushButton_6.clicked.connect(self.GeneralTest.gen_message)
        self.action_3.triggered.connect(self.GeneralTest.open_file)

        # 接口测试
        # 初始化 InterFace，传入主界面中的 tableWidget_2
        self.Interface = InterFace(self.tableWidget_2, self.elfAnalysis)
        self.Interface.settable()
        self.pushButton_17.clicked.connect(self.Interface.excel_toTable)
        self.pushButton_18.clicked.connect(self.Interface.add_rows)
        self.pushButton_19.clicked.connect(self.Interface.clear_table)
        self.pushButton_20.clicked.connect(self.Interface.export_to_excel)
        self.pushButton_21.clicked.connect(lambda :self.Interface.read_Messeage(self.filename, self.elfpath))

        #ID路由测试
        #初始化RoutingTset,传入主界面中的 tableWidget_3
        self.RoutingTest = RoutingTest(self.tableWidget_3)
        self.RoutingTest.settable()
        self.pushButton_22.clicked.connect(self.RoutingTest.excel_toTable)
        self.pushButton_23.clicked.connect(self.RoutingTest.add_rows)
        self.pushButton_24.clicked.connect(self.RoutingTest.clear_table)
        self.pushButton_25.clicked.connect(self.RoutingTest.export_to_excel)
        self.pushButton_26.clicked.connect(lambda :self.RoutingTest.readMesseage(self.filename))

        #信号路由测试
        # 初始化SignalRoutingTest,传入主界面中的 tableWidget_4
        self.SignalRoutingTest = SignalRoutingTest(self.tableWidget_4)
        self.SignalRoutingTest.settable()
        self.pushButton_27.clicked.connect(self.SignalRoutingTest.excel_toTable)
        self.pushButton_28.clicked.connect(self.SignalRoutingTest.add_rows)
        self.pushButton_29.clicked.connect(self.SignalRoutingTest.clear_table)
        self.pushButton_30.clicked.connect(self.SignalRoutingTest.export_to_excel)
        self.pushButton_31.clicked.connect(lambda :self.SignalRoutingTest.readMesseage(self.filename, './信号路由自动测试.xlsx', './硬件输入的信号路由测试.xlsx'))


    # 切换不同的页面
    def click_pushButton_10(self):
        self.stackedWidget.setCurrentIndex(0)  # 设置 page_0 为当前页面  配置页面

    def click_pushButton_1(self):
        self.stackedWidget.setCurrentIndex(1)  # 设置 page_1 为当前页面  通用测试用例编写

    def click_pushButton_2(self):
        self.stackedWidget.setCurrentIndex(2)  # 设置 page_2 为当前页面  接口测试用例编写

    def click_pushButton_3(self):
        self.stackedWidget.setCurrentIndex(3)  # 设置 page_3 为当前页面  id信号路由测试用例编写

    def click_pushButton_4(self):
        self.stackedWidget.setCurrentIndex(4)  # 设置 page_4 为当前页面  信号路由测试用例编写

    def show_message1(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile("./参考文档.pdf"))

    def show_message2(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile("./使用说明.pdf"))

    def get_current_table_widget(self):
        """根据当前页面索引返回对应的表格控件"""
        current_index = self.stackedWidget.currentIndex()
        if current_index == 1:
            return self.tableWidget_5
        elif current_index == 2:
            return self.tableWidget_2
        elif current_index == 3:
            return self.tableWidget_3
        elif current_index == 4:
            return self.tableWidget_4
        else:
            return None


    # 保存表格所有数据到列表
    def get_table_data(self):
        """获取当前表格的所有数据"""
        table_widget = self.get_current_table_widget()
        if not table_widget:
            QMessageBox.warning(None, "警告", "当前页面没有对应的表格控件")
            return []

        data = []
        for row in range(table_widget.rowCount()):
            row_data = []
            for col in range(table_widget.columnCount()):
                item = table_widget.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        return data


    # 从列表恢复到表格
    def set_table_data(self, data):
        """将数据恢复到当前表格"""
        table_widget = self.get_current_table_widget()
        if not table_widget:
            QMessageBox.warning(None, "警告", "当前页面没有对应的表格控件")
            return

        table_widget.blockSignals(True)  # 暂停信号，防止死循环
        table_widget.setRowCount(len(data))
        table_widget.setColumnCount(len(data[0]) if data else 0)

        for row_idx, row_data in enumerate(data):
            for col_idx, cell_data in enumerate(row_data):
                item = QtWidgets.QTableWidgetItem(cell_data)
                table_widget.setItem(row_idx, col_idx, item)

            # 重新绑定删除按钮（假设删除按钮在最后一列）
            delete_button = QPushButton("删除")
            delete_button.clicked.connect(lambda checked, row=row_idx: self.delete_row(row))
            table_widget.setCellWidget(row_idx, len(row_data) - 1, delete_button)

        table_widget.blockSignals(False)

    def delete_row(self, row):
        """删除指定行的操作"""
        table_widget = self.get_current_table_widget()
        table_widget.removeRow(row)
        self.record_operation()  # 删除操作后记录状态

    # 有任何修改就记录
    def record_operation(self):
        if not self.record_enabled:
            return  # 如果当前不允许记录，直接返回

        data_snapshot = self.get_table_data()
        if not self.undo_redo_stack.undo_stack or data_snapshot != self.undo_redo_stack.undo_stack[-1]:
            self.undo_redo_stack.push(data_snapshot)

    # 执行撤销
    def undo(self):
        operation = self.undo_redo_stack.undo()
        if operation is not None:
            self.set_table_data(operation)

    # 执行重做
    def redo(self):
        operation = self.undo_redo_stack.redo()
        if operation is not None:
            self.set_table_data(operation)



    def open_window(self):
        # 打开新窗口并将主窗口作为父窗口传递
        self.window = NewWindow(parent=self)
        self.window.resize( 300, 100)

    #搜索函数
    def search_in_table(self, keyword, columns=None):
        try:
            table_widget = self.get_current_table_widget()
            # 检查表格控件是否有效
            if table_widget is None:
                QMessageBox.warning(None, "警告", "表格控件未初始化")
                return

            # 清除之前的高亮
            self.clear_highlights(table_widget)

            # 如果未指定列，则默认搜索所有列
            if columns is None:
                columns = range(table_widget.columnCount())

            # 高亮颜色
            highlight_brush = QBrush(QColor(255, 255, 0))  # 黄色高亮

            # 遍历表格中的每一行和指定列
            found = False
            for row in range(table_widget.rowCount()):
                for col in columns:
                    item = table_widget.item(row, col)
                    if item and keyword.lower() in item.text().lower():  # 忽略大小写
                        item.setBackground(highlight_brush)
                        found = True

            # 提示用户搜索结果
            if not found:
                QMessageBox.information(None, "搜索结果", f"未找到关键字：{keyword}")
            else:
                QMessageBox.information(None, "搜索结果", f"已高亮显示关键字：{keyword}")

        except Exception as e:
            logging.error(f"搜索过程中发生错误: {str(e)}")
            QMessageBox.critical(None, "错误", f"搜索过程中发生错误: {str(e)}")

    #高亮清除函数
    def clear_highlights(self, table_widget):
        """
        清除表格中所有单元格的高亮背景。
        参数：
            table_widget (QTableWidget): 需要清除高亮的表格控件。
        """
        default_brush = QBrush(Qt.transparent)  # 默认透明背景
        for row in range(table_widget.rowCount()):
            # 获取第一列的内容，判断是否为备注行
            first_item = table_widget.item(row, 0)
            if first_item and "备注" in first_item.text():
                continue  # 如果是备注行，则跳过

            for col in range(table_widget.columnCount()):
                item = table_widget.item(row, col)
                if item:
                    item.setBackground(default_brush)



    def contextMenuEvent(self, event: QtGui.QContextMenuEvent) -> None:
        cmenu = QMenu(self)  # 实例化Qmenu对象

        copy_action = cmenu.addAction("复制")
        paste_action = cmenu.addAction("粘贴 ")
        explain_action = cmenu.addAction("插入备注 ")
        insert_action = cmenu.addAction("插入新行 ")

        copy_action.triggered.connect(self.copy_selected_cells)
        paste_action.triggered.connect(self.paste_cells)
        explain_action.triggered.connect(self.explain_selected_cells)
        insert_action.triggered.connect(self.insert_selected_cells)
        # 显示菜单
        cmenu.exec_(QCursor.pos())

    def copy_selected_cells(self):
        """复制选中的单元格内容到剪贴板"""
        table_widget = self.get_current_table_widget()
        # 获取选中的范围
        selected_ranges = table_widget.selectedRanges()
        clipboard_text = ""
        for selected_range in selected_ranges:
            top_row, bottom_row = selected_range.topRow(), selected_range.bottomRow()
            left_col, right_col = selected_range.leftColumn(), selected_range.rightColumn()

            # 构造复制内容
            for row in range(top_row, bottom_row + 1):
                row_data = []
                for col in range(left_col, right_col + 1):
                    item = table_widget.item(row, col)
                    if item:
                        row_data.append(item.text())
                    else:
                        widget = table_widget.cellWidget(row, col)
                        if isinstance(widget, QComboBox):
                            row_data.append(widget.currentText())
                        else:
                            row_data.append("")
                clipboard_text += "\t".join(row_data) + "\n"

        # 将内容复制到剪贴板
        clipboard = QApplication.clipboard()
        clipboard.setText(clipboard_text.strip())


    def paste_cells(self):
        """将剪贴板内容粘贴到选中的单元格"""
        try:
            # 获取当前页面索引
            current_index = self.stackedWidget.currentIndex()
            # 根据页面索引选择对应的表格控件
            if current_index == 1:
                table_widget = self.tableWidget_5
                delete_slot = self.GeneralTest.delete_clicked  # 绑定通用测试用例的删除逻辑
            elif current_index == 2:
                table_widget = self.tableWidget_2
                delete_slot = self.Interface.delete_clicked  # 绑定接口测试用例的删除逻辑
            elif current_index == 3:
                table_widget = self.tableWidget_3
                delete_slot = self.RoutingTest.delete_clicked  # 绑定 ID 路由测试用例的删除逻辑
            elif current_index == 4:
                table_widget = self.tableWidget_4
                delete_slot = self.SignalRoutingTest.delete_clicked  # 绑定信号路由测试用例的删除逻辑
            else:
                return

            # 检查表格控件是否有效
            if table_widget is None:
                return

            # 获取目标起始位置
            selected_ranges = table_widget.selectedRanges()
            start_row = selected_ranges[0].topRow()
            start_col = selected_ranges[0].leftColumn()

            # 获取剪贴板内容
            clipboard = QApplication.clipboard()
            clipboard_text = clipboard.text()

            # 解析剪贴板内容为二维数组
            rows = clipboard_text.split("\n")
            data = [row.split("\t") for row in rows]

            # 计算粘贴数据的最大列数
            max_columns = max(len(row_data) for row_data in data)

            # 动态扩展表格列数
            current_columns = table_widget.columnCount()
            if max_columns > current_columns:
                table_widget.setColumnCount(max_columns)

            # 填充表格
            for i, row_data in enumerate(data):
                for j, cell_value in enumerate(row_data):
                    target_row = start_row + i
                    target_col = start_col + j

                    # 动态扩展表格行数
                    if target_row >= table_widget.rowCount():
                        table_widget.insertRow(target_row)

                    # 如果目标列超出当前列数，跳过该列
                    if target_col >= table_widget.columnCount():
                        continue

                    # 设置单元格内容
                    widget = table_widget.cellWidget(target_row, target_col)
                    if isinstance(widget, QComboBox):
                        widget.setCurrentText(cell_value)
                    else:
                        item = QTableWidgetItem(cell_value)
                        table_widget.setItem(target_row, target_col, item)

                # 在每一行的最后一个单元格添加删除按钮
                delete_button = QPushButton("删除")
                delete_button.clicked.connect(lambda checked, row=target_row: delete_slot())
                last_column = table_widget.columnCount() - 1
                table_widget.setCellWidget(target_row, last_column, delete_button)

        except Exception as e:
            logging.error(f"粘贴过程中发生错误: {str(e)}")
            msg_box = QMessageBox(QMessageBox.Critical, "粘贴错误", f"粘贴过程中发生错误: {str(e)}")
            msg_box.exec_()


    def explain_selected_cells(self):
        try:
            current_index = self.stackedWidget.currentIndex()
            if current_index == 1:
                table_widget = self.tableWidget_5
            elif current_index == 2:
                table_widget = self.tableWidget_2
            elif current_index == 3:
                table_widget = self.tableWidget_3
            elif current_index == 4:
                table_widget = self.tableWidget_4
            else:
                QMessageBox.warning(None, "警告", "当前页面没有对应的表格控件")
                return
            selected_ranges = table_widget.selectedRanges()
            if not selected_ranges:
                QMessageBox.warning(None, "警告", "未选中任何单元格")
                return

            selected_row = selected_ranges[0].bottomRow()
            insert_row = selected_row + 1

            # 动态获取合并列数
            merge_columns = self.merge_column_config.get(current_index, 9)  # 默认值为 9

            # 确保列数 >= 合并列数
            col_count = table_widget.columnCount()
            if col_count < merge_columns:
                table_widget.setColumnCount(merge_columns)

            # 插入新行
            table_widget.insertRow(insert_row)

            # 设置灰色背景和空白项
            # brush = QBrush(QColor(230, 230, 230))
            brush = QBrush(QColor(119, 136, 153))  # 使用柔和的深灰色
            # brush = QBrush(QColor(47, 79, 79))      # 或者更暗的深灰色
            # brush = QBrush(QColor(255, 255, 0))  # 设置背景颜色为黄色
            for i in range(col_count):
                item = QTableWidgetItem("" if i < col_count - 1 else None)
                if item:
                    item.setBackground(brush)
                    table_widget.setItem(insert_row, i, item)

            # 设置备注文字
            table_widget.setItem(insert_row, 0, QTableWidgetItem("备注：请填写备注内容"))
            table_widget.item(insert_row, 0).setBackground(brush)

            # 合并前 col_count - 1 列为一个备注区域
            table_widget.setSpan(insert_row, 0, 1, col_count - 1)

            # 添加删除按钮到最后一列
            delete_button = QPushButton("删除")
            delete_button.clicked.connect(self.delete_clicked)  # 修改为删除当前选中的行
            table_widget.setCellWidget(insert_row, col_count - 1, delete_button)

        except Exception as e:
            logging.error(f"添加备注行时发生错误: {str(e)}")
            QMessageBox.critical(None, "错误", f"添加备注行时发生错误: {str(e)}")

    def delete_clicked(self):
        try:
            current_index = self.stackedWidget.currentIndex()
            if current_index == 1:
                table_widget = self.tableWidget_5
            elif current_index == 2:
                table_widget = self.tableWidget_2
            elif current_index == 3:
                table_widget = self.tableWidget_3
            elif current_index == 4:
                table_widget = self.tableWidget_4
            else:
                QMessageBox.warning(None, "警告", "当前页面没有对应的表格控件")
                return
            current_row = table_widget.currentRow()
            table_widget.removeRow(current_row)  # 删除行
            table_widget.verticalScrollBar().setSliderPosition(current_row)  # 滚动条调整位置
        except Exception as e:
            logging.error(f"删除接口测试用例行过程中发生错误: {str(e)}")
            msg_box = QMessageBox(QMessageBox.Critical, "删除错误", f"删除接口测试用例行过程中发生错误: {str(e)}")
            msg_box.exec_()

    def insert_selected_cells(self):
        try:
            current_index = self.stackedWidget.currentIndex()
            if current_index == 1:
                selected_ranges = self.tableWidget_5.selectedRanges()
                if not selected_ranges:
                    QMessageBox.warning(None, "警告", "未选中任何单元格")
                    return
                selected_row = selected_ranges[0].bottomRow()
                insert_row = selected_row + 1
                self.GeneralTest.add_row_2(insert_row)
                return
            if  current_index == 2:
                table_widget = self.tableWidget_2
            elif current_index == 3:
                table_widget = self.tableWidget_3
            elif current_index == 4:
                table_widget = self.tableWidget_4
            else:
                QMessageBox.warning(None, "警告", "当前页面没有对应的表格控件")
                return

            selected_ranges = table_widget.selectedRanges()
            if not selected_ranges:
                QMessageBox.warning(None, "警告", "未选中任何单元格")
                return
            selected_row = selected_ranges[0].bottomRow()
            insert_row = selected_row + 1
            table_widget.insertRow(insert_row)

            column_count = table_widget.columnCount()
            for col in range(column_count):
                if col == column_count - 1:
                    # 最后一列添加“删除”按钮
                    delete_button = QPushButton("删除")
                    delete_button.clicked.connect(self.delete_clicked)
                    table_widget.setCellWidget(insert_row, col, delete_button)
                else:
                    # 其余列插入空白单元格
                    table_widget.setItem(insert_row, col, QTableWidgetItem(""))
        except Exception as e:
            print(traceback.format_exc())
            QMessageBox.critical(None, "错误", f"插入行时发生错误:\n{str(e)}")

    def closeEvent(self, event):
        # 弹出提示框询问用户是否保存更改并退出
        reply = QtWidgets.QMessageBox.question(
            self, u'温馨提示', u'是否退出?',
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        if reply == QtWidgets.QMessageBox.Yes:
            event.accept()  # 关闭窗口
        elif reply == QtWidgets.QMessageBox.No:
            event.accept()  # 不保存直接关闭窗口
        else:
            event.ignore()  # 取消退出操作，保持窗口打开


    def open_file(self):
        try:
            # 检查是否是第一次调用
            if self.first_open_file_call:
                QMessageBox.information(None, "温馨提示！",
                                        "导入协议文件后，别忘了导入map文件，最后才是导入对应的测试文件（仅第一次导入文件时提醒）")
                self.first_open_file_call = False  # 设置标志变量为False，表示已经显示过了

            self.openfile = self.open_file or os.getcwd()
            fname = QFileDialog.getOpenFileName(None, "选择文件", '/', "xlsx Files(*.xlsx)")
            self.filename = fname[0]

            if not self.filename:
                QMessageBox.warning(None, "提示", "未选择任何文件！")
                return

            book = xlrd2.open_workbook(self.filename)
            required_sheets = [
                'CAN用例解析矩阵',
                'LIN用例解析矩阵',
                '测试盒CAN通讯矩阵',
                '电阻模拟输出表',
                '硬件配置表',
                '软件接口定义表'
            ]

            # 检查所有必需的工作表是否存在
            missing_sheets = [sheet for sheet in required_sheets if sheet not in book.sheet_names()]
            if missing_sheets:
                error_message = f"文件 '{self.filename}' 缺少以下工作表: {', '.join(missing_sheets)}\n导入失败！"
                logging.error(error_message)
                QMessageBox.critical(None, "导入失败", error_message)
                return

            # 如果所有必需的工作表都存在，则继续处理
            sheet = book.sheet_by_name('CAN用例解析矩阵')
            datalist = []
            for i in range(1, sheet.nrows):
                values = sheet.row_values(i)
                if values[25] != '':
                    item = values[25].replace("：", ':').replace('\r', '').replace(';', '').replace('；', '')
                    item = item.split('\n')
                    signalname = values[8].replace('\n', '').replace('\r', '')
                    item2 = [signalname] + item
                    self.signalD1.append(item2)
                    for j in range(len(item)):
                        if ':' in item[j]:
                            item[j] = item[j].split(":")[1]
                        else:
                            item.remove(item[j])
                    item1 = [signalname] + item
                    self.signalD.append(item1)
                if values[8] != '' and values[0] != '' and values[1] == '':
                    canN = values[0].replace('\r', '').replace('\n', '')
                    signalN = values[8].replace('\r', '').replace('\n', '')
                    data1 = [canN, signalN, values[15], values[16], values[11], values[13], values[10]]
                    datalist.append(signalN)
                    self.diclist.append(data1)
            datalist = list(set(datalist))
            self.datalist = sorted(datalist, key=str.lower)
            data1 = []
            for k in range(sheet.nrows - 1, 0, -1):
                values = sheet.row_values(k)
                if values[5] == '' and values[8] != '':
                    data1.append(values[8].replace('\r', '').replace('\n', ''))
                elif values[5] != "":
                    self.Canname = values[5]
                    if data1 != [] and self.Canname != '':
                        self.diclist1.append([self.Canname, values[3], values[2]] + data1)
                        data1 = []

            sheet5 = book.sheet_by_name('LIN用例解析矩阵')
            linlist = []
            for i in range(1, sheet5.nrows):
                values = sheet5.row_values(i)
                if values[25] != '':
                    item = values[25].replace("：", ':').replace('\r', '').replace(';', '').replace('；', '')
                    item = item.split('\n')
                    signalname = values[8].replace('\n', '').replace('\r', '')
                    item2 = [signalname] + item
                    self.linD1.append(item2)
                    for j in range(len(item)):
                        if ':' in item[j]:
                            item[j] = item[j].split(":")[1]
                        else:
                            item.remove(item[j])
                    item1 = [signalname] + item
                    self.linD.append(item1)
                if values[8] != '' and values[0] != '' and values[1] == '':
                    linN = values[0].replace('\r', '').replace('\n', '')
                    signalN = values[8].replace('\r', '').replace('\n', '')
                    data1 = [linN, signalN, values[15], values[16], values[11], values[13], values[10]]
                    linlist.append(signalN)
                    self.lindic.append(data1)
            linlist = list(set(linlist))
            self.linlist = sorted(linlist, key=str.lower)
            data2 = []
            for k in range(sheet5.nrows - 1, 0, -1):
                values = sheet5.row_values(k)
                if values[6] == '' and values[8] != '':
                    data2.append(values[8].replace('\r', '').replace('\n', ''))
                elif values[6] != "":
                    self.Linname = values[6]
                    if data2 != [] and self.Linname != '':
                        self.lindic1.append([self.Linname, values[3], values[2]] + data2)
                        data2 = []

            sheet2 = book.sheet_by_name('测试盒CAN通讯矩阵')
            self.candata = []
            self.cansignal1 = []
            self.cansignal = []
            self.candiclist = []
            self.candiclist1 = []
            candata1 = []
            Vinput1 = []
            for m in range(sheet2.nrows - 1, 0, -1):
                values = sheet2.row_values(m)
                if values[5] == '' and values[8] != '':
                    candata1.append(values[8].replace('\r', '').replace('\n', ''))
                elif values[5] != "":
                    self.Canname = values[5]
                    if candata1 != [] and self.Canname != '':
                        self.candiclist1.append([values[0], self.Canname, values[3]] + candata1)
                        candata1 = []
            for n in range(1, sheet2.nrows):
                value = sheet2.row_values(n)
                if value[8] in self.candiclist1[0]:
                    self.candata.append(value[8])
                    canlist = value[25].replace('\r', '').split('\n')
                    data = [value[8]]
                    data2 = [value[8]]
                    for l in canlist:
                        if ":" in l and l != '':
                            data.append(l)
                            param = l.split(':')[1]
                            data2.append(param)
                    self.cansignal1.append(data)
                    self.cansignal.append(data2)
                    signalN = value[8].replace('\r', '').replace('\n', '')
                    data1 = [signalN, value[10], value[11], value[13]]
                    self.candiclist.append(data1)
            self.candata = list(set(self.candata))

            sheet6 = book.sheet_by_name('电阻模拟输出表')
            for k in range(0, sheet6.ncols, 2):
                dic = {}
                temp = sheet6.col_values(k)[4:]
                ch = sheet6.col_values(k + 1)[4:]
                for u in range(len(temp)):
                    dic[temp[u]] = ch[u]
                self.RTlist.append(dic)

            sheet3 = book.sheet_by_name('硬件配置表')
            self.ID = sheet3.row_values(2)[4:12]
            for l in range(2, 12):
                lst = sheet3.col_values(l)
                result = list(filter(None, lst))
                if l == 2:
                    self.CAN = result[2:]
                elif l == 3:
                    self.LIN = result[2:]
                elif l == 9:
                    self.openclose = result[2:-1]
                elif l == 4:
                    self.temp = result[2:-1]
                elif l == 5:
                    self.CV = result[2:-1]
                elif l == 6:
                    self.Vinput = result[2:-1]
                elif l == 7:
                    self.PWM = result[2:-1]
                elif l == 10:
                    self.outputPWM = result[2:-1]
                elif l == 11:
                    self.inputPWM = result[2:-1]
                elif l == 8:
                    self.banqiao = result[2:-1]
            lst = sheet3.col_values(18)
            result = list(filter(None, lst))
            self.comboBoxList3 = result[2:]

            sheet4 = book.sheet_by_name('软件接口定义表')
            self.interfaceList = [s.replace('\n', '').replace('\t', '').strip() for s in sheet4.col_values(2)[1:]]
            for k in range(1, sheet4.nrows):
                values = [str(s).replace('\n', '').replace('\t', '').strip() for s in sheet4.row_values(k)]
                self.interfaceDic.append(values)

            # 成功导入后的提示框
            QMessageBox.information(None, "导入成功", f"文件 '{self.filename}' 导入成功！")

        except Exception as e:
            logging.error(f"文件导入失败：{traceback.format_exc()}")
            QMessageBox.critical(None, "导入失败", f"文件 '{self.filename}' 导入失败！\n错误详情：{str(e)}")

    #导入ELF文件
    def getELFAnalysis(self):
        # 提取文件过滤器和默认路径为类属性，提高灵活性
        file_filter = "elf Files(*.elf);;out files (*.out)"
        default_path = self.last_open_path if hasattr(self, 'last_open_path') else '/'

        # 打开文件选择对话框
        path = QFileDialog.getOpenFileName(self, "选择文件", default_path, file_filter)
        self.elfpath = path[0]

        # 如果用户取消选择，给出友好提示
        if not self.elfpath:
            msg_box = QMessageBox(QMessageBox.Warning, "提示", "未选择文件！")
            msg_box.exec_()
            return

        # 记住上次打开的路径
        self.last_open_path = os.path.dirname(self.elfpath)

        try:
            # 启动后台线程加载 ELF 文件
            self.thread = LoadELFThread(self.elfAnalysis, self.elfpath)
            self.thread.finished.connect(self.onLoadFinished)
            self.thread.error.connect(self.onLoadError)
            self.thread.start()

            # 使用 QProgressDialog 显示加载中的提示信息
            self.progress_dialog = QProgressDialog("正在加载文件，请稍候...", "取消", 0, 0, self)
            self.progress_dialog.setWindowModality(Qt.WindowModal)
            self.progress_dialog.setWindowTitle("加载中")
            self.progress_dialog.show()

            # 连接取消按钮到终止线程的槽函数
            self.progress_dialog.canceled.connect(self.thread.requestInterruption)

            # 确保在加载完成后关闭进度对话框
            self.thread.finished.connect(self.progress_dialog.close)
            self.thread.error.connect(self.progress_dialog.close)

        except Exception as e:
            # 捕获异常并提示用户
            logging.error(f"文件导入失败：{str(e)}")
            error_title = "导入失败"
            error_message = f"文件导入失败：{str(e)}"
            msg_box = QMessageBox(QMessageBox.Critical, error_title, error_message)
            msg_box.exec_()

    def onLoadFinished(self, message):
        self.progress_dialog.close()
        QMessageBox.information(self, "导入成功", message)

    def onLoadError(self, message):
        self.progress_dialog.close()
        QMessageBox.critical(self, "导入失败", message)


#加载ELF文件线程
class LoadELFThread(QThread):
    finished = pyqtSignal(str)  # 成功信号，传递成功消息
    error = pyqtSignal(str)     # 错误信号，传递错误消息

    def __init__(self, elf_analysis, elf_path):
        super().__init__()
        self.elf_analysis = elf_analysis
        self.elf_path = elf_path

    def run(self):
        try:
            # 模拟加载文件的过程，这里使用一个循环来模拟长时间操作
            for i in range(100):
                if self.isInterruptionRequested():
                    self.error.emit(f"MAP文件加载被取消")
                    return
                time.sleep(0.1)  # 模拟加载时间

            self.elf_analysis.loadPath(self.elf_path)
            self.finished.emit(f"文件 '{self.elf_path}' 导入成功！")
        except FileNotFoundError:
            self.error.emit(f"文件 '{self.elf_path}' 不存在，请检查路径是否正确。")
        except PermissionError:
            self.error.emit(f"无法访问文件 '{self.elf_path}'，请检查文件权限。")
        except Exception as e:
            self.error.emit(f"文件导入失败：{str(e)}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(':./sign.png'))
    win = UiMain()
    win.show()
    sys.exit(app.exec_())
