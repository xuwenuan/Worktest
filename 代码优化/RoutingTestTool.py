import configparser
import os
import re
import logging
import traceback
import xlrd2
import csv
from xlsxwriter import Workbook


class RoutingTest:
    logging.basicConfig(
            filename="log.log",
            filemode="w",
            datefmt="%Y-%m-%d %H:%M:%S %p",
            format="%(asctime)s - %(name)s - %(levelname)s - %(module)s: %(message)s",
            level=40
        )
    routingList = []
    CANList = []
    def readMesseage(self):
        try:
            if path1 and path2:
                with open(path1, 'r') as file:
                    reader = csv.reader(file)
                    for row in reader:
                        values = list(row)
                        self.routingList.append(values)
                book2 = xlrd2.open_workbook(path2)
                sheetbook = book2.sheet_by_name('硬件配置表')
                self.CANList = sheetbook.col_values(2)[3:]
                self.genautotable()
        except:
            logging.error(traceback.format_exc() + "\n")

    def genautotable(self):
        self.routingList = self.routingList[1:]
        dataList=[]
        for item in self.routingList:
            preview = []
            try:
                data1 = '0'+str(self.CANList.index(item[0]+'_CAN'))+'00'
            except:
                data1 = '0000'
            if item[3] == '1':
                ID = '8'+item[2].replace('0x', '').replace('0X', '').zfill(3)
            else:
                ID = item[2].replace('0x', '').replace('0X', '').zfill(4)
            if int(item[5]) < 20:
                data3 = '0014'
                cycle = '20'
            else:
                data3 = str(hex(int(item[5]))).upper()[2:].zfill(4)
                cycle = item[5]
            data2 = ID + str(hex(int(cycle) * 5)).upper()[2:].zfill(4)
            data4 = str(hex((int(item[4])-1)*8)).upper()[2:].zfill(4)
            data5 = str(hex(int(item[4])*8)).upper()[2:].zfill(4)
            message1 = data1+data2+data3+data4+data5
            Mac = ' '.join(message1[i:i + 2] for i in range(0, len(message1), 2)).upper()
            if item[1] != '/':
                name = item[0]+'CAN_'+item[2]+"_"+item[1]
            else:
                name = item[0]+'CAN_'+item[2]+"_"
            data = ['event', name, Mac, cycle, '3', '0x700', '', '', '--', 'Motorola']
            dataList.append(data)
            preview.append(data)
            seconditem = ['check', '单条测试用例处理状态', '0', cycle, '100', '0x710', '16', '8', '--', 'Motorola']
            dataList.append(seconditem)
            preview.append(seconditem)
            s = 'FF'*(int(item[4])+1)
            string = ' '.join(s[i:i + 2] for i in range(2, len(s), 2)).upper()
            thirditem = ['event', name, string, cycle, '3', item[2], '', '', '--', 'Motorola']
            anthirditem = ['event', name, string.replace('FF', '00'), cycle, '3', item[2], '', '', '--', 'Motorola']
            dataList.append(thirditem)
            preview.append(anthirditem)
            if item[8] != '/':
                name = item[7]+'CAN_'+item[9]+"_"+item[8]
            else:
                name = item[7]+'CAN_'+item[9]+"_"
            try:
                channel = '0' + str(self.CANList.index(item[7] + '_CAN')) + '00'
            except:
                channel = '0000'
            if item[10] == '1':
                ID = '8'+item[9].replace('0x', '').replace('0X', '').zfill(3)
            else:
                ID = item[9].replace('0x', '').replace('0X', '').zfill(4)
            if int(item[12]) < 20:
                data3 = '0014'
                cycle2 = '20'
            else:
                data3 = str(hex(int(item[12]))).upper()[2:].zfill(4)
                cycle2 = item[12]

            data4 = str(hex((int(item[11]) - 1) * 8)).upper()[2:].zfill(4)
            data5 = str(hex(int(item[11]) * 8)).upper()[2:].zfill(4)
            message2 = channel+ID+str(hex(int(cycle2)*5)).upper()[2:].zfill(4)+data3+data4+data5
            Mac2 = ' '.join(message2[i:i + 2] for i in range(0, len(message2), 2)).upper()
            forthitem = ['event', name, Mac2, cycle2, '3', '0x701', '', '', '--', 'Motorola']
            dataList.append(forthitem)
            preview.append(forthitem)
            fifthitem = ['check', '单条测试用例处理状态', '1', int(cycle2)*5, '100', '0x710', '16', '8', '--', 'Motorola']
            dataList.append(fifthitem)
            preview.append(fifthitem)
            t = 'FF' * (int(item[11])+1)
            text = ' '.join(t[i:i + 2] for i in range(2, len(t), 2)).upper()
            sixth = ['routercheck', name, text, cycle2, '100', item[9], '', '', '--', 'Motorola']
            ansixth = ['routercheck', name, text.replace('FF', '00'), cycle2, '100', item[9], '', '', '--', 'Motorola']
            dataList.append(sixth)
            preview.append(ansixth)
            dataList = dataList+preview
            text = item[0] + ' ' + item[1] + ' ' + item[2] + '      ' + item[7] + ' ' + item[8] + ' ' + item[9]
            note = ['测试用例说明', text, '12', '', '', '', '', '', '', '']
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
    configpath = './configpath.ini'
    if configpath:
        config = configparser.ConfigParser()
        config.read(configpath,encoding='utf-8')
        path1 = config.get('path', 'routingTestPATH')
        path2 = config.get('path', 'Protocolpath')
        path3 = config.get('path', 'autopath')
    # path1="C:\\Users\\pc\\Downloads\\MessageRoutingTestTable_V01.csv"
    # path2="C:\\Users\\pc\\Downloads\\自动化测试盒协议_BDM_T1EJ.xlsx"
        aa = RoutingTest()
        aa.readMesseage()

