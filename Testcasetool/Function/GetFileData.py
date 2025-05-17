import xlrd2

class GetFileData():
    def readCANConfig(self,path):
        book = xlrd2.open_workbook(path)
        sheet = book.sheet_by_name('测试盒CAN通讯矩阵')
        itemwithId =[]
        candata = []
        for m in range(sheet.nrows - 1, 0, -1):
            values = sheet.row_values(m)
            if values[5] == '' and values[8] != '':
                candata.append(values[8].replace('\r', '').replace('\n', ''))
            elif values[5] != "":
                self.Canname = values[5]
                if candata != [] and self.Canname != '':
                    itemwithId.append([values[0], self.Canname, values[3]] + candata)
                    candata = []
        InformationList=[]
        for n in range(1,sheet.nrows):
            item =[]
            values = sheet.row_values(n)
            for i in range(len(itemwithId)):
                if values[8]!='' and values[8] in itemwithId[i]:
                    # id，信号名称，起始位，长度，信号描述
                    item = [itemwithId[i][2],values[8].replace("：",':').replace('\r','').replace(';','').replace('；',''),values[11],values[13],values[25],values[15]]
                    break
            if item!=[]:
                InformationList.append(item)
        return InformationList

    def getBit(self,channel,param,index,path):
        index = str(index)
        num =''
        startbit=''
        length=''
        channeldic ={'开关输入':'706','电压输入':'704','PWM输入':'708'}
        paramdic = {'LSD':'0','HSD':'1','高阻':'2'}
        datalist = self.readCANConfig(path)
        for item in datalist:
            if (channel=='电压输入' or channel=='PWM输入') and channeldic[channel] in item[0] and index in item[1]:
                startbit = item[2]
                length = item[3]
                num = param
                break
            elif channel=='开关输入' and channeldic[channel] in item[0] and index in item[1]:
                num = paramdic[param]
                startbit = item[2]
                length = item[3]
                break
        return num,startbit,length


    def getresolution(self,signalname,path):
        resolution=''
        datalist = self.readCANConfig(path)
        for item in datalist:
            if item[1]==signalname or signalname in item[1]:
                resolution=item[-1]
                break
        return resolution




if __name__ == '__main__':
    path = r"C:\\Users\\se25atk\\Downloads\\自动化测试盒协议_BDM_Envoy.xlsx"
    data = GetFileData()
    data.readCANConfig(path)