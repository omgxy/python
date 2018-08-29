

# coding=utf-8
# @File  : OdaClass.py
# @Author: lijun4
# @Date  : 2018/8/17
# @Desc  :统计ODA每日交易文件，并生成到Excel中。
import re
import sys
import shutil,os
import operator
import xlwt
import xlrd
from xlutils.copy import copy
from xlwt import Style


jiaoyi_list =['日期','原始消费交易总笔数','原始消费交易成功笔数','原始消费成功率','再请款交易总笔数','再请款交易成功笔数','再请款交易成功率','云闪付交易总笔数','云闪付交易成功笔数','云闪付交易成功率','云闪付再请款总笔数','云闪付再请款成功笔数','云闪付再请款成功率','接入机构总数','交易量top1机构','top1交易量','交易量top2机构','top2交易量','交易量top3机构','top3交易量','交易量top4机构','top4交易量','交易量top5机构','top5交易量']
class OdaClass(object):
    def __index__(self,filePath,fileName,fileNameBAK,sheetName,dirList):
        self.filePath = filePath
        self.fileName = fileName
        self.fileNameBAK = fileNameBAK
        self.sheetName = sheetName
        self.dirList = dirList

    def excelIsExit(self):
        # 判断文件是否存在，若存在，则先备份，后删除文件
        if os.path.exists(self.fileName):
            shutil.copy(self.fileName, self.fileNameBAK)
            os.unlink(self.fileName)

    def set_style(self, name, height, bold=False):
        style = xlwt.XFStyle()  # 初始化样式

        font = xlwt.Font()  # 为样式创建字体
        font.name = name
        font.bold = bold  # 粗体
        font.color_index = 4
        font.height = height

        style.font = font
        return style

    def createExcel(self):

        self.excelIsExit()
        # 创建工作簿
        workbook = xlwt.Workbook(encoding='utf-8')
        # 创建sheet
        data_sheet = workbook.add_sheet('每日交易统计总量')
        data_sheet1 = workbook.add_sheet('交易排行榜')
        raw = 0
        for i in jiaoyi_list:
            # print(i)
            data_sheet.write(raw, 0, i, self.set_style('Times New Roman', 220, True))
            raw = raw + 1
        # 保存文件
        workbook.save('oda.xls')

    def cdPath(self):
        self.dirList = os.listdir(self.filePath)
        self.dirList.sort()
        os.chdir(self.filePath)

    def parseOdaData(self,fileName, savefilename, sheetName):


        count = 0  # 原始消费交易总笔数
        success_count = 0  # 原始消费交易成功笔数
        succ_per = 0  # 原始消费成功率
        reptran_count = 0  # 再请款交易总笔数
        reptran_success_count = 0  # 再请款交易成功总笔数
        rep_succ_per = 0  # 再请款交易成功率
        tsp_conut = 0  # 云闪付交易总笔数
        tsp_success_count = 0  # 云闪付交易成功笔数
        tsp_succ_per = 0  # 云闪付交易成功率
        tsp_reqtrans_count = 0  # 云闪付再请款总笔数
        tsp_reqtrans_succ_count = 0  # 云闪付再请款成功笔数
        tsp_reqtrans_succ_per = 0  # 云闪付再请款成功率

        countDict = {}


        with open(fileName) as file:
            for line in file.readlines():
                line = line.strip()
                if not len(line):
                    pass
                else:
                    if line.startswith('================机构[') and line.endswith(']交易信息统计==============='):
                        name = line.split('[')[1].split(']')[0]
                    elif line.startswith('===========================概要统计END==============================='):
                        break

                    elif line.startswith('原始消费交易总笔数'):
                        line = line.split("：")[1]
                        line = line.strip()
                        countDict[name]=int(line)
                        count = count + int(line)
                    elif line.startswith('原始消费交易成功笔数 '):
                        line = line.split("：")[1].split("，")[0]
                        line = line.strip()
                        success_count = success_count + int(line)
                    elif line.startswith('再请款交易总笔数'):
                        line = line.split("：")[1]
                        line = line.strip()
                        reptran_count = reptran_count + int(line)
                    elif line.startswith('再请款交易成功笔数'):
                        line = line.split("：")[1].split("，")[0]
                        line = line.strip()
                        reptran_success_count = reptran_success_count + int(line)
                    elif line.startswith("云闪付交易总笔数"):
                        line = line.split("：")[1]
                        line = line.strip()
                        tsp_conut = tsp_conut + int(line)
                    elif line.startswith("云闪付交易成功笔数"):
                        line = line.split("：")[1].split("，")[0]
                        line = line.strip()
                        tsp_success_count = tsp_success_count + int(line)
                    elif line.startswith("云闪付再请款总笔数"):
                        line = line.split("：")[1]
                        line = line.strip()
                        tsp_reqtrans_count = tsp_reqtrans_count + int(line)
                    elif line.startswith("云闪付再请款成功笔数"):
                        line = line.split("：")[1].split("，")[0]
                        line = line.strip()
                        tsp_reqtrans_succ_count = tsp_reqtrans_succ_count + int(line)

        succ_per = success_count / count
        rep_succ_per = reptran_success_count / reptran_count
        tsp_succ_per = tsp_success_count / tsp_conut
        tsp_reqtrans_succ_per = tsp_reqtrans_succ_count / tsp_reqtrans_count
        print('日期                   : ', fileName[13:21])
        print('原始消费交易总笔数      : ', count)
        print('原始消费交易成功笔数    : ', success_count)
        print('再请款交易总笔数        : ', reptran_count)
        print('再请款交易成功总笔数    : ', reptran_success_count)
        print('云闪付交易总笔数        : ', tsp_conut)
        print('云闪付交易成功笔数      : ', tsp_success_count)
        print('云闪付再请款总笔数      : ', tsp_reqtrans_count)
        print('云闪付再请款成功笔数    : ', tsp_reqtrans_succ_count)

        #sorted(countDict.items(), key=operator.itemgetter(1))
        countDict1 = sorted(countDict.items(), key=lambda x: x[1], reverse=True)
        #print(countDict1)

        workbook = xlrd.open_workbook(savefilename)
        Charts_sheet = workbook.sheet_by_name(u'每日交易统计总量')#交易总量统计
        Charts_sheet_up = workbook.sheet_by_name(u'交易排行榜')  # 交易总量统计
        # print(Charts_sheet.nrows)
        # print(Charts_sheet.ncols)

        rb = xlrd.open_workbook(savefilename, formatting_info=True)
        wb = copy(rb)
        ws = wb.get_sheet(sheetName)
        ws.write(0, Charts_sheet.ncols, fileName[13:21], self.set_style('Times New Roman', 220, False))
        ws.write(1, Charts_sheet.ncols, count, self.set_style('Times New Roman', 220, False))
        ws.write(2, Charts_sheet.ncols, success_count, self.set_style('Times New Roman', 220, False))
        ws.write(3, Charts_sheet.ncols, succ_per, self.set_style('Times New Roman', 220, False))
        ws.write(4, Charts_sheet.ncols, reptran_count, self.set_style('Times New Roman', 220, False))
        ws.write(5, Charts_sheet.ncols, reptran_success_count, self.set_style('Times New Roman', 220, False))
        ws.write(6, Charts_sheet.ncols, rep_succ_per, self.set_style('Times New Roman', 220, False))
        ws.write(7, Charts_sheet.ncols, tsp_conut, self.set_style('Times New Roman', 220, False))
        ws.write(8, Charts_sheet.ncols, tsp_success_count, self.set_style('Times New Roman', 220, False))
        ws.write(9, Charts_sheet.ncols, tsp_succ_per, self.set_style('Times New Roman', 220, False))
        ws.write(10, Charts_sheet.ncols, tsp_reqtrans_count, self.set_style('Times New Roman', 220, False))
        ws.write(11, Charts_sheet.ncols, tsp_reqtrans_succ_count, self.set_style('Times New Roman', 220, False))
        ws.write(12, Charts_sheet.ncols, tsp_reqtrans_succ_per, self.set_style('Times New Roman', 220, False))
        ws.write(13, Charts_sheet.ncols, len(countDict1), self.set_style('Times New Roman', 220, False))

        ws = wb.get_sheet('交易排行榜')
        i = 1
        ws.write(0, Charts_sheet_up.ncols, fileName[13:21], self.set_style('Times New Roman', 220, False))
        for key,value in countDict1:

            #ws.write(i, Charts_sheet_up.ncols, key+':'+str(value), self.set_style('Times New Roman', 220, False))
            i = i+1





        os.remove(savefilename)
        wb.save(savefilename)

    # 创建Excel
    def createOdaData(self):
        for i in self.dirList:
            if os.path.splitext(i)[1] == ".log":
                self.parseOdaData(i, self.fileName, self.sheetName)


#主程序入口
oda = OdaClass()
oda.filePath = "D://3.git//oddrs//oddrs_doc//12.生产相关//3.ODA交易量//信总统计数据"
oda.fileName = "oda.xls"
oda.fileNameBAK = "oda_bak.xls"
oda.sheetName = "每日交易统计总量"
oda.cdPath()
oda.createExcel()
oda.createOdaData()


