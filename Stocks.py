#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import xlrd
import datetime
import time
import re

class GetData():
    def GetGuoXinDataList(filepath):
        wb = xlrd.open_workbook(filepath)
        RawData = wb.sheet_by_index(0)
        StockData = []
        for c in range(1,6):
            subdata = []
            for r in range(4,len(RawData.col_values(c))):
                subcell = RawData.col_values(0)[r]
                if u'\u4e00' <= subcell <= u'\u9fff':
                    continue
                subcell = str(subcell)
                subcell = re.sub(' ','',subcell)
                subcell = re.sub('/','-',subcell)
                subcell = datetime.datetime.strptime(subcell,'%Y-%m-%d')
                subcell = datetime.date(subcell.year,subcell.month,subcell.day)
                subdata.append({subcell:RawData.col_values(c)[r]})
            StockData.append(subdata)
        return StockData

    def GetExcelDate(date):
        rawdate = datetime.date(1899, 12, 31).toordinal() - 1
        if isinstance(date,float):
            date = int(date)
        date = datetime.date.fromordinal(rawdate + date)
        return date

if __name__=='__main__':
    filepath = r'E:\zain\data colletion\600807.xlsx'
    data = GetData.GetGuoXinDataList(filepath)
    print(data[4])
    filepath = (r'E:\zain\newproject\researchproject\HSI.xlsx')
    Workbook = xlrd.open_workbook(filepath)
    Rawdata = Workbook.sheet_by_index(0)
    HSIdata = []
    print(Rawdata.col_values(0)[2])
    print(GetData.GetExcelDate(Rawdata.col_values(0)[2]))