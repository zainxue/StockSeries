#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xlrd
from Stocks import GetData
from pandas import Series,DataFrame
import datetime
import re

def todate(self):
    sub = re.findall(r"\d+\.?\d*", self)
    return datetime.date(int(sub[0]),int(sub[1]),int(sub[2]))

if __name__=='__main__':
    filepath = (r'E:\zain\newproject\StockSeries\HSI.xlsx')
    Workbook = xlrd.open_workbook(filepath)

    Rawdata = Workbook.sheet_by_index(0)
    HSIdata = []

    sublist = []
    for r in range(1,len(Rawdata.col_values(0))):
        subcell = GetData.GetExcelDate(Rawdata.col_values(0)[r])
        sublist.append(subcell)

    subpricelist = Rawdata.col_values(1)
    del subpricelist[0]
    testdata = DataFrame(subpricelist,index=sublist,columns=['HSIPrice'])
    print(testdata)
    print(testdata.loc[todate('1986/12-31')])
    print(testdata[(testdata.HSIPrice==2568.300049)].index)
