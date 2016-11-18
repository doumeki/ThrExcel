#!/usr/bin/python
# -*- coding:utf-8 -*-
#@Author DOUMEKI

import ThrExcel,re
import sys
class CopyExpress:
    def __init__(self):
        self.table = []
        self.express = ''
        self.expressopt = [] #计算公式,存放操作数
        self.eft = {} #公式中的操作数对应的tabletitle K为title reg, V为表达式操作数
        self.titlereg = [] #表格的reg
        self.excel = None
        self.sheet = None
        self.calcValues = []

    def openExcel(self,filename,sheetname):
        # self.excel = ThrExcel.ThrExcel(filename="C:\\Users\\Administrator\\Desktop\\mytest.csv")
        # self.sheet = self.excel.getSheet("mytest")
        try:
            self.excel = ThrExcel.ThrExcel(filename=filename)
            self.sheet = self.excel.getSheet(sheetname)
        except Exception as e:
            raise e

    def copyExpressAnlys(self,express):
        #titlereg = '(Ulsmgres + RRSNkuks.avg)/Lupussh'
        self.express = express
        op_partner = r'[\(\\+\-\*\\/\)]'
        self.expressopt = list(filter(lambda x: x != '', re.split(op_partner, express)))

    def expressMapping(self):
        tabletitle = []
        for t in self.table:
            tabletitle.extend(list(t.keys()))
        t = sorted(tabletitle)
        ops = {}
        # print (self.table)
        for eo in self.expressopt:
            sortedtemp = []
            reg = self.eft[eo]
            for title in t:
                if re.match(reg,title):
                    for kv in self.table:
                       for k,v in kv.items():
                           if k == title:
                               sortedtemp.extend(v)
            ops[eo] = sortedtemp
        print (ops)
        self.calcValues = ops

    def calcexpress(self):
        map(self.runexpress(),self.calcValues)

    def runexpress(self):
        rag = len(self.calcValues[self.expressopt[0]])
        allexpress = []
        for i in range(rag):
            tempexpress = self.express
            for exo in self.expressopt:
                for k,v in self.calcValues.items():
                    if exo == k:
                        try:
                            vl = v.pop(0)
                        except:
                            vl = 0
                        tempexpress = tempexpress.replace(exo,str(vl))
            allexpress.append(tempexpress)
        allresult = []
        for exp in allexpress:
            try:
                allresult.append(eval(exp))
            except:
                allresult.append("None")
        print(allresult)







    def getColumnMapping(self, tableheadreg):
        try:
            self.getColumnValues(tableheadreg)
        except Exception as e:
            raise e
        finally:
            self.excel.close()

    def getColumnValues(self, tableheadreg):
        for reg in tableheadreg:
            tablecolumn = self.sheet.getColumnCellsByTableName(reg, getValue=True)
            self.table.append(tablecolumn)


if __name__ == "__main__":
    #if len(sys.argv) != 3:
        #raise Exception("请输入两Excel名称和Sheet名称")
    #else:
    ce = CopyExpress()
    #ce.openExcel(sys.argv[1],sys.argv[2])
    ce.openExcel('C:\\Users\\Administrator\\Desktop\\mytest.csv','mytest')
    # ip = input("请输入表达式")
    ip = '(Ulsmgres + RRSNkuks.avg)/Lupussh'
    ce.copyExpressAnlys(ip)
    print("请输入表达式对应表的关系，使用正则表达式关系，请注意转义符号的使用")
    ce.titlereg = ['L2_USERCHR_SCH_INFO\(0\).DRB_512MS\(\d\).PUSCH_DRB_512MS\(0\).ulPuschMcsSum',\
               'L2_USERCHR_SCH_INFO\(0\).DRB_512MS\(\d\).ulPuschSchSum',\
               'L2_USERCHR_VOLTE_SCH_INFO\(0\).DRB_512MS\(0\).PUSCH_DRB_512MS\(\d\).ulPuschMcsSum']
    for op,e in zip(ce.titlereg, ce.expressopt):
        # ip = input("请输入" + exp + "对应用的列表")
        # titlereg.append(ip)
        ce.eft[e] = op
    ce.getColumnMapping(ce.titlereg)
    ce.expressMapping()
    ce.runexpress()
    # ce.expressMapping()
    # print(ce.table)
    # ce.excel.close()

    '''
    LLC.Rank0.GRS	LLC.Rank0.GRS	UlemssK0.RRC1.DMK	UlemssK0.RRC0.DMK	PPSHCsum

    '''