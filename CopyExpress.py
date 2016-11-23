#!/usr/bin/python
# -*- coding:utf-8 -*-
#@Author DOUMEKI

import ThrExcel,re
import sys,math
class CopyExpress:
    def __init__(self):
        self.table = {}
        self.express = ''
        self.expressopt = [] #计算公式,存放操作数
        self.eft = {} #公式中的操作数对应的tabletitle
        self.titlereg = [] #表格的reg
        self.excel = None
        self.sheet = None
        self.alltitles = {}
        self.registedExpress = ['AVG','SIN'] #已注册的公共方法，//todo://后期作为配置项来做
        self.callanditem = []
        self.reportexcel = None

    def clearUp(self):
        self.table.clear()
        self.express = ''
        self.expressopt.clear()
        self.eft.clear()
        self.titlereg.clear()
        self.alltitles.clear()
        self.callanditem.clear()

    def openExcel(self,filename,sheetname):
        # self.excel = ThrExcel.ThrExcel(filename="C:\\Users\\Administrator\\Desktop\\mytest.csv")
        # self.sheet = self.excel.getSheet("mytest")
        try:
            self.excel = ThrExcel.ThrExcel(filename=filename,visiable=0)
            self.sheet = self.excel.getSheet(sheetname)
        except Exception as e:
            raise e

    def copyExpressAnlys(self,express):
        self.express = express
        pattern = r'[\+\-\*\/\s*]'
        pattern2 = r'[\(*\)*\s*]'
        first = re.split(pattern,express)
        expressopt = set()
        for oe in first:
            temp = re.split(pattern2, oe)
            for t in temp:
                if t != '' and t not in self.registedExpress and not t.isdigit(): #非空，非注册函数，非纯数字
                    expressopt.add(t)
        self.expressopt = list(expressopt)
        # expressopt = list(filter(lambda x: x != '', re.split(op_partner, express)))
        # self.expressopt = list(filter(lambda x : x not in self.registedExpress,expressopt))


    def expressMapping2(self):
        for exo in self.expressopt:
            ktitle = []
            for k in self.table.keys():
                if re.match(self.eft[exo],k) or self.eft[exo] == k:
                    ktitle.append(k)
            ktitle =  sorted(ktitle)
            self.alltitles[exo] = ktitle #别名 -> 真实列名，已排序

    # 返回每行的平均值
    def runexpress2(self):
        optlen = len(list(self.alltitles.keys())) # 每行的操作数的个数
        titlelen = 0
        for k,v in self.alltitles.items():
            if len(v) > titlelen:
                titlelen = len(v)
        # titlelen = len(list(self.alltitles[self.expressopt[0]])) #title 行的个数
        valuelen = len(self.table[self.alltitles[self.expressopt[0]][0]]) #列的个数

        allexpress = self._gen_express(optlen, titlelen, valuelen) #生成每个的表达式计算值
        rowallexpress = self._col_to_row_sort(allexpress, titlelen, valuelen) #将表达式由列式转为行式显示
        rowavgs = self._calc_row_avg(rowallexpress, titlelen) #根据行式表达式计算最终结果
        return rowavgs

    def GenResultToExcel(self,calcname, callid_and_time, result_avg):
        if self.reportexcel is None:
            excel = ThrExcel.ThrExcel(visiable=1,newfile=True)
            self.reportexcel = excel
        else:
            excel = self.reportexcel
        sht = excel.createSheet(calcname,False)
        i = 1
        for i in range(3):
            j = 1
            for ct,r in zip(callid_and_time,result_avg):
                sht.setRowData(j,[ct[0],ct[1],r])
                j += 1

    #对每条数据生成表达式用于计算
    def _gen_express(self, optlen, titlelen, valuelen):
        allexpress = []
        for tl in range(titlelen):
            tempexpress = self.express
            nowtitleexpress = []
            for ol in range(optlen):
                nowopt = self.expressopt[ol]
                try:
                    nowtitle = self.alltitles[nowopt][tl]
                except:
                    nowtitle = self.alltitles[nowopt][0] #如果取值出错，说明有些是单列的Title，则取第一个列名，不考滤各类列值不一致的情况。
                    #此类情况过于复杂，可手工调整列名来操作
                tempexpress = tempexpress.replace(nowopt, nowtitle)
                nowtitleexpress.append(nowtitle)
            nowtempexpress = tempexpress
            for vl in range(valuelen):
                tempexpress = nowtempexpress
                for nte in nowtitleexpress:
                    tempexpress = tempexpress.replace(nte, str(self.table[nte][vl]))
                allexpress.append(tempexpress)
        return allexpress


    #根据行排表达式计算平均值
    def _calc_row_avg(self, rowallexpress, titlelen):
        avgresult = []
        avgvs = []
        for ex, i in zip(rowallexpress, range(len(rowallexpress))):
            if (i % titlelen == 0 and i != 0):
                avgresult.append(AVG(avgvs))
                avgvs.clear()
            try:
                result = eval(ex)
            except:
                result = None
            avgvs.append(result)
            if i == len(rowallexpress) - 1:
                avgresult.append(AVG(avgvs))
        return avgresult

    # 表达式的由列排转为行排，为计算平均值做准备
    def _col_to_row_sort(self, allexpress, titlelen, valuelen):
        m = 0
        row = 0
        rowallexpress = list.copy(allexpress)
        for i in range(titlelen):
            rowallexpress[i] = allexpress[m * valuelen]
            m += 1
        m = 0
        row += 1
        for j in range(1, valuelen):
            m = 0
            for k in range(titlelen):
                rowallexpress[row * titlelen + k] = allexpress[m * valuelen + j]
                m += 1
            row += 1
        return rowallexpress

    def getColumnMapping(self, tableheadreg):
        try:
            self.getColumnValues(tableheadreg)
            self.getIDColumns()
        except Exception as e:
            raise e

    def getColumnValues(self, tableheadreg):
        for reg in tableheadreg:
            tablecolumn = self.sheet.getTableColumnCells(reg, getValue=True)
            self.table.update(tablecolumn)
        self.end = self.sheet.getEndingRowsCount(1)
        print(self.table)

    def getIDColumns(self):
        t = {}
        callID = self.sheet.getTableColumnCells(r'.*CallID',getValue=True)
        time = self.sheet.getTableColumnCells(r'TIME',getValue=True)

        for c,t in zip(list(callID.values())[0],list(time.values())[0]): #只有一列，
            c = str(c).split('.')[0] #//todo://暂时去掉.,应该使用Format方法来设置Column
            self.callanditem.append([c,t])

def AVG(args):
    i = 0
    sum =0
    for a in args:
        if a is not None:
            sum += a
            i += 1
    return sum / i


def calcMethord(ce):
    # ce.openExcel('C:\\Users\\Administrator\\Desktop\\mytest.csv','mytest') 固定用值,仅作测试
    ip = input("请输入表达式")
    # ip = '(Ulsmgres + RRSNkuks.avg)/Lupussh' 固定用值，仅作测试
    ce.copyExpressAnlys(ip.strip())
    print("请输入表达式对应表的关系，使用正则表达式关系，请注意转义符号的使用")
    # ce.titlereg = ['L2_USERCHR_SCH_INFO\(0\).DRB_512MS\(\d\).PUSCH_DRB_512MS\(0\).ulPuschMcsSum',\
    #            'L2_USERCHR_SCH_INFO\(0\).DRB_512MS\(\d\).ulPuschSchSum',\
    #            'L2_USERCHR_VOLTE_SCH_INFO\(0\).DRB_512MS\(0\).PUSCH_DRB_512MS\(\d\).ulPuschMcsSum']
    # for op,e in zip(ce.titlereg, ce.expressopt):
    #     ce.eft[e] = op
    for exo in ce.expressopt:
        ip = input("请输入" + exo + "对应用的列表")
        ce.eft[exo] = ip.strip()
        ce.titlereg.append(ip)
    ce.getColumnMapping(ce.titlereg)
    ce.expressMapping2()
    allavgexp = ce.runexpress2()
    ip = input('请输入此表达式生成报告的结果名称[表格名称]')
    ce.GenResultToExcel(ip.strip(),ce.callanditem, allavgexp)
    # del ce


def fileOpen():
    en = input('请输入cvs路径名称')
    sn = input('请输入cvs表格名称，如果不输入则是文件名称')
    ce = CopyExpress()
    if len(sys.argv) == 1:
        tn = re.split(r'\\|\.', en)
        sn = tn[-2]
        print('表格名称是', sn)
    ce.openExcel(en, sn)
    return ce


if __name__ == "__main__":
    try:
        ce = None
        ce = fileOpen()
        while True:
            calcMethord(ce)
            ce.clearUp()
            ip = input('是否要继续计算其它公式,请输入y，否则任意键')
            if ip == 'y' or ip == 'Y':
                continue
            else:
                break
        ce.excel.close()
        ce.reportexcel.close()
    except BaseException as e:
        print('出错，请查看出错类型',e)




