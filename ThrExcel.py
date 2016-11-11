#!/usr/bin/python
# -*- coding:utf-8 -*-
#@Author DOUMEKI

import win32com.client
import win32clipboard
import re
import pythoncom


class ThrExcel:
    # fileName：要打开的文档名称
    # visiable: Excel是否可见
    # newfile: 是否是新开进程方式打开
    # multithread: 是否多线程操作Excel (多线程访问Sheet) 注：指多个线程同时操作Excel文档,为True为多线程，为False为当前线程
    # subthread: 是否是子线程 注：指打开Excel的线程是否为程序的子线程
    def __init__(self, filename=None, visiable=0, newfile=True, multithread=False, subthread=False):
        '''

        :param filename: 要打开的文档名称
        :param visiable: Excel是否可见
        :param newfile: 是否是新开进程方式打开
        :param multithread: 是否多线程操作Excel (多线程访问Sheet) 注：指多个线程同时操作Excel文档,为True为多线程，为False为当前线程
        :param subthread: 是否是子线程 注：指打开Excel的线程是否为程序的子线程
        '''
        self._setup(filename, multithread, newfile, visiable, subthread)
        #self.subthread = subthread #是否是子线程打开程序，关系到是否要UnMarsh一次stream
        self.multithread = multithread
        self.subthread = subthread

    def _setup(self, filename, multithread, newfile, visiable, subthread):
        # self.xlApp = win32com.client.Dispatch('Excel.Application') #使用这个，不能单独设置visiable
        self._myStream = None
        if multithread:
            self.multiThreadOperationInit()  # com组件多线程操作,在其子线程上也要运行该操作
        # elif (subthread and not multithread) or (not subthread):
        else:
            pythoncom.CoInitialize()  # COM组件单线程操作，注：这个可以不是主线程。
        if newfile:
            self.xlApp = win32com.client.DispatchEx("Excel.Application")  # 使用这个后，粘贴Excel会有提示，原因不明
        else:
            self.xlApp = win32com.client.dynamic.Dispatch('Excel.Application')  # 使用现有的Excel程序打开，如果没有才新建
        if multithread:
            self._myStream = pythoncom.CreateStreamOnHGlobal()
            pythoncom.CoMarshalInterface(self._myStream,
                                         pythoncom.IID_IDispatch,
                                         self.xlApp._oleobj_,
                                         pythoncom.MSHCTX_LOCAL,
                                         pythoncom.MSHLFLAGS_TABLESTRONG)
        elif not multithread and subthread:
            self._myStream = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch,
                                                                             self.xlApp._oleobj_)
        if filename:
            self.filename = filename
            self._xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self._xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''
        self.xlApp.Visible = visiable
        self.xlApp.DisplayAlerts = visiable

    # 关闭Excel，
    def close(self, finishExcelInitSubThreading=True):
        '''
        关闭Excel
        :param finishExcelInitSubThreading: 是否结束打开Excel的线程的初始信息
        :return:
        '''
        try:
            self._excelexit(self.xlApp)
            #if self.subthread: #如果是用子线程打开，需要release Marshal Data
            if self.multithread or self.subthread:
                self.multiThreadReleaseThreadData()
        except: #当前一步出错时，一般是多线程操作
            try:
                excel = self.GetExcelThroughMultiThreads()
                self.multiThreadReleaseThreadData()
                self._excelexit(excel)
            except Exception as e:
                raise Exception('关闭异常')
        if finishExcelInitSubThreading:
            self.subThreadingOperationFinish()
            # 一个线程上只需要运行一次就行，不管开几个Excel，此API不存在此问题


    def _excelexit(self,excel):
        excel.Workbooks(1).Close(SaveChanges= 0)
        excel.Quit()  # 大小写都可以

    # 根据sheetName查找sheet对象
    def getSheet(self, sheetName):
        '''
        根据sheetName查找sheet对象
        :param sheetName: sheet 的名称
        :return: 该sheet的内部对象，非COM对象
        '''
        sht = ThrSheet(sheetName, self.xlApp, stream=self._myStream)
        self._activeExcelSheet(sht._sheet)
        return sht

    def multiThreadOperationInit(self):
        '''
        多线程开始时初始化
        :return:
        '''
        pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)

    #子线程结束时调用，若要在此线程上关闭Excel时则不能调用（由Close方法的默认参数设置来调用）
    def subThreadingOperationFinish(self):
        '''
        在非主线程（子线程）结束时手动调用。
        若要在此线程上关闭Excel时则不能调用（由Close方法的默认参数设置来调用）
        :return:
        '''
        pythoncom.CoUninitialize(pythoncom.COINIT_MULTITHREADED)

    # 多线程的单个线程操作完成时释放MarshData，注意使用锁操作机制保证取得的对象唯一性
    def multiThreadReleaseThreadData(self):
        '''
        多线程的单个线程操作完成时释放Marshalled Interface，注意使用锁操作机制保证取得的对象唯一性
        :return:
        '''
        self._myStream.Seek(0, 0)
        pythoncom.CoReleaseMarshalData(self._myStream)


    def getSheetWithReg(self, sheetName):
        '''
        根据正则表达式获取sheet内部对象
        :param sheetName: sheet的名称
        :return: 该sheet的内部对象，非COM对象
        '''
        sht = ThrSheet(self._searchSpecialSheetName(sheetName), self.xlApp, self._xlBook, stream=self._myStream)
        return sht

    def _getsheetobj(self, sheetName, excelobj=None):
        '''
        根据sheet名称获取sheet的COM对象
        :param sheetName: sheet的名称
        :param excelobj: excel的COM对象
        :return: 该Sheet的COM对象
        '''
        # self._searchSpecialSheetName(sheetName,excelobj)

        if excelobj is None:
            sht = self.xlApp.Worksheets[sheetName]
        else:
            sht = excelobj.Worksheets[sheetName]
        self._activeExcelSheet(sht)
        return sht

    # 激活当前Excel的一个sheet
    def _activeExcelSheet(self, sheet=None):
        if sheet:
            sheet.Activate()
        else:
            self.sht.Activate()

    # 查找特字名字的sheet
    def _searchSpecialSheetName(self, sheetName, _excelobj=None):
        cmp = re.compile(sheetName)
        if _excelobj is not None:
            excel = _excelobj.Worksheets
        else:
            excel = self.xlApp.Worksheets
        for s in excel:
            if cmp.search(s.Name):  # 找得到Monkey 字样的sheet，这样便于模糊查找，如果以后出现Monkey Maki也可以用Monkey来对应
                return s.Name
        return None


    def setVisiable(self, visible):
        self.xlApp.Visible = visible  # 设定Excel是否可见

    # 多线程操作时取得Excel对象，注意使用锁操作机制保证取得的对象唯一性
    def GetExcelThroughMultiThreads(self):
        '''
        线程操作时取得ExcelCOM对象，注意使用锁操作机制保证取得的对象唯一性
        :return:
        '''
        self._myStream.Seek(0, 0)
        unMarshaledInterface = pythoncom.CoUnmarshalInterface(self._myStream, pythoncom.IID_IDispatch)
        xlApp = win32com.client.dynamic.Dispatch(unMarshaledInterface)
        return xlApp

    # 多线程操作时取得sheet对象,注意使用销操作机制保证取得的对象唯一性
    def GetSheetThroughMultiThreads(self, sheetName):

        '''
        在多线程multithread方式操作时，根据sheet的名称获取Sheet的内部对象，非COM对象
        :param sheetName: Sheet的名称
        :return: sheet的内部对象
        '''
        _excelobj = self.GetExcelThroughMultiThreads()
        return ThrSheet(sheetName, _excelobj, self._myStream)




class ThrSheet():
    def __init__(self,sheetname,excelobj,stream = None):
        if sheetname is None or excelobj is None : #todo://Raise的条件
            raise BaseException('sheet对象为空')
        else:
            try:
                self._sheet = excelobj.Worksheets[sheetname]
                self._excelobj = excelobj  # xlApp
                self._workbooksobj = excelobj.Workbooks(1)
                self._stream = stream
            except Exception:
                raise BaseException("查找Sheet对象出错")

    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self._workbooksobj.SaveAs(newfilename)
        else:
            self._workbooksobj.Save()



    # 获取一个Cell
    # @cell 的格式： A1, B3
    def getOneCellByRowNameColumnName(self, cell):
        "Get value of one cell"
        c = self._cellsplit(cell)
        return self._sheet.Cells(c[1], c[0])

    # 获取一个Cell的值
    # @cell 的格式： A1, B3
    def getOneCellValueByRowNameColumnName(self, cell):
        return self.getOneCellByRowNameColumnName(cell).Value


    # # 根据一个Cell对像取得Value值
    # def getOneCellValueByCellObject(self, cell):
    #     return cell.Value
    #
    # # 根据行数，列数取得一个Cell对象
    # def getOneCellByRowColumnIndex(self, row, column):
    #     return self._sheet.Cells(row, column)

    # 获取行对象
    # @index_row: 行数　整数型
    def getRowObjectByRowIndex(self, index_row):
        return self._sheet.Rows(index_row)

    # 获取一行已使用的Cell，已使用的意思是该cell有值或该cell的背景色。要求连续
    # @index_row：目标行，
    # @start_column: 默认从列１开始计算
    def getUsedRowCellsByRowIndex(self, index_row, start_column=1):
        indexColum = self._usedRange(index_row)
        if indexColum == 0:
            return None
        return self.getRange(index_row, start_column, index_row, indexColum)

    def _usedRange(self, index_row):
        allCells = self.getRowCellsByRowIndex(index_row)
        indexColum = 0
        for c in allCells:
            if not c.Value and c.Interior.Color == 16777215:  # 如果这个cell没有数据或是没有背景色，就认为是没有使用的。
                break
            elif c.Value is not None:
                try:  # 之所以用try,是因为1900年会被识别成1899年，unicode转换不了。当一量是1899年，说明有数据，则不用break
                    t = unicode(c.Value).strip()
                    if t == '':
                        indexColum += 1  # 如果存在数据‘’，则当前列是有数据列，则要+1
                        break
                except:
                    pass
            indexColum += 1
        return indexColum

    def isEmptyRow(self, row_index):
        if self.getUsedRowCellsByRowIndex(row_index):
            return False
        else:
            return True

    # 获取一行的所有Cell
    def getRowCellsByRowIndex(self, index_row):
        return self._sheet.Rows(index_row).Cells

    # 获取一个wk的usedRange
    def getUsedRangeInWorkSheet(self, sheet=None):
        if sheet:
            return sheet.UsedRange #此处取值不准确，原API的原因，暂不修正
        else:
            return self._sheet.UsedRange

    # 获取一个Range,根据star,end　的行数
    # @start_index_row: 开始行　包含　
    # @end_index_row 结束行　包含
    def getRowObjectByStartEndIndex(self, start_index_row, end_index_row):
        stringSelectRow = str(start_index_row) + ":" + str(end_index_row)
        return self._sheet.Rows(stringSelectRow)

    # 复制一行到另一行，包括值和格式,在同一个表中
    # @sour_index: 复制这个行
    # @dest_index: 粘贴到这个行
    # 要求要在同一个sheet中
    def copyRowFromRow(self, sour_index, dest_index):
        s = self._sheet.Rows(sour_index)
        d = self._sheet.Rows(dest_index)
        d.Rows.Insert(CopyOrigin=s.Copy())
        self.clearRowValue(dest_index)
        self.clearClipboard()
        return self.getRowObjectByRowIndex(dest_index)

    # 复制当前文档的一行到另一个文档的行
    # @otherExcel: 另一个文档对象
    # @sour_index: 源行
    # @dest_index: 另一个文档的目标行
    def copyRowToWithInsert(self, otherExcel, sour_index, dest_index):
        s = self._sheet.Rows(sour_index)
        d = otherExcel.sht.Rows(dest_index)
        d.Rows.Insert(CopyOrigin=s.Copy())

        # 清除一行的所有Value

    # @index_row: 要清除的目标行
    def clearRowValue(self, index_row):
        for n in self.getUsedRowCellsByRowIndex(index_row):
            n.Value = None

    # 清空目标sheet的数所属性，包括值，格式
    def clearWorkSheet(self, sheet):
        sheet.UsedRange.Clear()

    # 清空剪切板
    def clearClipboard(self):
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()

    # 设置一个cell的值
    # cell格式：A1，B1
    # value: 值
    def setCellValue(self, cell, value):
        "set value of one cell"
        self.getOneCellByRowNameColumnName(cell).Value = value

    # 对于Cell格式的简单判断
    def _cellsplit(self, cell):
        if len(cell) < 2:
            raise "cell given is wrong"
        cell = list(cell)
        t = []
        t.append(cell[0])
        t.append(''.join(cell[1:]))
        return t

    # 根据*行*列到*行*列得到一个Range
    # Ｎ行　Ｎ列　到　Ｍ行，Ｍ列
    def getRange(self, row1, col1, row2, col2):
        return self._sheet.Range(self._sheet.Cells(row1, col1), self._sheet.Cells(row2, col2))

    def addPicture(self, sheet, pictureName, Left, Top, Width, Height):
        "Insert a picture in sheet"
        self._sheet.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)

    def copySheet(self, before):
        "copy sheet"
        shts = self._workbooksobj.Worksheets
        shts(1).Copy(None, shts(1))

        # 得到最后一个组合的起始行数

    def getLastGroupedRowLineNumber(self):
        totalnumber = self.getUsedMaxRowIndex()
        for r in range(totalnumber, 1, -1):  # 倒着查找outline = 1的，这样不用编历整个excel,提高效率
            now = self.getRowObjectByRowIndex(r)
            if now.OutlineLevel == 1.0 and \
                            r < totalnumber and \
                    not self.isEmptyRow(r):
                return r

    # 得到已使用行的最后一行的行数
    def getUsedMaxRowIndex(self):
        return self.getUsedRangeInWorkSheet().Rows.Count

    # 取得某列的所有列单元
    def getColumnCellsByColumnIndex(self, columnIndex):
        endrow = self.getUsedMaxRowIndex()
        cells = []
        for i in range(1, endrow + 1, 1):
            cells.append(self.getOneCellByRowColumnIndex(i, columnIndex))
        return cells

    # 取得某列的所有列单元的值
    def getColumnCellsValueByColumnIndex(self, columnIndex):
        cellsValue = []
        for i in self.getColumnCellsByColumnIndex(columnIndex):
            cellsValue.append(i.Value)
        return cellsValue

    # 根据目标行获取一个Cell
    # @Row: 目标行
    # @index_column 代表这个行的第几列
    def getCellByGivenRow(self, row, index_column):
        return row.Cells(index_column)

    # 根据行号及列号获取一个Cell
    # @index_row: 行号
    # @index_column 列号
    def getOneCellByGivenRowColumnIndex(self, index_row, index_column):
        return self.getRowCellsByRowIndex(index_row)(index_column)




