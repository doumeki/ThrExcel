#!/usr/bin/python
# -*- coding:utf-8 -*-
#@Author DOUMEKI
import ThrExcel,threading,time
from unittest import TestCase
#此Demo分为：
#1. 单线程打开并操作Excel，并在单线程上结束Excel
#2. 子线程打开并操作Excel，并在子/主线程上结束Excel
#3. 主线程打开Excel，多线程操作Excel，并在主线程上结束Excel
#4. 子线程打开Excel，多线程操作Excel，并在子/主线程上结束Excel
class DemoThrExcel(TestCase):
    def setUp(self):
        self.path ="C:\\Users\\Administrator\\Desktop\\黄金链资料表最新(1).xls"
        self.excel = None
        self.locker1 = threading.Condition()
        self.locker2 = threading.Condition()

    def tearDown(self):
        pass

    #所有的操作在一个线程中完成
    def test_one_thread_in_all(self,sub = False,closed = True):
        excel = ThrExcel.ThrExcel(filename=self.path,subthread=sub)
        sheet  = excel.getSheet('买一个送一')
        columns = sheet.getColumnCellsValueByColumnIndex(1)
        print (sheet.getUsedMaxRowIndex())#原API产生行数大于实际值的问题，请自行修正
        print(columns)
        print(len(columns))
        if closed:
            excel.close()
        else:
            self.excel = excel

    #在子线程中打开Excel,但操作仍是在这个子线程中进行,并在子线程中关闭
    def test_in_sub_thread_no_close(self):
        t = threading.Thread(target=self._test_in_sub_thread_func_closed_in_subthread)
        t.start()

    # 在子线程中打开Excel,但操作仍是在这个子线程中进行,并在主线程中关闭
    def test_in_sub_thread_with_close(self):
        t = threading.Thread(target=self._test_in_sub_thread_func_closed_in_mainthread)
        t.start()
        time.sleep(10) #自行检测读取完成事件
        self.excel.close()

    #线程FUNC
    def _test_in_sub_thread_func_closed_in_subthread(self):
        self.test_one_thread_in_all(True)
    #线程FUNC
    def _test_in_sub_thread_func_closed_in_mainthread(self):
        self.test_one_thread_in_all(True,False)

    #主线程打开Excel，多个线程同时操作Excel，并在主线程中关闭
    def test_open_Excel_in_mainthread_operation_in_multi_thread(self,close = True):
        excel = ThrExcel.ThrExcel(self.path,multithread=True,subthread=False)
        self._test_multi_thread_operation(close, excel)

    #多线程操作FUNC
    def _test_multi_thread_operation(self, close, excel):
        for i in range(2):
            t = threading.Thread(target=self._test_open_Excel_in_mainthread_opration_in_multi_thread_func,
                                 args=(excel, i))
            t.start()
        if close:
            time.sleep(10)  # //自行检测读取完成事件
            excel.close()
        else:
            self.excel = excel

    #线程FUNC
    def _test_open_Excel_in_mainthread_opration_in_multi_thread_func(self,excel,i):
        excel.multiThreadOperationInit() #多线程是初始化
        # sheet = excel.GetSheetThroughMultiThreads("买一个送一") #多线程取得Sheet对象,不使用锁时报错
        sheet = self._lockSheet(self.locker1,excel.GetSheetThroughMultiThreads,"买一个送一") #多线程取得Sheet对象,要使用锁来操作，否则会出错
        columns = sheet.getColumnCellsValueByColumnIndex(i+1)
        print (columns)
        #excel.multiThreadReleaseThreadData() #释放对象,不使用锁时报错？
        self._lockSheet(self.locker1,excel.multiThreadReleaseThreadData)#释放对象,要使用锁来操作，否则会报错

    def _lockSheet(self,locker,func,*args):
        #locker = self.locker1
        if locker.acquire():
            r = func(*args)
            locker.notify()
            locker.release()
        else:
            locker.wait()
            r = func(*args)
            locker.notify()
            locker.release()
        return r

    # 子程打开Excel，多个线程同时操作Excel，并在主线程中关闭
    def test_open_excel_subthread_operation_multi_thread_with_no_closed_in_main_thread(self):
        t1 = time.clock()
        print("this is main thread")
        t = threading.Thread(target=self._subthread_func_with_Main_thread_close)
        t.start()
        t2 = time.clock()
        sp = t2 - t1
        print(sp, "main finished")
        time.sleep(10)
        self.excel.close()

    # 子程打开Excel，多子线程同时操作Excel，并在子线程中关闭
    def test_open_excel_subthread_operation_multi_thread_with_closed_in_sub_thread(self):
        t1 = time.clock()
        print("this is main thread")
        t = threading.Thread(target=self._subthread_func_within_sub_thread_close)
        t.start()
        t2 = time.clock()
        sp = t2-t1
        print(sp,"main finished")

    #线程FUN在子线程中关闭
    def _subthread_func_within_sub_thread_close(self):
        excel = ThrExcel.ThrExcel(self.path, multithread=True, subthread=True)
        self._test_multi_thread_operation(True,excel)
    #线程FUN在主线程中关闭
    def _subthread_func_with_Main_thread_close(self):
        excel = ThrExcel.ThrExcel(self.path, multithread=True, subthread=True)
        self._test_multi_thread_operation(False,excel)
