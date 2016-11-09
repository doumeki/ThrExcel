# ThrExce
#利用COM组件，多线程操作Excel的API
#此API未包含所有的操作，大部分为读取操作，可通过<ThrExcel实例>.xlApp自行添加功能.
#具体使用方法查看DemoThrExcel.py
#在Test_XXX方法上运行每个Test查看结果
#---------------------Demo 列出了以下几种场景-------------------------------

#1. 单线程打开并操作Excel，并在单线程上结束Excel
#2. 子线程打开并操作Excel，并在子/主线程上结束Excel
#3. 主线程打开Excel，多线程操作Excel，并在主线程上结束Excel
#4. 子线程打开Excel，多线程操作Excel，并在子/主线程上结束Excel


