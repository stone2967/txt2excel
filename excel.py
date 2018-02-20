#coding=utf-8
import os
import xlwt
import time_calculate
import datetime

def getInstallTime(times,app_num):
    '''
    :param times: 测试了几轮
    :param app_num:测试了几个app
    :return:
    '''
    #建立存放每次测试结果的字典
    test_result = {}
    for time in range(times):
        # 打开log文件
        startTimeFile = 'd:/app_install_speed_test' + '/' + 'log' + str(time) + 'starttime'
        endTimeFile = 'd:/app_install_speed_test' + '/' + 'log' + str(time) + 'endtime'
        f1 = open(startTimeFile)
        f2 = open(endTimeFile)

        #开始安装时间和安装结束时间 的局部变量
        start_time = '1991-02-04 00:00:00'
        end_time = '1991-02-04 00:00:00'
        time_delta = {}
        # 读时间
        startTimeLines = f1.readlines()
        endTimeLines = f2.readlines()
        for num in range(app_num):
            start_time = '2018-02-10 ' + startTimeLines[num].split()[1].split('.')[0]
            end_time = '2018-02-10 ' + endTimeLines[num].split()[1].split('.')[0]
            #不足1s的部分先计算
            microSecondDelta = 0.001*(int(endTimeLines[num].split()[1].split('.')[1]) - int(startTimeLines[num].split()[1].split('.')[1]))
            #将时间转化为秒数
            start_second = time_calculate.ISOString2Time(start_time)
            end_second = time_calculate.ISOString2Time(end_time)
            #计算安装app耗费的时间
            time_delta[str(num+1)+'.apk'] = end_second - start_second + microSecondDelta

        test_result[str(time)] = time_delta
        f1.close()
        f2.close()

    return test_result



def getTXTfile(logs,times):
    for time in range(times):
        # 打开log文件
        f1 = open('d:/logs' + '/' + logs[time])
        #创建目标文件
        f2 = open('d:/app_install_speed_test/log' +  str(time) + 'starttime','w+')
        f3 = open('d:/app_install_speed_test/log' +  str(time) + 'endtime','w+')
        lines = f1.readlines()
        for line in lines:
            if "install /sdcard/app_enter_time" in line:
                f2.write(line)
            if "PackageManager: result of install" in line:
                f3.write(line)
        f1.close()
        f2.close()

def saveInExcel(times,app_num,test_result):
    # 创建一个Workbook对象，这就相当于创建了一个Excel文件
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    '''
    Workbook类初始化时有encoding和style_compression参数
    encoding:设置字符编码，一般要这样设置：w = Workbook(encoding='utf-8')，就可以在excel中输出中文了。
    默认是ascii。当然要记得在文件头部添加：
    #!/usr/bin/env python
    # -*- coding: utf-8 -*-
    style_compression:表示是否压缩，不常用。
    '''
    # 创建一个sheet对象，一个sheet对象对应Excel文件中的一张表格。
    # 在电脑桌面右键新建一个Excel文件，其中就包含sheet1，sheet2，sheet3三张表
    #sheet = book.add_sheet('result_'+ datetime.datetime.now().strftime("%Y/%m/%d_%H_%M_%S"), cell_overwrite_ok=True)
    sheet = book.add_sheet('test1', cell_overwrite_ok=True)
    # 其中的test是这张表的名字,cell_overwrite_ok，表示是否可以覆盖单元格，其实是Worksheet实例化的一个参数，默认值是False
    # 向表test中添加数据
    for num in range(app_num):
        app_name = str(num+1) + '.apk'
        sheet.write(0, num,app_name)  # 其中的'0-行, 0-列'指定表中的单元，'app_name'是向该单元写入的内容
        for time in range(times):
            sheet.write(time+1, num,test_result[str(time)][app_name] )

    # 最后，将以上操作保存到指定的Excel文件中
    #book.save(r'e:\test1.xls')  # 在字符串前加r，声明为raw字符串，这样就不会处理其中的转义了。否则，可能会报错
    filename = 'Result_'+ datetime.datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    book.save('d:/app_install_speed_test/' + filename + '.xls')