
#coding=utf-8
import xlrd
import os
import makedir
import excel


def main():
    #创建一个目标输出文件夹
    makedir.mkdir('d:/app_install_speed_test')

    #确认共有几个log
    logs = os.listdir('d:/logs')
    if logs[0] == ' ':
        print '请确保桌面上log文件夹里有目标log文件，如：act_dumpstate1'
    else:
        times = len(logs)

    #先从log文件中将需要的信息筛选出来并保存到/Users/Roger/Desktop/app_install_speed_test文件夹
    excel.getTXTfile(logs, times)
    #30是指测试了30个app
    test_result = excel.getInstallTime(times,30)
    #将结果保存到excel中,第一个参数是测试次数，第二个参数是测试apk数
    excel.saveInExcel(times,30,test_result)



if __name__ == "__main__":
    main()