
import os
import time

print('数据存入表格成功。。。。。。。。。。')
f = os.popen(r"D:\Users\Administrator\桌面\买菜销量\买菜销量.xls", "r")
print('打开表')
time.sleep(10)
print('等10秒')
os.system('taskkill /f /im wps.exe')
print('关闭表')
