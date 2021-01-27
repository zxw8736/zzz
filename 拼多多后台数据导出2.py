
from time import sleep
import requests
import sys
import datetime
import xlwt#写excel
import xlrd
import os
import time
#导入模块
from xlutils.copy import copy        #导入copy模块


class xie_xls():
    def __init__(self,name,biandan='sheet1'):
        self.name=name
        self.biandan=biandan

    # 创建表
    def chuangbiao(self,tou):
        book=xlwt.Workbook()#创建一个工作簿
        sheet=book.add_sheet(self.biandan)#创建一个表名
        # 第一个参数是行，第二个参数是列，都从0开始计算
        # tou = ['总积压', '通过', '拒绝', '今日积压']

        for i in range(0, len(tou)):
            sheet.write(0, i, tou[i])

        book.save('{}'.format(self.name))#保存一个excel文件
        print('表格创建成功。。。。。。。')


    # 修改表02
    def xiubiao02(self,data,tou,biao01=1):
        rb = xlrd.open_workbook('{}'.format(self.name),formatting_info=True)  # 打开weng.xls文件   保留表格格式 formatting_info=True
        wb = copy(rb)  # 利用xlutils.copy下的copy函数复制

        dan_name=rb.sheet_names()
        print('表单：',dan_name)

        if self.biandan not in dan_name:
            ws = wb.add_sheet(self.biandan)  # 创建一个表名
            hang=0
            pass
        else:
            sh = rb.sheet_by_name(self.biandan)
            hang = sh.nrows  # 读取有几行，从1开始记
            for i01 in range(0,len(dan_name)):
                if dan_name[i01]==self.biandan:
                    ws = wb.get_sheet(i01)  # 获取表单0

        # 设置背景为 red
        # style = xlwt.easyxf('pattern:pattern solid,fore_colour red')
        # ws.write(0,0,'',style)

        if hang==0:
            print('写表头。。。。。')
            for i01 in range(0,len(tou)):
                ws.write(hang,i01,[hang+1]+tou[i01])  # 改变（0,0）的值

            hang=1

        for i01 in range(0,len(data)):
            ws.write(hang,i01,[hang+1]+data[i01])  # 改变（0,0）的值

        wb.save('{}'.format(self.name))  # 保存文件

        if biao01==1:
            print('数据存入表格成功。。。。。。。。。。')

            os.popen(r"D:\Users\Administrator\桌面\买菜销量\买菜销量.xls", "r")
            print('打开表')
            time.sleep(10)
            print('等10秒')
            os.system('taskkill /f /im wps.exe')
            print('关闭表')


    # 读取表
    def dubiao(self,xuan=-1):
        fname = "{}".format(self.name)
        bk = xlrd.open_workbook(fname)
        # shxrange = range(bk.nsheets)
        name01=bk.sheet_names()
        print('表单名：',name01)

        if type(xuan)==int:
            print('读取表单名：', name01[xuan])
            sh = bk.sheet_by_name(name01[xuan])
        else:
            sh = bk.sheet_by_name(xuan)

        # sh = bk.sheet_by_name(self.biandan)

        nrows = sh.nrows  # 读取有几行，从1开始记
        ncols = sh.ncols  # 读取最后列的位置，从1开始记
        print("nrows %d, ncols %d" % (nrows, ncols))

        # cell_value = sh.cell_value(0, 0)  # 读取（0,0）位置的内容
        # print(cell_value)

        row_list = []
        for i in range(0, nrows):
            row_data = sh.row_values(i)  # 读取一行的内容，返回一个列表
            # print(row_data)
            row_list.append(row_data)
        # print(row_list)

        print('读取表格成功。。。。。。。。。。。。。')

        return row_list


def qing_post(url):
    while 1:
        try:
            headers = {
                'accept': '''application/json''',
                'accept-encoding': '''gzip, deflate, br''',
                'accept-language': '''zh-CN,zh;q=0.9''',
                # 'anti-content': '''0aoAfxvUDiQYq9EVF88Jf2EB2UTHV4KAOPymeoYZdgwkIZP6Ljj2xM7ZCdbHunkpASY_A-y4xmaHR97JJzw9AMprxLQcLaWKoX9h9vuAXSY01yBcuQPpVSr2DRVfhiyDOc1Azvefen-fDurIR_e-_1Od1zxgPnvIpfw0bKNvwv6_NvmyXj9ZbDfoFx8eWBBHykg22XDjLDDlYCAlMwPAJmKtYvSan_CvoEFw5oYP-GTx0eZyJCZCqSNjmwD6OSNvChjCEAg9Q9EscHcl6h_2AW8J9zdJ5UVhcW3j9-BMcmJM5loQwhL-s97AJ_4mLVlci0wQ7CfxoGa70CjudZ9XiG3TAbMTdExp9GBkdIN90x9c3aqy6sZlVL4caYggQxgIWCad_uesXok-V5ys8Stuz5TuPbHDlG7k0sDAlJsBcUvagtUENFkTrm86TRyaiRMUiOGvXZnTlNgOFg73lj3v1zoWBfZIw0Q0BxdHfqvr-rIZwgaFjZZVYa1efHE7Xf5YyNzXvhZzBy8eD1HbHsxR-yjmRnaQGILh1Tqf13I3umUQxDk8IWRENubRNihbIqX2-7ezPvkx_YJCNdACBATLeRzYkWQbnkxRsgxPyAtJnj7kGTtoqNPvtAv''',
                'content-length': '''38''',
                'content-type': '''application/json;charset=UTF-8''',
                'cookie': cookie01,
                'origin': '''https://mai.pinduoduo.com''',
                'referer': '''https://mai.pinduoduo.com/mobile-grocer-supply/orders-detail.html?date=20210114&wid=3959&wName=%E9%9D%92%E5%B2%9B1%E4%BB%93&isEnd=1&areaId=22''',
                'user-agent': '''Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.81 Safari/537.36 SE 2.X MetaSr 1.0''',
            }

            # 忽略不验证证书出现的警告,proxies=http
            requests.packages.urllib3.disable_warnings()

            # r = requests.get(url, headers=headers, verify=False,proxies=http,timeout=10)

            nowtime = datetime.datetime.now().strftime('%Y%m%d')  # 现在
            data='{"date":"%s","warehouseId":3959}'%nowtime
            r = requests.post(url,headers=headers,verify=False,timeout=10,data=data)
            r.encoding='utf-8'

            return r
        except Exception as e:
            print('请求错误：',e)
            sleep(1)


print('程序开启。。。。。。。。。。。。。。。。')
cookie01='api_uid=rBQR6WACx/u3HDsxdk3nAg==; _nano_fp=XpEaX0mbnp9onpTan9_45j2vSPNqFr_ll6j3ugPB; _bee=ewuZSWfRaBbK6FMJP7dm2FWtAOHCde9r; _f77=b37320fd-eb21-49f1-a796-6ea1aa81f8bb; _a42=7812b02a-84d5-488c-8b38-e2533217be9f; rckk=ewuZSWfRaBbK6FMJP7dm2FWtAOHCde9r; ru1k=b37320fd-eb21-49f1-a796-6ea1aa81f8bb; ru2k=7812b02a-84d5-488c-8b38-e2533217be9f; x-visit-time=1611285928323; JSESSIONID=DBF68839BB0A5955E6E17C712F6F2C76; mms_b84d1838=120,1202,1203,1204,1205,3423; PASS_ID=1-N4Vl1cnkkNtv+XzlwBFHwZ2VhBY70dzK3N5xqlt4gp/Z4b1JGd/TAF8cECl/01f6njL1CnU1LCbjTWeZUTZV8Q_912018337_85453916'
bian_name = datetime.datetime.now().strftime('%m-%d')  # 现在
yue01=xie_xls('买菜销量.xls',bian_name)
try:
    yue01.dubiao(0)
except:
    print('文件不存在，新建表格。。。。。')
    yue01.chuangbiao([])


try:
    url01 = 'https://mms.pinduoduo.com/patronus-mms/order/daily/statisticList'
    data01 = qing_post(url01)
    data02 = data01.json()
    print('源码：',data02)
    data03 = data02['result']['orderList']



except:
    s = sys.exc_info()
    print("错误第{}行,详情：【'{}' 】".format(s[2].tb_lineno, s[1]).replace('\n', ''))

    while 1:
        input('程序出错，请检查后重新运行程序。。。。。。。。。。。。。。。。。。。')


sj01=60*1
while 1:
    nowtime = datetime.datetime.now().strftime('%Y/%m/%d %H:%M')  # 现在
    bian_name = datetime.datetime.now().strftime('%m-%d')  # 现在
    nowtime01 = datetime.datetime.now().strftime('%M')  # 现在

    if nowtime01 in ['02', '32']:
        print('开始运行。。。。。。。。。。')
    else:
        print('\r检测时间【{}】【{}】'.format(nowtime,nowtime01),end='')
        sleep(1)
        continue
##
    while 1:
        try:
            print('*'*120)
            print('时间：',nowtime)
            biao_tou=['序号','时间']
            url01='https://mms.pinduoduo.com/patronus-mms/order/daily/statisticList'
            data01=qing_post(url01)
            data02=data01.json()
            data03=data02['result']['orderList']

            ge_dirt={}
            for y01 in range(0,len(data03)):
                data04=data03[y01]
                n01=data04['goodsName']
                n02=data04['total']

                if n01 not in biao_tou:
                    biao_tou.append(n01)

                zhuan01=[n01,n02]

                print('数据：',zhuan01)
                cha01=ge_dirt.get(n01,[])

                if not cha01:
                    ge_dirt[n01]=n02
                else:
                    print('数据存在。。。。。。。。。')

            # print(ge_dirt)
            zhuan02=[nowtime]
            for i01 in biao_tou[2:]:
                cha02=ge_dirt.get(i01,'-')
                zhuan02.append(cha02)
#
            print('表头：',biao_tou)
            print('内容：',zhuan02)

            yue01=xie_xls('买菜销量.xls',bian_name)
            # yue01.chuangbiao([])
            yue01.xiubiao02(zhuan02,biao_tou)

            print('等待下次更新。。。。。。。。。。。。。。。')

         #   f = os.popen(r"D:\Users\Administrator\桌面\买菜销量\买菜销量.xls", "r")
           # print('打开表')
          #  time.sleep(10)
          #  print('等10秒')
          #  os.system('taskkill /f /im wps.exe')
          #  print('关闭表')
            sleep(61)
            break

            # for i01 in range(0,sj01):
            #     print('\r还剩下【{}】秒更新。。。。。。。。。。。'.format(sj01-i01),end='')
            #     sleep(1)
        except:
            s = sys.exc_info()
            print("错误第{}行,详情：【'{}' 】".format(s[2].tb_lineno, s[1]).replace('\n', ''))
            sleep(3)






