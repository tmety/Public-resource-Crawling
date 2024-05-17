import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import Chrome
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from time import time
from datetime import datetime, timedelta
import time
import glob
import os
from openpyxl import Workbook
import xlsxwriter
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains#鼠标调用
import io
import PyPDF2
from dateutil.relativedelta import relativedelta
import re
from urllib.parse import quote
from selenium.webdriver.common.by import By  # 选择器
from selenium.webdriver.common.keys import Keys   # 按钮
from selenium.webdriver.support.wait import WebDriverWait   # 等待页面加载完毕
from selenium.webdriver.support import expected_conditions as EC
import os
import shutil

url_wuhan = 'https://www.whzbtb.com/V2PRTS/WinBidBulletinInfoListInit.do'
url_shenzhen = ['https://www.szggzy.com/jygg/list.html?id=jsgc']
url_guangdong = 'https://ygp.gdzwfw.gov.cn/#/44/jygg'
url_shanghai = 'http://www.shcpe.cn/jyfw/xxfw/u1ai51.html'
url_suzhou = 'http://www.jszb.com.cn/jszb/'
url_hangzhou = 'http://zjpubservice.zjzwfw.gov.cn/jyxxgk/list.html'
url_henan = 'http://hnztbkhd.fgw.henan.gov.cn/'
url_chongqin = ['https://www.ccgp-chongqing.gov.cn/info-notice/result-notice','https://www.cqggzy.com/']
url_guangzhou = ['http://www.gzebpubservice.cn/fjzbgg/index.htm']
url_chengdu = 'https://www.cdggzy.com/index.aspx'
url_hainan = 'https://zw.hainan.gov.cn/ggzy/ggzy/jgzbgs/index.jhtml'

# 控制浏览器运行完程序后不会自动关闭
option = Options()
option.add_argument('--ignore-certificate-errors')

option = webdriver.ChromeOptions()
option.add_experimental_option("detach", True)
option.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36")
# 创建浏览器对象
path = Service('./chromedriver.exe')
web = webdriver.Chrome(service=path, options=option)
my_action = ActionChains(web)  # 调用鼠标
# web = Chrome(executable_path='./chromedriver.exe',options=option)
web.maximize_window()  # 最大化浏览器2

def wuhan(url_wuhan):
    # 登录网址
    web.get(url_wuhan)
    time.sleep(5)
    # 找到时间框选项，并输入搜索时间区间
    web.find_element(By.ID, 'searchStartTime').send_keys(start_time)
    web.find_element(By.ID, 'searchEndTime').send_keys(end_time)
    web.implicitly_wait(10)
    time.sleep(3)
    # 点击搜索
    web.find_element(By.CSS_SELECTOR, '#search > table > tbody > tr:nth-child(3) > td:nth-child(3) > a').click()
    time.sleep(5)
    # 找到页面数量展现数量选择项
    new_1 = web.find_element(By.XPATH,
                             '/html/body/div[4]/div[2]/div[2]/div[2]/form/div[2]/div/div[2]/table/tbody/tr/td[1]/select')
    web_new = Select(new_1)
    print(len(web_new.options))
    # 选择页面数量最多的进行切换
    web_new.select_by_index(3)
    time.sleep(5)
    # 确定更换页面数量后总共具有多少页
    yemian = web.find_element(By.XPATH,
                              '/html/body/div[4]/div[2]/div[2]/div[2]/form/div[2]/div/div[2]/table/tbody/tr/td[8]/span')
    yemian = list(yemian.text)[1]
    print('总共需要翻', yemian, '页')
    # 确定总共有多少条数据
    datanum = web.find_element(By.XPATH, '/html/body/div[4]/div[2]/div[2]/div[2]/form/div[2]/div/div[2]/div[1]')
    datanum = datanum.text.split(',')[-1]
    print(datanum)
    # 获取页面数据
    workname = []  # 项目名称
    diyu = []  # 地域
    sucname = []  # 中标人名称
    price = []  # 中标价格
    pepolename = []  # 招标人
    worktime = []  # 发布时间
    for i in range(int(yemian)):
        trs_1 = web.find_elements(By.XPATH,
                                  '/html/body/div[4]/div[2]/div[2]/div[2]/form/div[2]/div/div[1]/div[1]/div[2]/div/table/tbody/tr')
        # aa,bb,cc,dd,ee=[],[],[],[],[]
        for tr_1 in trs_1:
            tr_11 = tr_1.find_elements(By.XPATH, './td')
            a = []
            for i in tr_11:
                ii = i.text
                if ii is None:
                    ii = 0
                else:
                    pass
                a.append(ii)
            print('该条数据爬取完毕')
            print(a)
            if a[4] == '':
                print('添加0')
                price.append(0)
            else:
                print('添加原数据')
                price.append(a[4])
            workname.append(a[1])
            sucname.append(a[3])
            diyu.append('武汉')
            print(a[1])
            print(a[3])
            print(a[4],type(a[4]))
        print('前半部份爬取完毕')
        trs_2 = web.find_elements(By.XPATH,
                                  '/html/body/div[4]/div[2]/div[2]/div[2]/form/div[2]/div/div[1]/div[2]/div[2]/table/tbody/tr')
        for tr_2 in trs_2:
            tr_22 = tr_2.find_elements(By.XPATH, './td')
            b = []
            for i in tr_22:
                ii = i.text
                b.append(ii)
            print('该条数据爬取完毕')
            print(b)
            print(b[2])
            print(b[4])
            pepolename.append(b[2])
            worktime.append(b[4])
            print('后半部份爬取完毕')
        print('这一页数据爬取完毕，开始转换到下一页')
        time.sleep(5)
        web.find_element(By.XPATH,'/html/body/div[4]/div[2]/div[2]/div[2]/form/div[2]/div/div[2]/table/tbody/tr/td[10]/a').click()
        time.sleep(5)
        print('转换到下一页成功')
    print(len(workname),workname)
    print(len(sucname),sucname)
    print(len(price),price)
    print(len(pepolename),pepolename)
    print(len(worktime),worktime)
    df1 = pd.DataFrame({'项目名称':workname,'备注':diyu,'招标人':pepolename,'中标人':sucname,'中标价/万元':price,'中标时间':worktime})
    df1['中标价/万元'] = df1['中标价/万元'].str.strip().str.replace(',', '').fillna(0).astype(str) + '万元'
    df1 = df1.drop_duplicates()
    wuhan_df.append(df1)
    wuhan_df.extend([df1])

def chongqin2(url_chongqin):
    # web.get(url_chongqin[1])
    a = url_chongqin[1]
    # print('输出当前窗口句柄', web.current_window_handle)
    js = 'window.open("' + a + '");'
    web.execute_script(js)
    # print('输出当前窗口句柄',web.current_window_handle)
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    # print('输出句柄集合',handles)
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.close()
            print('关闭原先窗口')
            web.switch_to.window(handle)
            # print('输出当前窗口句柄',web.current_window_handle)
            break
    diyu=[]
    gongshitime=[]
    name=[]
    zhaobiao=[]
    zhongbiao=[]
    pricedata1=[]
    web.implicitly_wait(30)
    time.sleep(3)
    print('开始点击中标结果')
    time.sleep(3)
    try:
        web.find_element(By.CSS_SELECTOR,'#viewGuid > div > div:nth-child(5) > div.span567 > div:nth-child(1) > div.middle-hd > div > a:nth-last-child(2)').click()
    except:
        web.find_element(By.XPATH,'/html/body/div[3]/div/div[5]/div[2]/div[1]/div[1]/div/a[5]').click()
    print('等待目标页面数据加载')
    web.implicitly_wait(30)
    time.sleep(3)
    original_handle = web.current_window_handle
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.switch_to.window(handle)
            break
    # 关闭原始页面
    web.switch_to.window(original_handle)
    web.close()
    # 将句柄切换到新的窗口
    web.switch_to.window(handles[-1])
    web.implicitly_wait(30)
    time.sleep(2)
    biaoshi = ''
    for i in range(50):
        lis = web.find_elements(By.XPATH,'/html/body/div[2]/div[1]/ul/li')
        print('正在爬取第',str(i+1),'页的数据，本页共有',str(len(lis)),'条数据')
        for li in lis:
            diyu1 = li.find_element(By.XPATH,'.//a/span').text.strip()
            name_time1 = li.find_element(By.XPATH,'./span').text.strip()
            print('所属地域：',diyu1,'公示时间：',name_time1)
            date_1 = datetime.strptime(name_time1, "%Y-%m-%d")
            timestamp_1 = datetime.timestamp(date_1)
            if timestamp_1 > timestamp_end:
                print('本条数据不在时间范围内，不做抓取')
            elif timestamp_1 <= timestamp_end and timestamp_1 >= timestamp_start:
                li.find_element(By.XPATH, './/a').click()
                handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
                for handle in handles:  # 切换窗口（切换回去）
                    if handle != web.current_window_handle:
                        web.switch_to.window(handle)
                        break
                web.implicitly_wait(20)
                time.sleep(2)
                print('成功进入到目标页面')
                trs = web.find_elements(By.XPATH,'/html/body/div[2]/div/div[2]/div[6]//table[last()]/tbody/tr')
                for tr in trs:
                    tds = tr.find_elements(By.XPATH,'./td')
                    data1=[]
                    for td in tds:
                        data_last = ''
                        try:
                            data_las = td.text.strip()
                            data_last += data_las
                        except:
                            data = td.find_elements(By.XPATH, './/p')
                            for iiii in data:
                                data_las = iiii.text.strip()
                                data_last += data_las
                        data1.append(data_last)
                        # print(data1)
                    try:
                        if '项目信息' in data1[0]:
                            print('项目信息：',data1[-1])
                            if '最高限价' not in data1[-1]:
                                name.append(data1[-1])
                                diyu.append(diyu1)
                                gongshitime.append(name_time1)
                        elif '招标人信息'in data1[0]:
                            print('招标人：',data1[-1])
                            zhaobiao.append(data1[-1])
                        elif '采购人信息'in data1[0]:
                            print('采购人：',data1[-1])
                            zhaobiao.append(data1[-1])
                        elif '采购信息'in data1[0]:
                            print('采购人：',data1[-1])
                            zhaobiao.append(data1[-1])
                        elif '比选人信息'in data1[0]:
                            print('比选人：',data1[-1])
                            zhaobiao.append(data1[-1])
                        elif '中标人信息'in data1[0]:
                            print('中标人：',data1[-1])
                            zhongbiao.append(data1[-1])
                        elif '成交人信息'in data1[0]:
                            print('成交人：',data1[-1])
                            zhongbiao.append(data1[-1])
                        elif '中选人信息'in data1[0]:
                            print('中选人：',data1[-1])
                            zhongbiao.append(data1[-1])
                        elif '中标金额' in data1[0]:
                            print('中标金额：',data1[-1])
                            pricedata1.append(data1[-1])
                        elif '成交金额' in data1[0]:
                            print('成交金额：',data1[-1])
                            pricedata1.append(data1[-1])
                        elif '中标下浮比例' in data1[0]:
                            print('中标下浮比例：',data1[-1])
                            pricedata1.append(data1[-1])
                        elif '中选价' in data1[0]:
                            print('中选价：',data1[-1])
                            pricedata1.append(data1[-1])
                        elif '中标价' in data1[0]:
                            print('中标价：',data1[-1])
                            pricedata1.append(data1[-1])
                        elif '成交价' in data1[0]:
                            print('成交价：',data1[-1])
                            pricedata1.append(data1[-1])
                        elif '中标折扣率' in data1[0]:
                            print('中标折扣率：',data1[-1])
                            pricedata1.append(data1[-1])
                        elif '中选金额'in data1[0]:
                            print('中选金额：',data1[-2])
                            pricedata1.append(data1[-1])
                        elif '中标比例'in data1[0]:
                            print('中标比例：',data1[-1])
                            pricedata1.append(data1[-1])
                        elif '中标费率'in data1[0]:
                            print('中标费率：',data1[-1])
                            pricedata1.append(data1[-1])
                        elif '中标投标报价'in data1[0]:
                            print('中选金额：',data1[-1])
                            pricedata1.append(data1[-1])
                        else:
                            # print('爬取为空，没有符合条件的数据')
                            pass
                    except:
                        print('爬取不了，不爬这条数据了')
                web.close()  # 关闭当前窗口
                web.switch_to.window(handles[0])
                print('切换回主页面成功==============================')
                web.implicitly_wait(20)
                a = max(len(diyu), len(name), len(zhaobiao), len(zhongbiao), len(pricedata1),len(gongshitime))
                b = min(len(diyu), len(name), len(zhaobiao), len(zhongbiao), len(pricedata1),len(gongshitime))
                if a != b :
                    lists = [diyu,name,zhaobiao,zhongbiao,pricedata1,gongshitime]
                    for lil in range(len(lists)):
                        if len(lists[lil]) < b:
                            print(f'最大的数据个数为：{b}，但这个只有：{len(lists[lil])}')
                            lists[lil].append('没有爬取！')
                        else:
                            pass
            elif timestamp_1 < timestamp_start:
                biaoshi = '结束'
                break
        if biaoshi == '结束':
            print('爬取数据结束**********************')
            break
        else:
            print('点击下一页')
            web.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[3]/div/a[last()-2]').click()  # 点击下一页
            web.implicitly_wait(10)
            time.sleep(3)
    print(len(diyu),diyu)
    print(len(name),name)
    print(len(zhongbiao),zhongbiao)
    print(len(zhaobiao), len(pricedata1), len(gongshitime))
    df2 = pd.DataFrame({'项目名称':name,'备注':diyu,'招标人':zhaobiao,'中标人':zhongbiao,'中标价/万元':pricedata1,'中标时间':gongshitime})
    df2 = df2.drop_duplicates()
    wuhan_df.append(df2)
    wuhan_df.extend([df2])
    result_df = pd.concat(wuhan_df, ignore_index=True)
    merged_df = pd.merge(result_df, df0, how='left', on='中标人')
    print(merged_df)
    merged_df['客户类型'] = np.where(merged_df['客户联系方式'].notnull(), '老客户', '新客户')
    merged_df = merged_df.drop_duplicates()
    merged_df1 = merged_df.sort_values(by='客户联系方式', key=pd.notnull, ascending=False)
    with pd.ExcelWriter(r'D:\工作事项\外部数据\武汉' + start_time1 + '.xlsx') as writer:
        merged_df1.to_excel(writer, index=False, sheet_name='Sheet1')
    time.sleep(3)

def chengdu(url_chengdu):
    a = url_chengdu
    # print('输出当前窗口句柄', web.current_window_handle)
    js = 'window.open("' + a + '");'
    web.execute_script(js)
    # print('输出当前窗口句柄',web.current_window_handle)
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    # print('输出句柄集合',handles)
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.close()
            print('关闭原先窗口')
            web.switch_to.window(handle)
            # print('输出当前窗口句柄',web.current_window_handle)
            break
    name=[]
    diyu=[]
    zhongbiaotime=[]
    zhaobiao=[]
    zhongbiao=[]
    price=[]
    time.sleep(3)
    try:
        web.find_element(By.XPATH,'/html/body/form/div[3]/div[2]/div[1]/div/div[1]/div[1]/a').click()
    except:
        web.find_element(By.CSS_SELECTOR, '#form1 > div.container-fluid.main.site > div:nth-child(8) > div:nth-child(1) > div > div.middle_1.clear > div:nth-child(1) > a').click()
    print('等待目标页面数据加载')
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    # print('输出句柄集合',handles)
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.close()
            print('关闭原先窗口')
            web.switch_to.window(handle)
            # print('输出当前窗口句柄',web.current_window_handle)
            break
    web.implicitly_wait(10)
    time.sleep(3)
    print('配置中标结果条件')
    web.find_element(By.CSS_SELECTOR,'#condition > div:nth-child(1) > div.optionlist.col-xs-10 > div:nth-child(7)').click()
    web.implicitly_wait(30)
    time.sleep(5)
    biaoshi = ''
    for i in range(50):
        print('条件配置完毕，开始爬取第',str(i+1),'页数据')
        divs = web.find_elements(By.CSS_SELECTOR,'#contentlist > div')
        print('本页有',str(len(divs)),'条数据需要爬取')
        for div in divs:
            timeoo = div.find_element(By.XPATH,'./div[@class="item-right"]/div[1]').text.strip()
            date_1 = datetime.strptime(timeoo, "%Y-%m-%d")
            timestamp_1 = datetime.timestamp(date_1)
            if timestamp_1 <= timestamp_end and timestamp_1 >= timestamp_start:
                div.find_element(By.XPATH,'./div[@class="col-xs-10 infotitle"]/a').click()
                handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
                for handle in handles:  # 切换窗口（切换回去）
                    if handle != web.current_window_handle:
                        web.switch_to.window(handle)
                        break
                web.implicitly_wait(20)
                time.sleep(2)
                try:
                    name1 = web.find_element(By.CSS_SELECTOR, '#noticecontent > table > tbody > tr:nth-child(1) > td').text
                    print('该页面出现问题，需要重新加载进入')
                except:
                    web.close()  # 关闭当前窗口
                    web.switch_to.window(handles[0])
                    print('切换回主页面成功==============================')
                    web.implicitly_wait(10)
                    time.sleep(5)
                    div.find_element(By.XPATH, './div[@class="col-xs-10 infotitle"]/a').click()
                    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
                    for handle in handles:  # 切换窗口（切换回去）
                        if handle != web.current_window_handle:
                            web.switch_to.window(handle)
                            break
                    web.implicitly_wait(20)
                    time.sleep(3)
                    # print('成功进入到目标页面')
                name1 = web.find_element(By.CSS_SELECTOR,'#noticecontent > table > tbody > tr:nth-child(1) > td').text
                diyu1 = web.find_element(By.CSS_SELECTOR,'#noticecontent > table > tbody > tr:nth-child(3) > td:nth-child(4)').text
                zhongbiaotime1 = web.find_element(By.CSS_SELECTOR,'#noticecontent > table > tbody > tr:nth-child(5) > td:nth-child(4)').text
                zhaobiao1 = web.find_element(By.CSS_SELECTOR,
                                            '#noticecontent > table > tbody > tr:nth-child(3) > td:nth-child(2)').text
                zhongbiao1 = web.find_element(By.CSS_SELECTOR,
                                         '#noticecontent > table > tbody > tr:nth-child(5) > td:nth-child(2)').text
                price1 = web.find_element(By.CSS_SELECTOR,
                                             '#noticecontent > table > tbody > tr:nth-child(6) > td:nth-child(2)').text
                data_list = [name1, diyu1, zhongbiaotime1, zhaobiao1, zhongbiao1,price1]
                for data_lil in range(len(data_list)):
                    if data_list[data_lil] == '':
                        data_list[data_lil] = '没有爬取！'
                    else:
                        pass
                print('项目名称：',name1,'\n','所属地区：',diyu1,'\n','中标通知书发出时间：',zhongbiaotime1,'\n','招标人：',zhaobiao1,'\n',
                      '中标人：',zhongbiao1,'\n','中标价：',price1)
                name.append(name1)
                diyu.append(diyu1)
                zhongbiaotime.append(zhongbiaotime1)
                zhaobiao.append(zhaobiao1)
                zhongbiao.append(zhongbiao1)
                price.append(price1)
                web.close()  # 关闭当前窗口
                web.switch_to.window(handles[0])
                print('切换回主页面成功==============================')
                web.implicitly_wait(10)
                time.sleep(2)
            elif timestamp_1 < timestamp_start:
                biaoshi = '结束'
                break
            elif timestamp_1 > timestamp_end:
                print('这条数据不在时间范围内，不做抓取')
        if biaoshi == '结束':
            print('数据全部爬取完毕')
            break
        else:
            web.find_element(By.CSS_SELECTOR,'#Pager > a:nth-child(14)').click()
            print('成功点击下一页成功')
            web.implicitly_wait(50)
            time.sleep(3)
    print(len(name),len(zhaobiao),len(zhongbiao),len(price))
    df3 = pd.DataFrame({'项目名称': name, '备注': diyu, '招标人': zhaobiao, '中标人': zhongbiao, '中标价/万元': price,'中标时间': zhongbiaotime})
    wuhan_df.append(df3)
    wuhan_df.extend([df3])
    result_df = pd.concat(wuhan_df, ignore_index=True)
    merged_df = pd.merge(result_df, df0, how='left', on='中标人')
    print(merged_df)
    merged_df['客户类型'] = np.where(merged_df['客户联系方式'].notnull(), '老客户', '新客户')
    merged_df = merged_df.drop_duplicates()
    merged_df1 = merged_df.sort_values(by='客户联系方式', key=pd.notnull, ascending=False)
    with pd.ExcelWriter(r'D:\工作事项\外部数据\武汉'+start_time1+'.xlsx') as writer:
        merged_df1.to_excel(writer, index=False, sheet_name='Sheet1')
    time.sleep(3)

def shenzhen1(url_shenzhen):
    a = url_shenzhen[0]
    # print('输出当前窗口句柄', web.current_window_handle)
    js = 'window.open("' + a + '");'
    web.execute_script(js)
    # print('输出当前窗口句柄',web.current_window_handle)
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    # print('输出句柄集合',handles)
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.close()
            print('关闭原先窗口')
            web.switch_to.window(handle)
            # print('输出当前窗口句柄',web.current_window_handle)
            break
    web.implicitly_wait(10)
    name_work = []
    name_zhaobiao = []
    name_zhongbiao = []
    work_price = []
    work_time = []
    name_data = []
    web.implicitly_wait(10)
    time.sleep(10)
    print('调整参数---点击其它公示')
    web.find_element(By.CSS_SELECTOR, '#gglx-content > dl > dd:nth-child(10) > a').click()
    web.implicitly_wait(10)
    time.sleep(2)
    print('调整参数---点击小型工程直接发包')
    web.find_element(By.CSS_SELECTOR, '#jianshegongcheng > dl > dd:nth-child(3) > a').click()
    web.implicitly_wait(10)
    time.sleep(2)
    print('参数调整完成，开始爬虫')
    biaoshi = ''
    for z in range(100):  # 用手工调整
        lis = web.find_elements(By.XPATH,'/html/body/div[1]/div[3]/div[2]/div[2]/div[2]/div[2]/div/div[5]/div/ul[@id="list_jsgc"]/li')
        print('该页面一共具有', len(lis), '条数据')
        for li in lis:
            data_shenzhen = []
            timeoo = li.find_element(By.XPATH, './/a/span[last()]').text.strip()#这条数据所属的时间
            date_1 = datetime.strptime(timeoo, "%Y-%m-%d")
            timestamp_1 = datetime.timestamp(date_1)
            if timestamp_1 <= timestamp_end and timestamp_1 >= timestamp_start:
                a = li.find_element(By.XPATH, './/a').get_attribute("href")
                # print('输出当前窗口句柄', web.current_window_handle)
                js = 'window.open("' + a + '");'
                web.execute_script(js)
                # print('输出当前窗口句柄',web.current_window_handle)
                handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
                # print('输出句柄集合',handles)
                for handle in handles:  # 切换窗口（切换回去）
                    if handle != web.current_window_handle:
                        # print('switch to ', handle)
                        web.switch_to.window(handle)
                        # print('输出当前窗口句柄',web.current_window_handle)
                        break
                time.sleep(2)
                data_1 = web.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div[5]/div/div[2]').text
                # print(data_1)
                data_shenzhen.append(data_1)
                print('目标数据汇总：',data_shenzhen)
                if data_shenzhen[0] == "":
                    pass
                    web.refresh()
                    time.sleep(4)
                try:
                    aa = re.findall('招标项目名称：\s*(.*?)\\n', data_shenzhen[0])
                    print('项目名称：', aa[0])
                    bb = re.findall('建设单位： \s*(.*?)\\n', data_shenzhen[0])
                    print('招标人：', bb[0])
                    cc = re.findall('承包商： \s*(.*?)\\n', data_shenzhen[0])
                    print('中标人：', cc[0])
                    dd = re.findall('合同价（万元）： \s*(.*?)\\n', data_shenzhen[0])
                    print('合同价(万元)：', dd[0])
                    ee = re.findall('公告发布时间： \s*(.*?)\\n', data_shenzhen[0])
                    print('公告发布时间：', ee[0])
                except:
                    aa = re.findall('招标项目名称：\s*(.*?)\\n', data_shenzhen[0])
                    print('项目名称：', aa[0])
                    bb = re.findall('建设单位：\s*(.*?)\\n', data_shenzhen[0])
                    print('招标人：', bb[0])
                    cc = re.findall('承包商：\s*(.*?)\\n', data_shenzhen[0])
                    print('中标人：', cc[0])
                    dd = re.findall('合同价（万元）： \s*(.*?)\\n', data_shenzhen[0])
                    print('合同价（万元）：', dd[0])
                    ee = re.findall('公告发布时间：\s*(.*?)\\n', data_shenzhen[0])
                    print('公告发布时间：', ee[0])
                if ';' in cc[0]:
                    zhong = cc[0].split(';')
                    for ioi in zhong:
                        name_work.append(aa[0])
                        name_zhaobiao.append(bb[0])
                        name_zhongbiao.append(ioi)
                        work_price.append(dd[0].split('万')[0])
                        work_time.append(ee[0])
                        name_data.append('工程建设-直接发包')
                elif '//' in cc[0]:
                    zhong = cc[0].split('//')
                    for ioi in zhong:
                        name_work.append(aa[0])
                        name_zhaobiao.append(bb[0])
                        name_zhongbiao.append(ioi)
                        work_price.append(dd[0].split('万')[0])
                        work_time.append(ee[0])
                        name_data.append('工程建设-直接发包')
                else:
                    name_work.append(aa[0])
                    name_zhaobiao.append(bb[0])
                    name_zhongbiao.append(cc[0])
                    work_price.append(dd[0])
                    work_time.append(ee[0])
                    name_data.append('工程建设-直接发包')
                web.close()  # 关闭当前窗口（搜狗）
                web.switch_to.window(handles[0])
                # print('输出当前窗口句柄', web.current_window_handle)
                time.sleep(3)  # 切换回窗口
            elif timestamp_1 < timestamp_start:
                biaoshi = '结束'
                break
            elif timestamp_1 > timestamp_end:
                print('本条数据不在时间范围内，不做抓取')
        if biaoshi == '结束':
            print('【工程建设-直接发包】爬取数据结束**********************')
            break
        print('第', z + 1, '页数据爬取完毕，开始点击下一页')
        web.find_element(By.XPATH,'/html/body/div[1]/div[3]/div[2]/div[2]/div[2]/div[2]/div/div[5]/div/div[3]/a[last()-1]').click()
        time.sleep(2)
        print('成功切换到下一页')
    print(len(name_work),len(name_zhaobiao),len(name_zhongbiao),len(work_time),len(work_price), len(name_data))
    print('===================================================\\n','调整参数---点击定标公示')
    web.find_element(By.CSS_SELECTOR,'#gglx-content > dl > dd:nth-child(8) > a').click()
    web.implicitly_wait(10)
    time.sleep(2)
    print('调整参数---点击中标结果公示')
    web.find_element(By.CSS_SELECTOR, '#jianshegongcheng > dl > dd:nth-child(4) > a').click()
    web.implicitly_wait(10)
    time.sleep(2)
    print('参数调整完成，开始爬虫')
    biaoshi = ''
    for z in range(100):#用手工调整
        lis = web.find_elements(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div[2]/div[2]/div[2]/div/div[5]/div/ul[@id="list_jsgc"]/li')
        print('该页面一共具有',len(lis),'条数据')
        for li in lis:
            data_shenzhen = []
            timeoo = li.find_element(By.XPATH, './/a/span[last()]').text.strip()
            date_1 = datetime.strptime(timeoo, "%Y-%m-%d")
            timestamp_1 = datetime.timestamp(date_1)
            if timestamp_1 <= timestamp_end and timestamp_1 >= timestamp_start:
                a = li.find_element(By.XPATH, './/a').get_attribute("href")
                # print('输出当前窗口句柄', web.current_window_handle)
                js = 'window.open("'+a+'");'
                web.execute_script(js)
                # print('输出当前窗口句柄',web.current_window_handle)
                handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
                # print('输出句柄集合',handles)
                for handle in handles:  # 切换窗口（切换回去）
                    if handle != web.current_window_handle:
                        # print('switch to ', handle)
                        web.switch_to.window(handle)
                        # print('输出当前窗口句柄',web.current_window_handle)
                        break
                time.sleep(2)
                data_1 = web.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div[5]/div/div[2]').text
                # print(data_1)
                data_shenzhen.append(data_1)
                print('总数据：',data_shenzhen)
                if data_shenzhen[0] == "":
                    pass
                else:
                    try:
                        aa = re.findall('招标项目名称：\s*(.*?)\\n', data_shenzhen[0])
                        print('项目名称：', aa[0])
                        bb = re.findall('招标人： \s*(.*?)\\n', data_shenzhen[0])
                        print('招标人：', bb[0])
                        cc = re.findall('中标人： \s*(.*?)\\n', data_shenzhen[0])
                        print('中标人：', cc[0])
                        dd = re.findall('中标价\(万元\)： \s*(.*?)\\n', data_shenzhen[0])
                        print('中标价(万元)：', dd[0])
                        ee = re.findall('公示时间： \s*(.*?)至', data_shenzhen[0])
                        print('公示时间：', ee[0])
                    except:
                        try:
                            aa = re.findall('工 程 名 称：\s*(.*?)\\n', data_shenzhen[0])
                            print('项目名称：', aa[0])
                            bb = re.findall('招 标 人：\s*(.*?)\\n', data_shenzhen[0])
                            print('招标人：', bb[0])
                            cc = re.findall('中 标 人：\s*(.*?)\\n', data_shenzhen[0])
                            print('中标人：', cc[0])
                            dd = re.findall('中 标 价：\s*(.*?)\\n', data_shenzhen[0])
                            print('中标价(万元)：', dd[0])
                            ee = re.findall('公 示 日 期：\s*(.*?)\\n', data_shenzhen[0])
                            print('公示时间：', ee[0])
                        except:
                            try:
                                aa = re.findall('招标项目名称：\s*(.*?)\\n', data_shenzhen[0])
                                print('项目名称：', aa[0])
                                bb = re.findall('建设单位：\s*(.*?)\\n', data_shenzhen[0])
                                print('招标人：', bb[0])
                                cc = web.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div[2]/table/tbody/tr[3]/td[2]').text.strip()
                                print('中标人：', cc)
                                dd = web.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div[2]/table/tbody/tr[3]/td[6]').text.strip()
                                print('中标价(万元)：', dd)
                                ee = re.findall('公示时间：\s*(.*?)\\n', data_shenzhen[0])
                                print('公示时间：', ee[0])
                            except:
                                aa = re.findall('项目名称：\s*(.*?)\\n', data_shenzhen[0])
                                print('项目名称：', aa[0])
                                bb = re.findall('招标单位：\s*(.*?)\\n', data_shenzhen[0])
                                print('招标人：', bb[0])
                                cc = web.find_element(By.XPATH,'/html/body/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div[2]/table/tbody/tr[3]/td[2]').text.strip()
                                print('中标人：', cc)
                                dd = web.find_element(By.XPATH,'/html/body/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div[2]/table/tbody/tr[3]/td[3]').text.strip()
                                print('中标价(万元)：', dd)
                                ee = re.findall('公示时间：\s*(.*?)\\n', data_shenzhen[0])
                                print('公示时间：', ee[0])
                    if ';' in cc[0]:
                        zhong = cc[0].split(';')
                        for ioi in zhong:
                            name_work.append(aa[0])
                            name_zhaobiao.append(bb[0])
                            name_zhongbiao.append(ioi)
                            work_price.append(dd[0].split('万')[0])
                            work_time.append(ee[0])
                            name_data.append('工程建设')
                    elif '//' in cc[0]:
                        zhong = cc[0].split('//')
                        for ioi in zhong:
                            name_work.append(aa[0])
                            name_zhaobiao.append(bb[0])
                            name_zhongbiao.append(ioi)
                            work_price.append(dd[0].split('万')[0])
                            work_time.append(ee[0])
                            name_data.append('工程建设')
                    else:
                        name_work.append(aa[0])
                        name_zhaobiao.append(bb[0])
                        name_zhongbiao.append(cc[0])
                        work_price.append(dd[0].split('万')[0])
                        work_time.append(ee[0])
                        name_data.append('工程建设')
                web.close()  # 关闭当前窗口（搜狗）
                web.switch_to.window(handles[0])
                # print('输出当前窗口句柄', web.current_window_handle)
                time.sleep(3)# 切换回窗口
            elif timestamp_1 < timestamp_start:
                biaoshi = '结束'
                break
            elif timestamp_1 > timestamp_end:
                print('本条数据不在时间范围内，不做抓取')
        if biaoshi == '结束':
            print('【工程建设】爬取数据结束**********************')
            break
        print('第',z+1,'页数据爬取完毕，开始点击下一页')
        web.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div[2]/div[2]/div[2]/div/div[5]/div/div[3]/a[last()-1]').click()
        time.sleep(2)
        print('成功切换到下一页')
    print(len(name_work),name_work)
    print(len(name_zhaobiao),name_zhaobiao)
    print(name_zhongbiao)
    print(work_price)
    print(work_time)
    print(name_data)
    df = pd.DataFrame({'项目名称':name_work,'备注':name_data,'招标人':name_zhaobiao,'中标人':name_zhongbiao,'中标价/万元':work_price,'中标时间':work_time})
    df['中标价/万元'] = df['中标价/万元'].str.strip().str.replace(',', '').fillna(0).astype(float)
    df = df.sort_values(by='中标价/万元', ascending=False)
    df = df.drop_duplicates()
    merged_df = pd.merge(df, df0, how='left', on='中标人')
    print(merged_df)
    merged_df['客户类型'] = np.where(merged_df['客户联系方式'].notnull(), '老客户', '新客户')
    merged_df = merged_df.drop_duplicates()
    merged_df = merged_df[merged_df['中标价/万元']>=100]
    merged_df = merged_df[~merged_df['中标人'].str.startswith(('中能', '中建', '中交', '中国', '中铁','中材','中冶','中铁建','中电','华电','国家','国新'))]
    merged_df2 = merged_df.sort_values(by=['客户联系方式','中标价/万元'], key=pd.notnull, ascending=[False, False])
    with pd.ExcelWriter(r'D:\工作事项\外部数据\深圳'+start_time1+'.xlsx') as writer:
        merged_df2.to_excel(writer, index=False, sheet_name='Sheet1')

def shanghai(url_shanghai):
    a = url_shanghai
    # print('输出当前窗口句柄', web.current_window_handle)
    js = 'window.open("' + a + '");'
    web.execute_script(js)
    # print('输出当前窗口句柄',web.current_window_handle)
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    # print('输出句柄集合',handles)
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.close()
            print('关闭原先窗口')
            web.switch_to.window(handle)
            # print('输出当前窗口句柄',web.current_window_handle)
            break
    time.sleep(3)
    #隐藏遮挡元素
    element_to_hide = web.find_element(By.ID, 'footer')  # 替换为实际元素的ID
    web.execute_script("arguments[0].style.display = 'none';", element_to_hide)
    fangshi = []
    leixing = []
    name = []
    zhaobiao = []
    daili = []
    zhongbiao = []
    price = []
    work_time = []
    web.switch_to.default_content()
    frame = web.find_elements(By.TAG_NAME, 'iframe')[0]
    web.switch_to.frame(frame)
    web.implicitly_wait(30)
    time.sleep(2)
    web.find_element(By.ID, 'txtZbrqBegin').send_keys(start_time)
    web.find_element(By.ID, 'txtZbrqEnd').send_keys(end_time)
    web.find_element(By.CSS_SELECTOR, '#btnSearch').click()
    web.implicitly_wait(10)
    time.sleep(3)
    print('条件配置加载完毕')
    tds = web.find_elements(By.CSS_SELECTOR, '#gvZbjgGkList > tbody > tr.pagestyle > td > table > tbody > tr > td')
    yemian = len(tds)
    print('本次爬取数据共有：',str(yemian),'个页面！')
    if yemian == 0:
        yemian = 1
    for i in range(yemian):  # 用手工调整翻页次数
        trs = web.find_elements(By.CSS_SELECTOR,'#gvZbjgGkList > tbody > tr')
        for ii in range(len(trs)-2):
            web.find_element(By.CSS_SELECTOR,'#gvZbjgGkList_lbXmmc_'+str(ii)).click()
            print('点击进入目标内容页面！')
            web.implicitly_wait(10)
            time.sleep(2)
            web.switch_to.default_content()
            frame = web.find_elements(By.TAG_NAME, 'iframe')[0]
            web.switch_to.frame(frame)
            trs = web.find_elements(By.XPATH,'/html/body//table/tbody/tr')
            data = []
            for tr in trs:
                tds = tr.find_elements(By.XPATH,'.//td')
                for td in tds:
                    print(td.text,'----------------')
                    data.append(td.text.strip())
            data = "/-/".join(data)
            # print(data)
            aa = re.findall('招标方式：/-/(.*?)/-/', data)
            print('招标方式：', aa)
            bb = re.findall('招标类型：/-/(.*?)/-/', data)
            print('招标类型：', bb)
            cc = re.findall('招标项目名称：/-/(.*?)/-/', data)
            print('项目名称：', cc)
            dd = re.findall('招标人：/-/(.*?)/-/', data)
            print('招标人：', dd)
            ee = re.findall('招标代理机构：/-/(.*?)/-/', data)
            print('招标代理机构：', ee)
            ff = re.findall('中标人：/-/(.*?)/-/', data)
            print('中标人：', ff)
            gg = re.findall('中标价：/-/(.*?)万元', data)
            print('中标价：', gg)
            hh = re.findall('中标日期：/-/(.*?)/-/', data)
            print('中标日期：', hh)
            if ',' in ff[0]:
                zhong = ff[0].split(',')
                for ioi in zhong:
                    fangshi.append(aa[0])
                    leixing.append(bb[0])
                    name.append(cc[0])
                    zhaobiao.append(dd[0])
                    daili.append(ee[0])
                    zhongbiao.append(ioi)
                    price.append(gg[0])
                    work_time.append(hh[0])
            else:
                fangshi.append(aa[0])
                leixing.append(bb[0])
                name.append(cc[0])
                zhaobiao.append(dd[0])
                daili.append(ee[0])
                zhongbiao.append(ff[0])
                price.append(gg[0])
                work_time.append(hh[0])
            web.switch_to.default_content()
            web.execute_script("window.history.go(-1)")
            web.implicitly_wait(10)
            print('成功退回至主页面')
            web.switch_to.default_content()
            frame = web.find_elements(By.TAG_NAME, 'iframe')[0]
            web.switch_to.frame(frame)
        print('已经成功爬完本页数据，开始爬取下一页数据')
        if i+1 < int(yemian):
            web.find_element(By.CSS_SELECTOR, '#gvZbjgGkList > tbody > tr.pagestyle > td > table > tbody > tr > td:nth-child('+str(i+2)+') > a').click()
            web.implicitly_wait(10)
            time.sleep(2)
        else:
            pass
            web.implicitly_wait(10)
            time.sleep(2)
            web.switch_to.default_content()
            frame = web.find_elements(By.TAG_NAME, 'iframe')[0]
            web.switch_to.frame(frame)
    print(len(name),name)
    print(len(work_time),work_time)
    df = pd.DataFrame({'项目名称':name,'备注':leixing,'招标人':zhaobiao,'中标人':zhongbiao,'中标价/万元':price,'中标时间':work_time})
    df['中标价/万元'] = df['中标价/万元'].str.strip().str.replace(',', '').fillna(0).astype(float)
    df = df.sort_values(by='中标价/万元', ascending=False)
    df = df.drop_duplicates()
    merged_df = pd.merge(df, df0, how='left', on='中标人')
    print(merged_df)
    merged_df['客户类型'] = np.where(merged_df['客户联系方式'].notnull(), '老客户', '新客户')
    merged_df = merged_df.drop_duplicates()
    merged_df = merged_df[merged_df['中标价/万元'] >= 100]
    merged_df = merged_df[~merged_df['中标人'].str.startswith(
        ('中能', '中建', '中交', '中国', '中铁', '中材', '中冶', '中铁建', '中电', '华电', '国家', '国新'))]
    merged_df3 = merged_df.sort_values(by=['客户联系方式','中标价/万元'], key=pd.notnull, ascending=[False, False])
    with pd.ExcelWriter(r'D:\工作事项\外部数据\上海'+start_time1+'.xlsx') as writer:
        merged_df3.to_excel(writer, index=False, sheet_name='Sheet1')

def suzhou(url_suzhou):
    # 登录网址
    a = url_suzhou
    # web.get(a)
    # print('输出当前窗口句柄', web.current_window_handle)
    js = 'window.open("' + a + '");'
    web.execute_script(js)
    # print('输出当前窗口句柄',web.current_window_handle)
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    # print('输出句柄集合',handles)
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.close()
            print('关闭原先窗口')
            web.switch_to.window(handle)
            # print('输出当前窗口句柄',web.current_window_handle)
            break
    time.sleep(3)
    data_diqu = []
    data_leibie = []
    data_xiangmu = []
    data_biaoduan = []
    data_fabiao = []
    data_zhongbiao = []
    data_price = []
    data_worktime = []
    data_time = []
    data_didian = []
    web.find_element(By.CSS_SELECTOR, '#more_6_1 > a').click()
    print('等待目标页面数据加载')
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    # print('输出句柄集合',handles)
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.close()
            print('关闭原先窗口')
            web.switch_to.window(handle)
            # print('输出当前窗口句柄',web.current_window_handle)
            break
    time.sleep(5)
    web.implicitly_wait(10)
    print('加载完毕，拿到输入框')
    web.find_element(By.ID, 'MoreInfoList1_StartDate').send_keys(start_time)
    web.find_element(By.ID, 'MoreInfoList1_EndDate').send_keys(end_time)
    web.find_element(By.CSS_SELECTOR, '#MoreInfoList1_btnOK').click()
    web.implicitly_wait(10)
    print('条件配置完毕，开始爬虫')
    yemian = web.find_element(By.XPATH, '/html/body/table[2]/tbody/tr/td[4]/table/tbody/tr[2]/td/form/table/tbody/tr[7]/td/div/table/tbody/tr/td[1]/font[2]/b').text
    yemian = int(yemian.split('/')[1])
    print('本次爬虫的页面数量为：',yemian,'页')
    for z in range(yemian):
        trs = web.find_elements(By.CSS_SELECTOR, '#MoreInfoList1_DataGrid1 > tbody > tr')
        print('该页面一共具有',len(trs),'条数据')
        for tr in trs:
            diqu = tr.find_element(By.XPATH, './/td[2]/a').text
            diqu = diqu.split(']')[0].split('[')[1]
            print('所属地区：',diqu)
            tr.find_element(By.XPATH, './/td[2]/a').click()
            web.implicitly_wait(10)
            # print('输出当前窗口句柄', web.current_window_handle)
            handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
            # print('输出句柄集合',handles)
            for handle in handles:  # 切换窗口（切换回去）
                if handle != web.current_window_handle:
                    web.switch_to.window(handle)
                    # print('输出当前窗口句柄',web.current_window_handle)
                    break
            time.sleep(3)
            web.implicitly_wait(10)
            # print('等待页面加载完毕')
            data = web.find_element(By.XPATH,'/html/body/form/font/table/tbody/tr/td/table/tbody').text
            # print('该页面所有数据：',data)
            data_1 =web.find_element(By.XPATH,'/html/body/form/font/table/tbody/tr/td/table/tbody/tr[3]/td[2]/span').text.strip()
            print('项目名称：',data_1)
            data_3 =web.find_element(By.XPATH,'/html/body/form/font/table/tbody/tr/td/table/tbody/tr[6]/td[2]/span').text.strip()
            print('发标人：',data_3)
            data_4 = web.find_element(By.XPATH,'/html/body/form/font/table/tbody/tr/td/table/tbody/tr[7]/td[2]/span').text.strip()
            print('项目类型：', data_4)
            data_5 =web.find_element(By.XPATH,'/html/body/form/font/table/tbody/tr/td/table/tbody/tr[9]/td[2]/span').text.strip()
            print('中标人：',data_5)
            data_6 =web.find_element(By.XPATH,'/html/body/form/font/table/tbody/tr/td/table/tbody/tr[12]/td[2]/span').text.strip()
            print('中标价万元：',data_6)
            data_8 = web.find_element(By.XPATH,'/html/body/form/font/table/tbody/tr/td/table/tbody/tr[14]/td[2]/span').text.strip()
            print('中标时间：',data_8)
            web.close()  # 关闭当前窗口（搜狗）
            web.switch_to.window(handles[0])
            # print('输出当前窗口句柄', web.current_window_handle)
            time.sleep(3)
            if ';' in data_5:
                zhong = data_5.split(';')
                for ioi in zhong:
                    if ioi == '':
                        pass
                    else:
                        data_leibie.append(data_4)
                        data_xiangmu.append(data_1)
                        data_fabiao.append(data_3)
                        data_zhongbiao.append(ioi)
                        data_price.append(data_6)
                        data_time.append(data_8)
            else:
                data_leibie.append(data_4)
                data_xiangmu.append(data_1)
                data_fabiao.append(data_3)
                data_zhongbiao.append(data_5)
                data_price.append(data_6)
                data_time.append(data_8)
        print('第',z+1,'页数据爬取完毕，开始点击下一页')
        if z+1 == yemian:
            print('数据全部爬取完毕，开始进入数据整理收尾阶段')
            pass
        else:
            web.find_element(By.XPATH, '/html/body/table[2]/tbody/tr/td[4]/table/tbody/tr[2]/td/form/table/tbody/tr[7]/td/div/table/tbody/tr/td[2]/a[last()-1]').click()
            time.sleep(3)
            print('成功切换到下一页')
    print(len(data_diqu),data_diqu)
    print(data_leibie,'\n',data_xiangmu,'\n',data_biaoduan,'\n',data_fabiao,'\n',data_zhongbiao,'\n',data_price,'\n',data_worktime,'\n',data_time,'\n',data_didian)
    df = pd.DataFrame({'项目名称':data_xiangmu,'备注':data_leibie,'招标人':data_fabiao
                          ,'中标人':data_zhongbiao,'中标价/万元':data_price,'中标时间':data_time})
    df['中标价/万元'] = df['中标价/万元'].str.strip().str.replace(',', '').str.replace(';', '').fillna(0).astype(float)
    df = df.sort_values(by='中标价/万元', ascending=False)
    df = df.drop_duplicates()
    merged_df = pd.merge(df, df0, how='left', on='中标人')
    print(merged_df)
    merged_df['客户类型'] = np.where(merged_df['客户联系方式'].notnull(), '老客户', '新客户')
    merged_df = merged_df.drop_duplicates()
    merged_df = merged_df[merged_df['中标价/万元'] >= 100]
    merged_df = merged_df[~merged_df['中标人'].str.startswith(
        ('中能', '中建', '中交', '中国', '中铁', '中材', '中冶', '中铁建', '中电', '华电', '国家', '国新'))]
    merged_df4 = merged_df.sort_values(by=['客户联系方式','中标价/万元'], key=pd.notnull, ascending=[False, False])
    with pd.ExcelWriter(r'D:\工作事项\外部数据\江苏'+start_time1+'.xlsx') as writer:
        merged_df4.to_excel(writer, index=False, sheet_name='Sheet1')

def hangzhou(url_hangzhou):
    # 登录网址
    a = url_hangzhou
    # print('输出当前窗口句柄', web.current_window_handle)
    js = 'window.open("' + a + '");'
    web.execute_script(js)
    # print('输出当前窗口句柄',web.current_window_handle)
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    # print('输出句柄集合',handles)
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.close()
            print('关闭原先窗口')
            web.switch_to.window(handle)
            # print('输出当前窗口句柄',web.current_window_handle)
            break
    time.sleep(3)
    data_name1 = []
    data_name2 = []
    data_name3 = []
    data_name4 = []
    data_price = []
    data_time = []
    print('等待页面数据加载完毕，点击中标结果选项')
    web.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/ul/li[4]/a').click()
    web.implicitly_wait(10)
    time.sleep(2)
    web.find_element(By.XPATH,'//ul[@id="leftmenu"]/li[last()-1]/a').click()
    web.implicitly_wait(10)
    time.sleep(2)
    print('找到城市选择项')
    for i in range(2,13):
        web.implicitly_wait(10)
        time.sleep(3)
        web.find_element(By.CSS_SELECTOR, '#shijiselect_chosen').click()
        web.implicitly_wait(10)
        time.sleep(5)
        web.find_element(By.CSS_SELECTOR, '#shijiselect_chosen > div > ul > li:nth-child('+str(i)+')').click()
        cityname = web.find_element(By.CSS_SELECTOR, '#shijiselect_chosen > div > ul > li:nth-child('+str(i)+')').text
        print('选择城市成功,该城市为：',cityname)
        web.implicitly_wait(10)
        time.sleep(5)
        #判断有多少页面
        try:
            yemian = web.find_elements(By.XPATH, '/html/body/div[2]/div/div[3]/div[2]/div[@class="pager"]/ul[@class="m-pagination-page"]/li')
            print('本次条件筛选下，共有',len(yemian),'个页面！')
            biaoshi = ''
            for li in yemian:
                li.find_element(By.XPATH,'./a').click()
                print('点击成功')
                web.implicitly_wait(30)
                time.sleep(3)
                print('开始爬取本页数据------------------------------')
                datas = web.find_elements(By.CSS_SELECTOR,'#list > li')
                for lili in datas:
                    qu = lili.find_element(By.XPATH,'.//a/span').text#市里面的区域
                    time00 = lili.find_element(By.XPATH,'./span[@class="ewb-date"]').text.strip()
                    print('所属区域：',qu,'时间：',time00)
                    date_1 = datetime.strptime(time00, "%Y-%m-%d")
                    timestamp_1 = datetime.timestamp(date_1)
                    if timestamp_1 <= timestamp_end and timestamp_1 >= timestamp_start:
                        data_href = lili.find_element(By.XPATH, './/a').get_attribute("href")
                        js = 'window.open("' + data_href + '");'
                        web.execute_script(js)
                        handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
                        for handle in handles:  # 切换窗口（切换回去）
                            if handle != web.current_window_handle:
                                web.switch_to.window(handle)
                                break
                        web.implicitly_wait(50)
                        time.sleep(3)
                        name1 = web.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div/div[1]/div/table/tbody/tr[1]/td[2]/div/b').text
                        print('项目名称：',name1)
                        name2 = web.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div/div[1]/div/table/tbody/tr[3]/td[2]/div/b').text
                        print('标段名称：',name2)
                        name3 = web.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div/div[1]/div/table/tbody/tr[2]/td[2]/div[1]/b').text
                        print('招标人：',name3)
                        name4 = web.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div/div[1]/div/table/tbody/tr[5]/td[1]/div/b').text
                        print('中标人：', name4)
                        price = web.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div/div[1]/div/table/tbody/tr[5]/td[2]/div/b').text.strip().split('元')[0]
                        print('中标价：', price)
                        if ';' in name4:
                            zhong = name4.split(';')
                            for ioi in zhong:
                                data_name1.append(name1)
                                data_name2.append(name2)
                                data_name3.append(name3)
                                data_name4.append(ioi)
                                data_price.append(price)
                                data_time.append(time00)
                        else:
                            data_name1.append(name1)
                            data_name2.append(name2)
                            data_name3.append(name3)
                            data_name4.append(name4)
                            data_price.append(price)
                            data_time.append(time00)
                        web.close()  # 关闭当前窗口（搜狗）
                        web.switch_to.window(handles[0])
                        # print('输出当前窗口句柄', web.current_window_handle)
                        time.sleep(2)  # 切换回窗口
                    elif timestamp_1 < timestamp_start:
                        biaoshi = '结束'
                        break
                    elif timestamp_1 > timestamp_end:
                        print('本条数据不在时间范围内，不做抓取')
                if biaoshi == '结束':
                    print(f'{cityname}爬取数据结束**********************')
                    break
        except:
            pass
    print(len(data_name1))
    print(data_name1)
    print(data_name2)
    print(data_name3)
    print(data_name4)
    print(data_price)
    df = pd.DataFrame({'项目名称':data_name1,'备注':data_name2,'招标人':data_name3,'中标人':data_name4,'中标价/万元':data_price,'中标时间': data_time})
    df['招标人'] = df['招标人'].str.strip().str.replace('名称:', '')
    df['中标价/万元'] = pd.to_numeric(df['中标价/万元'].str.replace(',', '').str.strip('%').str.replace('%', ''), errors='coerce')
    df['中标价/万元'] = df['中标价/万元'].fillna(0)
    df['中标价/万元'] = df['中标价/万元'].apply(lambda x: x / 10000)
    df = df.sort_values(by='中标价/万元', ascending=False)
    df = df.drop_duplicates()
    merged_df = pd.merge(df, df0, how='left', on='中标人')
    print(merged_df)
    merged_df['客户类型'] = np.where(merged_df['客户联系方式'].notnull(), '老客户', '新客户')
    merged_df = merged_df.drop_duplicates()
    merged_df = merged_df[merged_df['中标价/万元'] >= 100]
    merged_df = merged_df[~merged_df['中标人'].str.startswith(
        ('中能', '中建', '中交', '中国', '中铁', '中材', '中冶', '中铁建', '中电', '华电', '国家', '国新'))]
    merged_df5 = merged_df.sort_values(by=['客户联系方式','中标价/万元'], key=pd.notnull, ascending=[False, False])
    with pd.ExcelWriter(r'D:\工作事项\外部数据\浙江'+start_time1+'.xlsx') as writer:
        merged_df5.to_excel(writer, index=False, sheet_name='Sheet1')

def henan(url_henan):
    # 登录网址
    # web.get('about:blank')
    # web.get(url_henan)
    a0 = 'about:blank'
    a = url_henan
    # print('输出当前窗口句柄', web.current_window_handle)
    js = 'window.open("' + a0 + '");'
    web.execute_script(js)
    # print('输出当前窗口句柄',web.current_window_handle)
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    # print('输出句柄集合',handles)
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.close()
            print('关闭原先窗口')
            web.switch_to.window(handle)
            # print('输出当前窗口句柄',web.current_window_handle)
            break
    time.sleep(3)
    js = 'window.open("' + a + '");'
    web.execute_script(js)
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    # print('输出句柄集合',handles)
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.close()
            print('关闭原先窗口')
            web.switch_to.window(handle)
            # print('输出当前窗口句柄',web.current_window_handle)
            break
    time.sleep(3)
    data_name=[]
    data_hangye=[]
    data_fabutime=[]
    data_zhongbiao=[]
    data_price=[]
    zhaobiao=[]
    # 等待页面加载完成
    wait = WebDriverWait(web, 10)  # 最多等待10秒
    wait.until(EC.presence_of_element_located((By.ID, 'iframe')))  # 通过某个元素的ID来判断页面加载完成
    print('点击中标结果板块,并且调整好时间')
    web.find_element(By.XPATH,'//div[@class="tab_tit"]/ul/li[last()-1]/a').click()
    web.implicitly_wait(10)
    time.sleep(2)
    print('开始爬取数据')
    web.switch_to.frame('iframe')
    biaoshi = ''
    for i in range(50):#页面
        print('现在正在爬取第',str(i+1),'页的数据——————————————————————————————————')
        trs = web.find_elements(By.XPATH,'/html/body/table/tbody/tr')[1:]
        print(trs)
        for tr in trs:#数据
            print('开始爬取正式数据！')
            time00 = tr.find_element(By.XPATH, './/td[last()]').text.strip()
            date_1 = datetime.strptime(time00, "%Y-%m-%d")
            timestamp_1 = datetime.timestamp(date_1)
            if timestamp_1 > timestamp_end:
                print('本条数据不在时间范围内，不做抓取')
            elif timestamp_1 <= timestamp_end and timestamp_1 >= timestamp_start:
                print('开始爬取该条数据')
                qudao = tr.find_element(By.XPATH, './/td[4]').text
                if qudao == '中招联合招标采购平台':
                    print('这是采购项目，不获取这个数据')
                else:
                    hangye = tr.find_element(By.XPATH,'.//td[2]/span').text
                    fabu_time = tr.find_element(By.XPATH, './/td[5]').text
                    print(hangye,fabu_time)
                    tr.find_element(By.XPATH,'.//td[1]/a').click()
                    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
                    for handle in handles:  # 切换窗口（切换回去）
                        if handle != web.current_window_handle:
                            web.switch_to.window(handle)
                            break
                    web.implicitly_wait(10)
                    time.sleep(3)
                    name = web.find_element(By.XPATH,'/html/body/div[5]/div[2]/div[1]/h3').text
                    print(name)
                    print('切换到pdf的frame里面')
                    web.switch_to.frame('iframe')
                    web.implicitly_wait(10)
                    data = []
                    try:
                        a = web.find_element(By.CSS_SELECTOR, '#viewer > div:nth-child(1) > div.textLayer').text.strip()
                        data.append(a)
                    except:
                        try:
                            a = web.find_element(By.XPATH,
                                                 '/html/body/div[1]/div[2]/div[4]/div/div[1]/div[2]').text.strip()
                            data.append(a)
                        except:
                            a = '一、中标人信息：没有爬取到数据二、其他'
                            data.append(a)
                    print(data)
                    b = re.findall('一、中标人信息：(.*?)二、其他', data[0].replace("\n", ""))
                    print(b)
                    try:
                        data_1 = re.findall(r'中标人：(.*?)公司', b[0])
                        data_1 = data_1[0] + '公司'
                    except:
                        data_1 = '废标'
                    print(data_1)
                    if data_1 == '废标':
                        pass
                    else:
                        data_zhongbiao.append(data_1)
                        data_hangye.append(hangye)
                        data_fabutime.append(fabu_time)
                        data_name.append(name)
                        zhaobiao.append('未爬取')
                        try:
                            data_2 = re.findall(r'中标价格：(.*?)元', b[0])
                            data_2 = data_2[0] + '元'
                        except:
                            try:
                                data_2 = re.findall(r'中标费率：(.*?)%', b[0])
                                data_2 = data_2[0] + '%'
                            except:
                                data_2 = '其他类型中标价'
                        print(data_2)
                        data_price.append(data_2)
                    web.switch_to.default_content()
                    print('切换回去目标页面成功')
                    web.close()  # 关闭当前窗口（搜狗）
                    web.switch_to.window(handles[0])
                    print('切换回去主页面成功')
                    web.implicitly_wait(10)
                    time.sleep(1)
                    web.switch_to.frame('iframe')
            elif timestamp_1 < timestamp_start:
                biaoshi = '结束'
                break
        if biaoshi == '结束':
            print('爬取数据结束**********************')
            break
        else:
            print('点击下一页')
            web.find_element(By.XPATH,'/html/body/div[2]/a[last()-1]').click()
            web.implicitly_wait(10)
            time.sleep(3)
    print(len(data_name),data_name)
    print(len(data_hangye),data_hangye)
    print(len(data_zhongbiao),data_zhongbiao)
    print(len(data_price),data_price)
    print(len(data_fabutime),data_fabutime)
    df = pd.DataFrame({'项目名称':data_name,'备注':data_hangye, '招标人': zhaobiao ,'中标人':data_zhongbiao,'中标价/万元':data_price,'中标时间':data_fabutime})
    df = df.drop_duplicates()
    df['中标价/万元'] = df['中标价/万元'].str.strip().str.replace('万元', '').apply(pd.to_numeric, errors='coerce')
    df['中标价/万元'] = df['中标价/万元'].fillna(0)
    merged_df = pd.merge(df, df0, how='left', on='中标人')
    merged_df['客户类型'] = np.where(merged_df['客户联系方式'].notnull(), '老客户', '新客户')
    mask = ~(merged_df['中标价/万元'] == 0) | (merged_df['客户类型'] != '新客户')
    mergeddf6 = merged_df.loc[mask]
    mergeddf6 = mergeddf6.drop_duplicates()
    mergeddf6 = mergeddf6[merged_df['中标价/万元'] >= 300]
    mergeddf6 = mergeddf6[~mergeddf6['中标人'].str.startswith(
        ('中能', '中建', '中交', '中国', '中铁', '中材', '中冶', '中铁建', '中电', '华电', '国家', '国新'))]
    merged_df6 = mergeddf6.sort_values(by=['客户联系方式', '中标价/万元'], key=pd.notnull, ascending=[False, False])
    with pd.ExcelWriter(r'D:\工作事项\外部数据\河南'+start_time1+'.xlsx') as writer:
        merged_df6.to_excel(writer, index=False, sheet_name='Sheet1')

def hainan(url_hainan):
    # web.get(url_hainan)
    a = url_hainan
    # print('输出当前窗口句柄', web.current_window_handle)
    js = 'window.open("' + a + '");'
    web.execute_script(js)
    # print('输出当前窗口句柄',web.current_window_handle)
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    # print('输出句柄集合',handles)
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.close()
            print('关闭原先窗口')
            web.switch_to.window(handle)
            # print('输出当前窗口句柄',web.current_window_handle)
            break
    diyu0 = []
    name0 = []
    time0 = []
    zhongbiao0 = []
    price0 = []
    zhao0 = []
    time.sleep(2)
    web.implicitly_wait(30)
    biaoshi = ''
    for i in range(50):
        print(f'现在爬取第{i+1}页数据^^^^^^^^^^^^^^^^^')
        trs = web.find_elements(By.XPATH,'/html/body/div[4]/div[3]/div[2]/table[@class="newtable"]/tbody/tr')[:-2]
        print(f'此页共有{str(len(trs))}条数据')
        for tr in trs:
            diyu = tr.find_element(By.XPATH,'./td[2]').text.strip()
            time00 = tr.find_element(By.XPATH,'./td[last()]').text.strip()
            print('所属地域：',diyu,'中标时间：',time00,)
            date_1 = datetime.strptime(time00, "%Y-%m-%d")
            timestamp_1 = datetime.timestamp(date_1)
            if timestamp_1 <= timestamp_end and timestamp_1 >= timestamp_start:
                print('点击进入目标页面')
                tr.find_element(By.XPATH,'./td[3]/a').click()
                handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
                for handle in handles:  # 切换窗口（切换回去）
                    if handle != web.current_window_handle:
                        web.switch_to.window(handle)
                        break
                web.implicitly_wait(20)
                time.sleep(2)
                print('成功进入到目标页面')
                name = web.find_element(By.CSS_SELECTOR,'body > div.container > div > div.newsTex > h1').text.strip()
                if '中标结果' in name:
                    divs = web.find_elements(By.CSS_SELECTOR,
                                             'body > div.container > div > div.newsTex > div.newsCon > div')
                    if len(divs) > 5:
                        try:
                            zhongbiao = web.find_element(By.CSS_SELECTOR,'body > div.container > div > div.newsTex > div.newsCon > div:nth-child(6) > u').text.strip()
                            price = web.find_element(By.CSS_SELECTOR,'body > div.container > div > div.newsTex > div.newsCon > div:nth-child(7) > u').text.strip()
                            zhaobiao = web.find_element(By.CSS_SELECTOR, 'body > div.container > div > div.newsTex > div.newsCon > div:nth-child(18)').text.strip().split('：')[1]
                            print('第一种：', '\n', name, '\n', zhongbiao, '\n', price, '\n', zhaobiao)
                        except:
                            zhongbiao = web.find_element(By.CSS_SELECTOR,'body > div.container > div > div.newsTex > div.newsCon > div:nth-child(5)').text.strip().split(
                                '：')[1]
                            price = web.find_element(By.CSS_SELECTOR,'body > div.container > div > div.newsTex > div.newsCon > div:nth-child(6)').text.strip().split(
                                '：')[1]
                            zhaobiao = web.find_element(By.CSS_SELECTOR,'body > div.container > div > div.newsTex > div.newsCon > div:nth-child(11)').text.strip().split(
                                '：')[1]
                            print('第三种：', '\n', name, '\n', zhongbiao, '\n', price, '\n', zhaobiao)
                    else:
                        zhongbiao = web.find_element(By.CSS_SELECTOR,'body > div.container > div > div.newsTex > div.newsCon > div:nth-child(5) > div > div:nth-child(2) > div:nth-child(3) > p > span').text.strip()
                        price = web.find_element(By.CSS_SELECTOR,'body > div.container > div > div.newsTex > div.newsCon > div:nth-child(5) > div > div:nth-child(2) > div:nth-child(4) > div > p > span').text.strip()
                        zhaobiao = web.find_element(By.CSS_SELECTOR,'body > div.container > div > div.newsTex > div.newsCon > div:nth-child(5) > div > div:nth-child(5) > div > p:nth-child(1) > span').text.strip()
                        print('第二种：','\n',name,'\n',zhongbiao,'\n',price,'\n',zhaobiao)
                    if ',' in zhongbiao:
                        zhong = zhongbiao.split(',')
                        for ioi in zhong:
                            diyu0.append(diyu)
                            name0.append(name)
                            time0.append(time00)
                            zhongbiao0.append(ioi)
                            price0.append(price)
                            zhao0.append(zhaobiao)
                    elif ';' in zhongbiao:
                        zhong = zhongbiao.split(';')
                        for ioi in zhong:
                            diyu0.append(diyu)
                            name0.append(name)
                            time0.append(time00)
                            zhongbiao0.append(ioi)
                            price0.append(price)
                            zhao0.append(zhaobiao)
                    else:
                        diyu0.append(diyu)
                        name0.append(name)
                        time0.append(time00)
                        zhongbiao0.append(zhongbiao)
                        price0.append(price)
                        zhao0.append(zhaobiao)
                elif '中标公示' in name:
                    print('这条数据为中标公示公告')
                elif '中标候选人' in name:
                    print('这条数据为中标候选人公告')
                else:
                    print('该条数据没有爬取，应该是有未知标题出现，请详细查看！！！！！！！！！！！！！！！！！！！！！！！！！！！')
                web.close()  # 关闭当前窗口
                web.switch_to.window(handles[0])
                print('切换回主页面成功==============================')
                web.implicitly_wait(10)
                time.sleep(1)
            elif timestamp_1 < timestamp_start:
                print('数据已经爬取完毕！！！')
                biaoshi = '结束'
                break
            elif timestamp_1 > timestamp_end:
                print('该条数据不在指定爬取时间范围内，跳过。')
        if biaoshi == '结束':
            break
        else:
            print('开始点击下一页')
            web.find_element(By.CSS_SELECTOR,'body > div.containerNobg > div:nth-child(3) > div.w740 > table > tbody > tr:nth-child(11) > td > div > div > a:nth-child(3)').click()
            time.sleep(2)
            web.implicitly_wait(30)
    print(len(name0), len(zhao0), len(zhongbiao0), len(price0))
    df = pd.DataFrame({'项目名称': name0, '备注': diyu0,'招标人': zhao0,'中标人': zhongbiao0, '中标价/万元': price0,'中标时间': time0})
    df['中标价/万元'] = df['中标价/万元'].str.strip().str.replace('元', '').apply(pd.to_numeric,errors='coerce')
    df['中标价/万元'] = df['中标价/万元'].fillna(0)
    df['中标价/万元'] = df['中标价/万元'].apply(lambda x: x / 10000)
    df1 = df.sort_values(by='中标价/万元', ascending=False)
    df1 = df1.drop_duplicates()
    huanan_df.append(df1)
    huanan_df.extend([df1])
    result_df = pd.concat(huanan_df, ignore_index=True)
    merged_df = pd.merge(result_df, df0, how='left', on='中标人')
    print(merged_df)
    merged_df['客户类型'] = np.where(merged_df['客户联系方式'].notnull(), '老客户', '新客户')
    merged_df = merged_df.drop_duplicates()
    merged_df = merged_df[merged_df['中标价/万元'] >= 100]
    merged_df = merged_df[~merged_df['中标人'].str.startswith(
        ('中能', '中建', '中交', '中国', '中铁', '中材', '中冶', '中铁建', '中电', '华电', '国家', '国新'))]
    merged_df7 = merged_df.sort_values(by=['客户联系方式', '中标价/万元'], key=pd.notnull, ascending=[False, False])
    with pd.ExcelWriter(r'D:\工作事项\外部数据\华南' + start_time1 + '.xlsx') as writer:
        merged_df7.to_excel(writer, index=False, sheet_name='Sheet1')

def guangzhou(url_guangzhou):
    a = url_guangzhou[0]
    # print('输出当前窗口句柄', web.current_window_handle)
    js = 'window.open("' + a + '");'
    web.execute_script(js)
    # print('输出当前窗口句柄',web.current_window_handle)
    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
    # print('输出句柄集合',handles)
    for handle in handles:  # 切换窗口（切换回去）
        if handle != web.current_window_handle:
            web.close()
            print('关闭原先窗口')
            web.switch_to.window(handle)
            # print('输出当前窗口句柄',web.current_window_handle)
            break
    time.sleep(3)
    leibie,name,zhaobiao,zhongbiao,dataprice,datatime= [],[],[],[],[],[]
    class_names = web.find_elements(By.CSS_SELECTOR,'#contentType1 > p')
    for i in range(len(class_names)-1):
        class_name = web.find_element(By.XPATH,'/html/body/div[4]/ul[2]/li[1]/div[@id="contentType1"]/p['+str(i+1)+']').text.strip()
        print('工程类别：',class_name)
        web.find_element(By.XPATH,'/html/body/div[4]/ul[2]/li[1]/div[@id="contentType1"]/p['+str(i+1)+']').click()
        web.implicitly_wait(10)
        time.sleep(2)
        web.find_element(By.XPATH,'/html/body/div[4]/ul[2]/li[2]/div[@class="content"]/p[last()]').click()
        print('点击中标信息完成')
        web.implicitly_wait(10)
        time.sleep(1)
        for i in range(50):
            biaoshi = '开始'
            lis = web.find_elements(By.CSS_SELECTOR,'body > div.jyxx-center > ul.ej-list-type1.ej-list-type1-1 > li.unread')
            for li in lis:
                time00 = li.find_element(By.XPATH,'./p[2]').text.strip()#这条数据所属的时间
                date_1 = datetime.strptime(time00, "%Y-%m-%d")
                timestamp_1 = datetime.timestamp(date_1)
                if timestamp_1 <= timestamp_end and timestamp_1 >= timestamp_start:
                    a = li.find_element(By.XPATH, './p[1]/a').get_attribute("href")
                    # print('输出当前窗口句柄', web.current_window_handle)
                    js = 'window.open("' + a + '");'
                    web.execute_script(js)
                    # print('输出当前窗口句柄',web.current_window_handle)
                    handles = web.window_handles  # 获取当前窗口句柄集合（列表类型）
                    # print('输出句柄集合',handles)
                    for handle in handles:  # 切换窗口（切换回去）
                        if handle != web.current_window_handle:
                            # print('switch to ', handle)
                            web.switch_to.window(handle)
                            # print('输出当前窗口句柄',web.current_window_handle)
                            break
                    web.implicitly_wait(10)
                    time.sleep(2)
                    try:
                        data_name = web.find_element(By.CSS_SELECTOR, 'body > div.xwdt-xq-center > div.content > table > tbody > tr:nth-child(4) > td.label').text.strip()
                        print('项目名称：', data_name)
                        data_zhaobiao = web.find_element(By.CSS_SELECTOR, 'body > div.xwdt-xq-center > div.content > table > tbody > tr:nth-child(6) > td:nth-child(2)').text.strip()
                        print('招标方：', data_zhaobiao)
                        data_zhongbiao = web.find_element(By.CSS_SELECTOR,'body > div.xwdt-xq-center > div.content > table > tbody > tr:nth-child(8) > td:nth-child(1)').text.strip()
                        print('中标方：',data_zhongbiao)
                        data_price0 = web.find_element(By.CSS_SELECTOR,'body > div.xwdt-xq-center > div.content > table > tbody > tr:nth-child(8) > td:nth-child(2) > p').text.strip()
                        if '中标总价(元)' not in data_price0:
                                data_price = 0
                        else:
                            data_price = data_price0.split('(元)：')[1]
                        print('中标总价(元)：',data_price,'\n','----------------------------------------')
                        if ';' in data_zhongbiao:
                            zhong = data_zhongbiao.split(';')
                            for ioi in zhong:
                                leibie.append(class_name)
                                name.append(data_name)
                                zhaobiao.append(data_zhaobiao)
                                zhongbiao.append(ioi)
                                dataprice.append(data_price)
                                datatime.append(time00)
                        else:
                            leibie.append(class_name)
                            name.append(data_name)
                            zhaobiao.append(data_zhaobiao)
                            zhongbiao.append(data_zhongbiao)
                            dataprice.append(data_price)
                            datatime.append(time00)
                    except:
                        try:
                            data_name = web.find_element(By.CSS_SELECTOR,
                                                         'body > div.xwdt-xq-center > div.content > div > div > ul > li:nth-child(1) > span:nth-child(2)').text.strip()
                            print('项目名称：', data_name)
                            data_zhaobiao = web.find_element(By.CSS_SELECTOR,
                                                             'body > div.xwdt-xq-center > div.content > div > div > ul > li:nth-child(3) > span:nth-child(2)').text.strip()
                            print('招标方：', data_zhaobiao)
                            data_zhongbiao = web.find_element(By.CSS_SELECTOR,
                                                              'body > div.xwdt-xq-center > div.content > div > div > ul > li:nth-child(5) > span:nth-child(2)').text.strip()
                            print('中标方：', data_zhongbiao)
                            data_price0 = web.find_element(By.CSS_SELECTOR,
                                                           'body > div.xwdt-xq-center > div.content > div > div > ul > li:nth-child(6)').text.strip()
                            if '中标总价(元)' not in data_price0:
                                data_price = data_price0
                            else:
                                data_price = data_price0.split('(元)：')[1]
                            print('中标总价(元)：', data_price, '\n', '----------------------------------------')
                            if ';' in data_zhongbiao:
                                zhong = data_zhongbiao.split(';')
                                for ioi in zhong:
                                    leibie.append(class_name)
                                    name.append(data_name)
                                    zhaobiao.append(data_zhaobiao)
                                    zhongbiao.append(ioi)
                                    dataprice.append(data_price)
                                    datatime.append(time00)
                            else:
                                leibie.append(class_name)
                                name.append(data_name)
                                zhaobiao.append(data_zhaobiao)
                                zhongbiao.append(data_zhongbiao)
                                dataprice.append(data_price)
                                datatime.append(time00)
                        except:
                            print('没有爬取到这条数据!!!','\n','----------------------------------------')
                            pass
                    web.close()  # 关闭当前窗口（搜狗）
                    web.switch_to.window(handles[0])
                    # print('输出当前窗口句柄', web.current_window_handle)
                    time.sleep(3)  # 切换回窗口
                elif timestamp_1 < timestamp_start:
                    biaoshi = '结束'
                    break
                elif timestamp_1 > timestamp_end:
                    print('本条数据不在时间范围内，不做抓取')
            if biaoshi == '结束':
                print('爬取数据结束**********************')
                break
            print('点击下一页')
            web.find_element(By.CSS_SELECTOR,'body > div.jyxx-center > ul.ej-list-type1.ej-list-type1-1 > div > div > ul.go-after > li:nth-child(1)').click()
            web.implicitly_wait(10)
            time.sleep(2)
    web.close()
    web.quit()
    print(len(leibie),len(name),len(zhaobiao),len(zhongbiao),len(dataprice),len(datatime))
    df = pd.DataFrame({'项目名称':name,'备注':leibie,'招标人':zhaobiao,'中标人':zhongbiao,'中标价/万元':dataprice,'中标时间':datatime})
    if len(name) == 0:
        pass
    else:
        df['中标价/万元'] = df['中标价/万元'].str.strip().str.replace(',', '').apply(pd.to_numeric, errors='coerce')
        df['中标价/万元'] = df['中标价/万元'].fillna(0)
        df['中标价/万元'] = df['中标价/万元'].apply(lambda x: x / 10000)
        df2 = df.sort_values(by='中标价/万元', ascending=False)
        df2 = df2.drop_duplicates()
        huanan_df.append(df2)
        huanan_df.extend([df2])
    result_df = pd.concat(huanan_df, ignore_index=True)
    merged_df = pd.merge(result_df, df0, how='left', on='中标人')
    print(merged_df)
    merged_df['客户类型'] = np.where(merged_df['客户联系方式'].notnull(), '老客户', '新客户')
    merged_df = merged_df.drop_duplicates()
    merged_df = merged_df[merged_df['中标价/万元'] >= 100]
    merged_df = merged_df[~merged_df['中标人'].str.startswith(
        ('中能', '中建', '中交', '中国', '中铁', '中材', '中冶', '中铁建', '中电', '华电', '国家', '国新'))]
    merged_df8 = merged_df.sort_values(by=['客户联系方式','中标价/万元'], key=pd.notnull, ascending=[False, False])
    with pd.ExcelWriter(r'D:\工作事项\外部数据\华南' + start_time1 + '.xlsx') as writer:
        merged_df8.to_excel(writer, index=False, sheet_name='Sheet1')
    time.sleep(3)

def liushi():
    # 获取历史交易数据
    path1 = r'D:\工作事项\客户数据池\客户交易表\数据源\月度台账20截止24.4.xlsx'
    df1 = pd.read_excel(path1, header=0)
    df1 = df1.dropna(subset=['保函共享日'])
    df1['保函共享日'] = pd.to_datetime(df1['保函共享日'])
    path2 = r'D:\工作事项\客户数据池\客户交易表\数据源\5月实时数据.xlsx'
    df2 = pd.read_excel(path2, header=1)
    df2['保函共享日'] = pd.to_datetime(df2['保函共享日'])
    df2['担保金额（元）'] = df2['担保金额（元）'].astype(float)
    df2['担保金额（万）'] = df2['担保金额（元）'].apply(lambda x: x / 10000)
    print(df1.shape, df2.shape)
    # 下面这个时间代表最终要给出数据的时间段
    dataframes = []
    current_date = date_start
    while current_date <= date_end:
        a = current_date.strftime('%Y-%m-%d')
        print(a)
        current_date += timedelta(days=1)
        time_0 = pd.Timestamp(a)# 计算当前时间节点
        time_1 = time_0 - pd.DateOffset(months=6)# 计算半年前的时间节点
        time_2 = time_0 - pd.DateOffset(years=1)# 计算1年前的时间节点
        time_3 = pd.Timestamp(a) - pd.DateOffset(years=2)# 计算2年前的时间节点
        print(time_0,time_1, time_2, time_3)
        # 筛选日期晚于2年前且早于当前时间节点的数据
        se_column_liushi = ['被保证人名称', '保函共享日', '客户经理名称','客户经理所属机构' ,'担保金额（万）', '订单编号', '担保止期']
        new_df_liushi_1 = df1[se_column_liushi]
        new_df_liushi_2 = df2[se_column_liushi]
        new_df_liushi = pd.concat([new_df_liushi_1, new_df_liushi_2], ignore_index=True)
        new_df_liushi_a = new_df_liushi[new_df_liushi['保函共享日'] <= time_0]
        new_df_liushi = new_df_liushi_a[new_df_liushi_a['保函共享日'] >= time_3]
        new_dfls = new_df_liushi.drop_duplicates()
        new_dfls = new_dfls.sort_values(by=['被保证人名称', '客户经理名称', '保函共享日'], ascending=[1, 0, 0])
        new_dfls['保函共享日'] = pd.to_datetime(new_dfls['保函共享日'])
        new_dfls['保函共享日'] = new_dfls['保函共享日'].dt.strftime("%Y-%m-%d")
        df_liushi = new_dfls
        col_a, col_b, col_c, col_d, col_e = ['被保证人名称'], ['客户经理'], ['累积交易时间'], ['客户经理所属机构'],['匹配']
        for i, row in df_liushi.iterrows():
            col_1 = row.loc['被保证人名称']
            col_2 = row.loc['保函共享日']
            col_3 = row.loc['客户经理名称']
            col_4 = row.loc['客户经理所属机构']
            if col_1 == col_a[-1]:
                if col_3 == col_b[-1]:  # 判断是同一个客户经理
                    col_c[-1].append(col_2)
                else:  # 判断不是同一个客户经理
                    col_a.append(col_1)  # 被保证人名称
                    col_b.append(col_3)  # 客户经理
                    col_d.append(col_4)  # 客户经理所属机构
                    riqi3 = []
                    riqi3.append(col_2)
                    col_c.append(riqi3)  # 累积交易时间
                    col_e.append(col_1 + col_3)
            else:
                col_a.append(col_1)  # 被保证人名称
                col_b.append(col_3)  # 客户经理
                col_d.append(col_4)  # 客户经理所属机构
                riqi1 = []
                riqi1.append(col_2)
                col_c.append(riqi1)  # 累积交易时间
                col_e.append(col_1 + col_3)
        a11_liushi = pd.DataFrame({'被保证人名称': col_a, '客户经理': col_b, '累积交易时间': col_c, '客户经理所属机构': col_d, '匹配': col_e})
        a11_liushi.drop(0, inplace=True)
        col_a, col_b, col_c, col_d = ['被保证人名称'], ['客户经理'], ['流失状态'], ['客户经理所属机构']
        for i, row in a11_liushi.iterrows():
            col_1 = row.loc['被保证人名称']
            col_2 = row.loc['客户经理']
            col_3 = row.loc['累积交易时间']
            col_4 = row.loc['客户经理所属机构']
            time0 = pd.to_datetime(col_3[0])
            if time0 == time_2:  # 等于1年前
                col_c3 = '已流失'
            elif time0 == time_1:  # 等于半年前
                col_c3 = '流失预警'
            elif time0 == time_0:  # 等于现在
                if len(set(col_3)) == 1:
                    col_c3 = '新增'
                elif len(set(col_3)) > 1:
                    col_c3 = '重新唤醒老客户'
                    a_aa = [x for x in col_3 if x != col_3[0]]
                    for i in range(len(a_aa)):
                        aaa = pd.to_datetime(a_aa[i])
                        if aaa > time_2:
                            col_c3 = '正常老客户'
                            break
                        else:
                            pass
            else:
                col_c3 = '未知'
            if col_c3 == '未知':
                pass
            else:
                col_a.append(col_1)
                col_b.append(col_2)
                col_c.append(col_c3)
                col_d.append(col_4)
        data_liushi = pd.DataFrame({'客户经理所属机构': col_d, '客户经理': col_b, '被保证人名称': col_a,'流失状态': col_c})
        data_liushi.drop(0, inplace=True)
        data_liushi = data_liushi.assign(column_name=1)
        data_liushi.rename(columns={'column_name': '笔数'}, inplace=True)
        data_liushi = data_liushi.assign(column_name=a)
        data_liushi.rename(columns={'column_name': '时间'}, inplace=True)
        dataframes.append(data_liushi)
    result_df = pd.concat(dataframes, ignore_index=True)
    result_df = result_df[result_df['流失状态'].str.contains('已流失|流失预警')]
    # result_df = result_df[~result_df['客户经理所属机构'].str.contains('渠道')]
    print(result_df)
    result_df = result_df.sort_values(by='客户经理所属机构')
    with pd.ExcelWriter(r'D:\工作事项\外部数据\每日流失名单.xlsx') as writer:
        result_df.to_excel(writer, index=False, sheet_name='Sheet1')
    time.sleep(3)
    return result_df

def data_qixi(data):
    import PyOfficeRobot
    file_name = ['深圳', '上海', '江苏', '浙江', '武汉', '华南', '河南']
    data_all = []
    for i in file_name:
        path1 = r'D:\工作事项\外部数据\\' + i+ start_time1 + '.xlsx'
        df1 = pd.read_excel(path1, header=0)
        mask = df1['区域'].fillna('').str.contains('渠道|青岛|长沙', regex=True)
        row_indices = df1[mask].index.tolist()
        if not row_indices:
            print("没有找到包含指定字符串的行")
        else:
            df1.iloc[row_indices, 6:] = np.nan
            df1.iloc[row_indices, 12] = '新客户'
        df1['区域'] = df1['区域'].fillna(i)
        df1.loc[df1['区域'].str.contains('成都|武汉'), '区域'] = '武汉'
        df1.loc[df1['区域'].str.contains('广州|海口|华南'), '区域'] = '华南'
        df1.loc[df1['区域'].str.contains('苏州|泰昆|中颐'), '区域'] = '江苏'
        df1.loc[df1['区域'].str.contains('杭州|合肥'), '区域'] = '浙江'
        df1.loc[df1['区域'].str.contains('深圳'), '区域'] = '深圳'
        df1.loc[df1['区域'].str.contains('上海'), '区域'] = '上海'
        df1.loc[df1['区域'].str.contains('郑州'), '区域'] = '河南'
        data_all.append(df1)
    all_df = pd.concat(data_all, ignore_index=True)
    result_df = data
    result_df.loc[result_df['客户经理所属机构'].str.contains('武汉|成都|重庆'), '客户经理所属机构'] = '武汉'
    result_df.loc[result_df['客户经理所属机构'].str.contains('海口|广州|华南'), '客户经理所属机构'] = '华南'
    result_df.loc[result_df['客户经理所属机构'].str.contains('河南|郑州'), '客户经理所属机构'] = '河南'
    result_df.loc[result_df['客户经理所属机构'].str.contains('苏州'), '客户经理所属机构'] = '江苏'
    result_df.loc[result_df['客户经理所属机构'].str.contains('上海'), '客户经理所属机构'] = '上海'
    result_df.loc[result_df['客户经理所属机构'].str.contains('深圳|福田'), '客户经理所属机构'] = '深圳'
    result_df.loc[result_df['客户经理所属机构'].str.contains('杭州|合肥'), '客户经理所属机构'] = '浙江'
    for i in file_name:
        dfa = all_df[all_df['区域'] == i]
        dfa.loc[:, '近1年交易频率'] = dfa['近1年交易频率'].fillna(0).astype(float)
        dfa1 = dfa.sort_values(by=['客户类型', '近1年交易频率', '中标价/万元'], ascending=[False,False, False])
        dfb = result_df[result_df['客户经理所属机构'] == i]
        dfb1 = dfb.sort_values(by=['流失状态'], key=pd.notnull, ascending=[False])
        a = (dfa1['客户类型'].str.contains('客户')).sum()  # 所有中标个数
        b = (dfa1['客户类型'] == '老客户').sum()  # 老客户
        c = (dfa1['客户类型'] == '新客户').sum()  # 新客户
        d = (dfa1['近1年交易频率'] != 0).sum()  # 近1年交易的老客户
        e = (dfb1['流失状态'] == '已流失').sum()  # 已流失客户
        f = (dfb1['流失状态'] == '流失预警').sum()  # 流失预警客户
        text = i + start_time1 + '中标人个数' + str(a) + '{ctrL}{ENTER}' + '老客户有' + str(b) + '个，近1年成交过的有' + str(
            d) + '个' + '{ctrL}{ENTER}' + '新客户有' + str(c) + '个'+'{ctrL}{ENTER}' + '已流失客户有' + str(e) + '个，流失预警客户有' + str(
            f) + '个'
        print(text)
        with pd.ExcelWriter(r'D:\工作事项\外部数据\\' + i + start_time1 + '.xlsx') as writer:
            dfa1.to_excel(writer, index=False, sheet_name='及时中标名单')
            dfb1.to_excel(writer, index=False, sheet_name='流失名单')
        PyOfficeRobot.chat.send_message(who=name_data, message=text)
        PyOfficeRobot.file.send_file(who=name_data, file=r'D:\工作事项\外部数据\\' + i + start_time1 + '.xlsx')
        new_filename = r'D:\工作事项\外部数据\\' + i + start_time1 + '.xlsx'
        target_folder_path = 'D:\工作事项\外部数据\历史数据\\'+i
        shutil.move(new_filename, target_folder_path)


if __name__ == '__main__':
    start_time = '2024-05-16'
    end_time = '2024-05-16'
    name_data = '大树（张树颖）'
    date_start = datetime.strptime(start_time, "%Y-%m-%d")
    date_end = datetime.strptime(end_time, "%Y-%m-%d")
    timestamp_start = datetime.timestamp(date_start)
    timestamp_end = datetime.timestamp(date_end)
    start_time1 = start_time.split('24-')[1]
    time0 = time.time()
    path = r'D:\工作事项\外部数据\2020年至今客户数据（更新5.13）.xlsx'
    df0 = pd.read_excel(path, header=0)
    data_reult_df = pd.DataFrame()
    wuhan_df=[]
    huanan_df = []
    wuhan(url_wuhan)
    chongqin2(url_chongqin)
    chengdu(url_chengdu)
    shenzhen1(url_shenzhen)
    shanghai(url_shanghai)
    suzhou(url_suzhou)
    hangzhou(url_hangzhou)
    henan(url_henan)
    hainan(url_hainan)
    guangzhou(url_guangzhou)
    data_qixi(data = liushi())
    time1 = time.time()
    time_difference = time1 - time0
    minutes, seconds = divmod(time_difference, 60)
    time_format = "{:02d}:{:02d}".format(int(minutes), int(seconds))
    print("本次运行消耗的时间：", time_format)