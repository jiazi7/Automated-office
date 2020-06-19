import xlwings as xw
import re
import calendar
import json
from lxml import etree
import requests
import scrapy
import os
import csv
from aip import AipOcr
import datetime 
import time
import pandas as pd
import numpy as np
from WindPy import w


@xw.sub  # only required if you want to import it or run it via UDF Server
def main():
    book = xw.Book.caller()


    """爬取农业部数据"""
    today= datetime.date.today()
    #获取当月第一天
    firstmonthday=datetime.datetime(today.year,today.month,1)
    #获取当年第一天
    firstday=datetime.datetime(today.year,1,1)
    oneday=datetime.timedelta(days=1)
    all={}#存放所有爬取的网页链接，key为日期，value为链接
    url='http://www.scs.moa.gov.cn/scxxfb/'

    #爬取主页
    response=requests.get(url)
    content=response.content
    page=etree.HTML(content)
    data=page.find('.//div[@class="sj_e_tonzhi_list"]')

    for i in data:
        infos=i.findall('.//li')
        for info in infos:
            rrr=info.find('.//a')
            link=url+str(rrr.get('href'))
            date=re.findall(r'.\w+.t(\d+)\w+',link)
            all[date[0]]=str(link)

    for i in range(1,13):
        url='http://www.scs.moa.gov.cn/scxxfb/index_'+str(i)+'.htm'
        response=requests.get(url)
        content=response.content
        page=etree.HTML(content)
        data=page.find('.//div[@class="sj_e_tonzhi_list"]')

        for i in data:
            infos=i.findall('.//li')
            for info in infos:
                rrr=info.find('.//a')
                link='http://www.scs.moa.gov.cn/scxxfb/'+str(rrr.get('href'))
                date=re.findall(r'.\w+.t(\d+)\w+',link)
                all[date[0]]=str(link)



    #print(all)
    #爬取目标页、正则提取猪肉价格
    def price_get(link):
        response=requests.get(link)
        content=response.content
        page=etree.HTML(content)
        info=page.find('.//div[@class="TRS_Editor"]')
        text=info.find('.//p').text
        price=re.findall(r'猪肉\D+(\d+.\d+)元',text)
        return price


    price1 = {}#存放猪肉价格，key为日期，value为价格

    #今天的价格，若未更新则为前一天价格
    while today.strftime('%Y%m%d') not in all.keys():
        today-=oneday   
    else:
        d_p_price=price_get(str(all[today.strftime('%Y%m%d')]))
        price1[today.strftime('%Y%m%d')]=d_p_price

    #本月初价格，更新时间为本月第一个工作日
    while firstmonthday.strftime('%Y%m%d') not in all.keys():
        firstmonthday+=oneday   
    else:
        m_p_price=price_get(str(all[firstmonthday.strftime('%Y%m%d')]))
        price1[firstmonthday.strftime('%Y%m%d')]=m_p_price

    #本年初价格，更新时间为本年第一个工作日
    while firstday.strftime('%Y%m%d') not in all.keys():
        firstday+=oneday   
    else:
        y_p_price=price_get(str(all[firstday.strftime('%Y%m%d')]))
        price1[firstday.strftime('%Y%m%d')]=y_p_price

    #对应价格的列表
    #l=[price1[today.strftime('%Y%m%d')],price1[firstmonthday.strftime('%Y%m%d')],price1[firstday.strftime('%Y%m%d')]]

    print(price1)

    '''
    最终结果是price1是一个字典
    pric1e[today.strftime('%Y%m%d')]是今日价格
    price1[firstmonthday.strftime('%Y%m%d')]是本月初
    price1[firstday.strftime('%Y%m%d')]是本年初
    '''

    """爬取二元能繁母猪数据"""
    #百度云账号
    APP_ID = '#####'
    API_KEY = '########'
    SECRECT_KEY = '########'
    client = AipOcr(APP_ID, API_KEY, SECRECT_KEY)


    #爬取主页，获取目标网页链接
    url='http://sousuo.gov.cn/s.htm?q=%E4%BA%8C%E5%85%83%E6%AF%8D%E7%8C%AA%E9%94%80%E5%94%AE%E4%BB%B7%E6%A0%BC&t=govall&timetype=timeqb&mintime=&maxtime=&sort=pubtime&sortType=1&nocorrect='
    response = requests.get(url)
    content = response.content
    page = etree.HTML(content)
    table = page.find('.//h3[@class="res-title"]')
    channels = table.find('.//a')
    link=channels.get('href')
    #print(link)


    #爬取最新公告的标题
    html = requests.get(link)
    html.encoding = 'utf-8'
    text = html.text
    page1 = etree.HTML(text) 
    info=page1.find('.//div[@class="article oneColumn pub_border"]')
    t=info.find('.//h1')
    title = t.text
    #print(title)

    #从公告标题中提取更新数据对应的日期
    datestr = title[len(title)-14:len(title)-9]
    date= '2020年'+datestr
    date1=datetime.datetime.strptime(date,'%Y年%m月%d日')
    #print(date1)

    #爬取公告中的图片
    content1 = page1.find('.//div[@class="pages_content"]')
    channels1 = content1.find('.//img')
    link_img=channels1.get('src')
    links = str(link)
    pic_urls = links[:len(links)-19]+link_img

    #调用百度api对图片进行文本识别，从中提取价格内容
    prices=client.basicGeneralUrl(pic_urls)
    r=prices['words_result']
    info=r[5]
    price=info['words']
    #print(price)

    #更新每周二元母猪价格
    pork_price={}
    week = date1.strftime("%W")
    pork_price[week+'周']=price #存放每周二元母猪价格，key为周数，对应价格
    print(pork_price)


    """抓取wind数据 写入excel"""

    #链接到wind数据库
    w.start()
    w.isconnected()
    
    #统计仔猪数据
    ##download仔猪数据
    pig_baby_codes = ['############']###仔猪代码已打码
    pig_baby = w.edb(pig_baby_codes,datetime.date.today()+datetime.timedelta(days=-5), datetime.date.today(), usedf = True, ShowBlank=0)
    pig_baby = pig_baby[1]
    pig_baby.columns = ['###########']###仔猪地区标签已打码

    ##分地区统计仔猪数据
    pig_baby_mean = pd.DataFrame([])
    pig_baby_mean_names = ['##########']###仔猪分地区统计的地区标签已打码                      
    for i in range(1, 13, 2):
        pig_baby_mean[pig_baby_mean_names[int((i-1)/2)]] = (pig_baby.iloc[:,i-1]+pig_baby.iloc[:,i])/2
    print(pig_baby_mean)


                        
    #生猪                                            
    ##download生猪数据
    pig_codes = ["###############"]###生猪代码已打码
    pig = w.edb(pig_codes,datetime.date.today()+datetime.timedelta(days=-4), datetime.date.today(), usedf = True,ShowBlank=0)
    pig = pig[1]
    pig.columns = ["###############"]###生猪地区标签已打码

    ##分地区统计仔猪数据
    pig_mean = pd.DataFrame(np.zeros((4, 5)))
    pig_mean_names = ["###########"]###生猪分地区统计的地区标签已打码
    pig_mean.columns = pig_mean_names
    print(pig_mean)
    pig_mean.index = pig.index[1:]
    for name in pig_mean_names:
        i = 0
        for n in list(pig.columns):
            if name in n:
                pig_mean[name] = pig_mean[name] + pig[n]
                i += 1
        pig_mean[name] = pig_mean[name] / i

    print(pig_baby_mean)

    #统计玉米数据       
    ##donload玉米价格
    corn_codes = ['S5005793']
    corn = w.edb(corn_codes,datetime.date.today()+datetime.timedelta(days=-5), datetime.date.today(), usedf = True, ShowBlank=0)
    corn = corn[1]
    corn.columns = ['现货价:玉米:平均价']
    corn = corn.T
    print(corn)

    #关闭Wind接口
    w.stop()

    #仔猪、生猪、猪肉、玉米价格汇总
    pig_baby_mean = pig_baby_mean.T
    pig_mean = pig_mean.T
    pig_baby_data = list(pig_baby_mean[pig_baby_mean.columns[-1]])
    pig_baby_data.append(np.mean(pig_baby_data))
    pig_data = list(pig_mean[pig_mean.columns[-1]])
    pig_data.append(np.mean(pig_data))
    corn_data = list(corn[corn.columns[-1]])
    pig_baby_data.extend(pig_data)
    pig_baby_data.extend(corn_data)
    pig_baby_data.append(float(price1[today.strftime('%Y%m%d')][0]))
    alldata = pig_baby_data
    print(alldata)

    #最近5日日期的一个list——days是datetime格式列表，days1是字符格式列表
    days = [datetime.datetime.today()+datetime.timedelta(days=-i) for i in range(5)]
    days1 = [days[i].strftime( '%Y-%m-%d') for i in range(5)]
    days.reverse()
    days1.reverse()
    print(days)

    #最近五周的一个list——week_nows
    week_list={}
    today = datetime.date.today()
    weeks=today.strftime("%W")
    week_n=int(weeks)
    week_list[week_n]=week_n
    l=[week_list[week_n]-i for i in range(5)]
    for i in range(5):
        l[i]=str(l[i])+'周'
    l.reverse()
    print(l)
    week_nows = l


    #链接到目标表格
    sht = book.sheets[0]

    #判断二元能繁母猪年度数据、月度数据是否要更新
    firstday_week = datetime.datetime(datetime.date.today().year, datetime.date.today().month, 1).strftime("%W")+'周'
    if week_nows[-1] == '1周':
        sht.range('Q8').value = float(price)
    if  week_nows[-1] == firstday_week:
        sht.range('P8').value = float(price)

    #判断仔猪、生猪、猪肉、玉米年度数据、月度数据是否要更新
    if days1[-1][6:] == '01-01':
        sht.range('Q11:Q25').options(transpose=True).value = alldata
    if days1[-1][9:] == '01':
        sht.range('P11:P25').options(transpose=True).value = alldata
        
    #更新主体数据（若今天数据已更新则不再更）
    ##二元能繁母猪
    if sht.range('K7').value == week_nows[-1]:
        pass
    else:
        sht.range('G8:J8').value = sht.range('H8:K8').value
        sht.range('K8').value = float(price)

    ##仔猪、生猪、猪肉、玉米
    if sht.range('K9').value.date() == days[-1].date():
        pass
    else:
        sht.range('G7:K7').value = week_nows
        sht.range('G9:K9').value = days1
        sht.range('G11:J25').value = sht.range('H11:K25').value
        sht.range('K11:K25').options(transpose=True).value = alldata


if __name__ == "__main__":
    xw.books.active.set_mock_caller()
    main()
