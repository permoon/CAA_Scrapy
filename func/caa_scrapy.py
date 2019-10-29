# -*- coding: utf-8 -*-
import pandas as pd
from bs4 import BeautifulSoup
import requests


# 檔案清單爬蟲
def caa_scrapy():
    link=[]
    title=[]
    # 取得有效的檔案名稱及連結
    for i in range(1,100000):
        url='https://www.caa.gov.tw/StatisticsYearMonthFile.aspx?a=1091&lang=9&p=' + str(i)
        source = requests.get(url)
        check = source.text
        if check.find('查無資料')>-1:
            break
        soup = BeautifulSoup(source.text, 'html.parser')
        a_tags = soup.find_all('a', class_="download-filebase")     
        for i in a_tags:
            link.append(i.get('href'))
            title.append(i.get('title'))        
    # 創建下載清單df
    dl_list=pd.DataFrame({'link':link,'title':title})
    dl_list=dl_list[dl_list['title'].str.contains("下") & dl_list['title'].str.contains("xls")].reset_index(drop=True)    
    # 取代連結中的無用字元".."
    def replace_c(s):
        return s.replace('..','')    
    dl_list['link']=dl_list['link'].apply(replace_c)
    return dl_list # 回傳所有檔案清單