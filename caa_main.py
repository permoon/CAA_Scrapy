# -*- coding: utf-8 -*-

from os import listdir, path
import pandas as pd
import xlrd
from datetime import datetime
from func.caa_reshape import caa_reshape #資料重整模組
from func.caa_scrapy import caa_scrapy #民航局爬蟲模組

# 設定取得檔案路徑及想拆解的sheet清單
path=path.join(path.dirname(__file__)) 

# 設定拆解目標sheet
sheet_target = ['36-1','36-2', '36-3','36-4','36-5','36-6']


# 取得過去已下載檔案清單
files = listdir(path + '/src')

# 篩選網頁上新增的檔案，排除過去已經下載、處理過的檔案
dl_list=caa_scrapy()
dl_list=dl_list[~dl_list['title'].isin(files)]


# 下載本次新增檔案
from urllib.request import urlretrieve
for l,t in list(zip(list(dl_list['link']),list(dl_list['title']))):
    url= 'https://www.caa.gov.tw'+l
    urlretrieve(url, path + '/src/' + t )


# 創造初始空dataframe
df_all=pd.DataFrame(columns=['DESTINATION', 'AIRLINE', 'AIRPLANE_TOTAL', 
                             'SEAT_TOTAL', 'PEOPLE_TOTAL', 'OCCUPANCYRATE_TOTAL', 
                             'AIRPLANE_INBOUND', 'SEAT_INBOUND', 'PEOPLE_INBOUND', 
                             'OCCUPANCYRATE_INBOUND', 'AIRPLANE_OUTBOUND', 'SEAT_OUTBOUND', 
                             'PEOPLE_OUTBOUND', 'OCCUPANCYRATE_OUTBOUND']) 

# 整理+合併資料
for f in list(dl_list['title']):
    for s in sheet_target:
        try:
            df_add=caa_reshape(f,s, path+'/src')
            df_all=df_all.append(df_add,ignore_index = True, sort=False)
        except xlrd.XLRDError:
            print ("找不到"+f+"的"+s+"這張資料表")


# 新增timestamp
df_all['CREATE_DATE']=datetime.today().strftime("%Y%m%d")

# 設定資料型別
c_list=['AIRPLANE_TOTAL', 'SEAT_TOTAL', 'PEOPLE_TOTAL', 
        'AIRPLANE_INBOUND', 'SEAT_INBOUND', 'PEOPLE_INBOUND', 
        'AIRPLANE_OUTBOUND', 'SEAT_OUTBOUND', 
        'PEOPLE_OUTBOUND','DATA_YEAR','DATA_MONTH']
r_list=['OCCUPANCYRATE_TOTAL','OCCUPANCYRATE_INBOUND','OCCUPANCYRATE_OUTBOUND']

for i in c_list:
    df_all[i]=df_all[i].astype(int)

for i in r_list:
    df_all[i]=(df_all[i].astype(float)).round(2)
    


# 另存excel
df_all.to_excel(path + '/' + 'caa_' + str(datetime.now().strftime('%Y%m%dT%H%M')) + ".xlsx",sheet_name='caa') 

