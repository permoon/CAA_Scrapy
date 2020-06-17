# -*- coding: utf-8 -*-
#### logger ##################################################
import logging
from func.loggy import log_init
log_init('./log/caaStatsMonthly.log')

#### main ###################################################
# 民航統計月報
# https://www.caa.gov.tw/StatisticsYearMonthFile.aspx?a=1091&lang=9&p=

from os import listdir, path
import pandas as pd
import xlrd
import func.caaStatsMonthly as caaStatsMonthly

# ssl varify
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

### file update 
# Newfiles on web
caaStatsMonthly.scrapyUpdate(key_word='下')
caaStatsMonthly.scrapyUpdate(key_word='上')


# 下載近期新增檔案
from urllib.request import urlretrieve

new_list=pd.DataFrame({'link':[],
                       'title':[],
                       'type':[]})

for i,j in zip(['_filelist_outbound.csv','_filelist_local.csv'],
               ['outbound','local']):
    files = listdir('./src/caaStatsMonthly/'+j) # 取得過去已下載檔案清單
    dl_list=pd.read_csv('./src/caaStatsMonthly/'+i) #  讀取目前網站上所有檔案清單
    dl_list=dl_list[~dl_list['title'].isin(files)] # 保留尚未下載清單
    dl_list['type']=j
    new_list=new_list.append(dl_list)
    # 迴圈抓抓抓抓抓
    try: 
        for l,t in list(zip(list(dl_list['link']),list(dl_list['title']))):
            url= 'https://www.caa.gov.tw'+l
            urlretrieve(url, './src/caaStatsMonthly/'+j +'/'+ t )
            logging.info(t + ': FileDownloads OK')
    except:
        logging.error(t + ': FileDownloads ERROR')      



### main loop(outbound) ###

# 設定拆解目標sheet
sheet_target = ['36-1','36-2', '36-3','36-4','36-5','36-6']

# 創造初始空dataframe
df_outbound=pd.DataFrame() 

# 取國際航空新增資料
dl_list=new_list[new_list['type']=='outbound']

# 整理+合併資料
for f in list(dl_list['title']):
    for s in sheet_target:
        try:
            df_add=caaStatsMonthly.reshape(f,s)
            df_outbound=df_outbound.append(df_add,ignore_index = True, sort=False)
            logging.info(f + s + ': Data ETL OK')
        except xlrd.XLRDError:
            logging.error(f + s + ": Data ETL ERROR, Can't find the file!")
            
# 匯出國際航空資料
df_outbound.to_csv('./outbound.csv')

### main loop(local) ###


# 創造初始空dataframe
df_local=pd.DataFrame() 

# 取國際航空新增資料
dl_list=new_list[new_list['type']=='local']



# 整理+合併資料
for f in list(dl_list['title']):
    try:
        df_add=caaStatsMonthly.reshape_local(f)
        df_local=df_local.append(df_add,ignore_index = True, sort=False)
        logging.info(f + ': Data ETL OK')
    except xlrd.XLRDError:
        logging.error(f + ": Data ETL ERROR, Can't find the file!")
            

# 匯出國內航空資料
df_local.to_csv('./local.csv')
