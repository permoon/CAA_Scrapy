# -*- coding: utf-8 -*-
import pandas as pd
from bs4 import BeautifulSoup
import requests
from datetime import datetime
import logging
def scrapyInit(keyword='下',stop=1):
    link=[]
    title=[]
    # 取得有效的檔案名稱及連結
    for i in range(1,stop+1):
        url='https://www.caa.gov.tw/StatisticsYearMonthFile.aspx?a=1091&lang=9&p=' + str(i)
        source = requests.get(url,verify=False)
        print(url)
        check = source.text
        if check.find('查無資料')>-1:
            break
        soup = BeautifulSoup(source.text, 'html.parser')
        a_tags = soup.find_all('a', class_="download-filebase")     
        for i in a_tags:
            link.append(i.get('href'))
            title.append(i.get('title'))
    # 創建下載清單df
    df=pd.DataFrame({'link':link,'title':title})
    df=df[df['title'].str.contains(keyword) & df['title'].str.contains("xls")].reset_index(drop=True)   
    # 取代連結中的無用字元".."
    def replace_c(s):
        return s.replace('..','')    
    df['link']=df['link'].apply(replace_c)
    return df
 
#init=scrapyInit(keyword='下',stop=40)
#init.to_csv('./src/caaStatsMonthly/_filelist_outbound.csv', index=0)
#init=scrapyInit(keyword='上',stop=40)
#init.to_csv('./src/caaStatsMonthly/_filelist_local.csv', index=0)

def scrapyUpdate(key_word='下'):
    if key_word=='下':        
        exists=pd.read_csv('./src/caaStatsMonthly/_filelist_outbound.csv')
        new=scrapyInit(keyword='下',stop=1)
        exists.update(new)
        exists.to_csv('./src/caaStatsMonthly/_filelist_outbound.csv', index=0)
    elif key_word=='上':
        exists=pd.read_csv('./src/caaStatsMonthly/_filelist_local.csv')
        new=scrapyInit(keyword='上',stop=1)
        exists.update(new)
        exists.to_csv('./src/caaStatsMonthly/_filelist_local.csv', index=0)

####


# 定義檢查欄位是否包含中文字的函數
def is_contain_chinese(check_str):
    for ch in check_str:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False

# 國際航班parser+reshape
def reshape(file,sheet):
    df_output=pd.read_excel('./src/caaStatsMonthly/outbound/'+ file ,sheet_name=sheet)
    # rename columns
    df_output.columns=['sort','DESTINATION', 'AIRLINE', 'AIRPLANE_TOTAL', 
                       'SEAT_TOTAL', 'PEOPLE_TOTAL', 'OCCUPANCYRATE_TOTAL', 
                       'AIRPLANE_INBOUND', 'SEAT_INBOUND', 'PEOPLE_INBOUND', 
                       'OCCUPANCYRATE_INBOUND', 'AIRPLANE_OUTBOUND', 'SEAT_OUTBOUND', 
                       'PEOPLE_OUTBOUND', 'OCCUPANCYRATE_OUTBOUND']
    # 轉換columns類別
    df_output['sort']=df_output['sort'].astype(str)
    df_output['DESTINATION']=df_output['DESTINATION'].astype(str)
    # 檢查第一個欄位是否包含中文
    df_output['check']=df_output['sort'].apply(is_contain_chinese)
    # 假若包含中文，轉換第一個欄位的值為nan
    for i in range(len(df_output)):
        if df_output['check'][i]==True:
            df_output['sort'][i]='nan'        
    # 排除不需要的資料行
    df_output=df_output[(df_output.sort!='nan') | 
            ((df_output.DESTINATION!='nan') & (df_output.DESTINATION!='航線')&(df_output.DESTINATION!='總計'))].reset_index(drop=True)
    # 填補機場資料
    for i in range(1,len(df_output)):
        if df_output['sort'][i]!='nan':
            df_output['DESTINATION'][i]=(df_output['DESTINATION'][i-1])
    # 排除不需要的資料行        
    df_output=df_output[df_output['sort']!='nan']
    df_output=df_output.drop(columns=['check','sort'])
    # 補日期
    df_output['DATA_YEAR']=int(file[:file.find('年')])
    df_output['DATA_MONTH']=int(file[file.find('年')+1:file.find('月')])
    # 補出發機場
    df_output['START_AIRPORT']=sheet
    # 新增timestamp
    df_output['CREATE_DATE']=datetime.today().strftime("%Y%m%d")
    # 設定資料型別
    c_list=['AIRPLANE_TOTAL', 'SEAT_TOTAL', 'PEOPLE_TOTAL', 
            'AIRPLANE_INBOUND', 'SEAT_INBOUND', 'PEOPLE_INBOUND', 
            'AIRPLANE_OUTBOUND', 'SEAT_OUTBOUND', 
            'PEOPLE_OUTBOUND','DATA_YEAR','DATA_MONTH']
    r_list=['OCCUPANCYRATE_TOTAL','OCCUPANCYRATE_INBOUND','OCCUPANCYRATE_OUTBOUND']
    for i in c_list:
        df_output[i]=df_output[i].astype(int)
    for i in r_list:
        df_output[i]=(df_output[i].astype(float)).round(2)
    return df_output

# 國內航班parser+reshape
def reshape_local(file):
    # 讀取資料
    df_output=pd.read_excel('./src/caaStatsMonthly/local/'+file,sheet_name='28')
    # 移除無用欄位
    df_output=df_output[(pd.notnull(df_output['Unnamed: 1'])) & 
                        (df_output['Unnamed: 1']!='總　計') & 
                        (df_output['Unnamed: 1']!='航線')].reset_index(drop=True)    
    # 航線清單
    line_list=df_output['Unnamed: 1'][1:].reset_index(drop=True)
    # 航司清單+資料表中位置
    airline_list=['德安','遠東','立榮','華信']
    airline_position=list(df_output.iloc[0])
    # 創建產出資料表
    df_final=pd.DataFrame({'AIRPLANE_TOTAL':[],
              'SEAT_TOTAL':[],
              'PEOPLE_TOTAL':[],
              'OCCUPANCYRATE_TOTAL':[]})
    # 迴圈找出所有航空公司的資料+併資料
    for i in airline_list:
        try:
            add=df_output.iloc[1:,airline_position.index(i):airline_position.index(i)+4].reset_index(drop=True)
            add.columns=['AIRPLANE_TOTAL','SEAT_TOTAL','PEOPLE_TOTAL','OCCUPANCYRATE_TOTAL']
            add['DESTINATION']=line_list
            add['AIRLINE']=i
            df_final=df_final.append(add)
        except:
            logging.error('NOT Found '+ i + 'in the file')
    # 補日期
    df_final['DATA_YEAR']=int(file[:file.find('年')])
    df_final['DATA_MONTH']=int(file[file.find('年')+1:file.find('月')])
    # 新增timestamp
    df_final['CREATE_DATE']=datetime.today().strftime("%Y%m%d")
    # 設定資料型別
    c_list=['AIRPLANE_TOTAL', 'SEAT_TOTAL', 'PEOPLE_TOTAL']
    r_list=['OCCUPANCYRATE_TOTAL']
    for i in c_list:
        df_final[i]=df_final[i].astype(int)
    for i in r_list:
        df_final[i]=(df_final[i].astype(float)).round(2)
    return df_final