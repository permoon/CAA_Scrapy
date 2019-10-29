# -*- coding: utf-8 -*-

import pandas as pd

# 定義檢查欄位是否包含中文字的函數
def is_contain_chinese(check_str):
    for ch in check_str:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False


def caa_reshape(file,sheet, path):
    df_output=pd.read_excel(path + '/' + file ,sheet_name=sheet)
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
    return df_output




