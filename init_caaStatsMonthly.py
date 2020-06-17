# -*- coding: utf-8 -*-
"""
Created on Wed Jun 17 15:56:42 2020

@author: hectorhuang
"""
import func.caaStatsMonthly as caaStatsMonthly

init_outbound=caaStatsMonthly.scrapyInit(keyword='下',stop=40)
init_outbound.to_csv('./src/caaStatsMonthly/_filelist_outbound.csv', index=0)
init_local=caaStatsMonthly.scrapyInit(keyword='上',stop=40)
init_local.to_csv('./src/caaStatsMonthly/_filelist_local.csv', index=0)