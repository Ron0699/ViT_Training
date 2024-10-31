import requests, datetime, xlrd, os, random, time, talib, warnings
import pandas as pd
import pandas_datareader as web
import matplotlib.pyplot as plt
import numpy as np
import yfinance as yf
from scipy.stats import linregress
from scipy import signal
from sklearn import preprocessing

TimePeriod = '2y' #1y 6mo 5d
excel_file = 'C:\\Users\\王庭軒\\Desktop\\RON_Note2023\\Python\\ViT_Stock\\List.xls'
TrainPeriod = 20
Label_Trend = 1

#Label txt
Label_dir = 'C:/Users/王庭軒/Desktop/RON_Note2023/Python/ViT_Stock/Label.txt'
Label_list = open(Label_dir, mode='r')

workbook = xlrd.open_workbook(excel_file, formatting_info=True)
sheet1 = workbook.sheets()[0]
targets = []
for s in range(0, sheet1.nrows):
    StockName = sheet1.cell(s, 0).value.strip()
    warnings.filterwarnings('ignore')

    try:
        df = yf.Ticker(StockName).history(period=TimePeriod, interval='1d')
    except:
        continue
    
    #Skip None Value
    if df['Close'][-1] == None:
        continue
    
    AllData = []
    HighLowData = []
    BottomValue = []
    ColorBar = []
    for length in range(0, len(df.index)):
        if df['Open'][length] <= df['Close'][length]:
            AllData.append(df['Close'][length]-df['Open'][length])
            HighLowData.append(df['High'][length]-df['Low'][length])
            BottomValue.append(df['Open'][length])
            ColorBar.append('r')
        else:
            AllData.append(df['Open'][length]-df['Close'][length])
            HighLowData.append(df['High'][length]-df['Low'][length])
            BottomValue.append(df['Close'][length])
            ColorBar.append('g')

    #plt.plot(tw_time, close, color=(60/255,100/255,230/255))
    plt.bar(df.index, HighLowData, bottom=df['Low'], color='cyan', width=0.6)
    plt.bar(df.index, AllData, bottom=BottomValue, color=ColorBar, width=0.6)
    plt.title(StockName)
    plt.xlabel('Time')
    plt.ylabel('Price')
    plt.grid(True)
    
    for i in range(0, len(df)-TrainPeriod-Label_Trend+1):
        Axis = Label_list.readline().split(' ')
        targets.append(int(Axis[1]))
        if int(Axis[1]) == 5: #or int(Axis[1]) == 4:
            plt.plot(df.index[i+TrainPeriod-1], df.Close[i+TrainPeriod-1], '*r')
            #print(df.index[i+TrainPeriod])
        
    FolderName = workbook.sheet_names()[0]
    if not os.path.exists('DataVerify\\'):
        os.makedirs('DataVerify\\')
    plt.savefig('DataVerify\\' + StockName + '_' + '.jpg', dpi=800)
    plt.close()
    
    time.sleep(random.randrange(15)/10)

Label_list.close()