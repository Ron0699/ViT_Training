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
excel_file = 'C:\\Users\\王庭軒\\Desktop\\RON_Note2023\\Python\\ViT_Stock\\List2.xls'
TrainPeriod = 20
Label_Trend = 1

#Label txt
fLabel = open('Label2.txt', 'w')

workbook = xlrd.open_workbook(excel_file, formatting_info=True)
sheet1 = workbook.sheets()[0]
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
    
    for i in range(0, len(df)-TrainPeriod-Label_Trend+1):
        dfS = df[i:i+TrainPeriod]
        AllData = []
        HighLowData = []
        BottomValue = []
        ColorBar = []
        for length in range(0, len(dfS.index)):
            if dfS['Open'][length] <= dfS['Close'][length]:
                AllData.append(dfS['Close'][length]-dfS['Open'][length])
                HighLowData.append(dfS['High'][length]-dfS['Low'][length])
                BottomValue.append(dfS['Open'][length])
                ColorBar.append('r')
            else:
                AllData.append(dfS['Open'][length]-dfS['Close'][length])
                HighLowData.append(dfS['High'][length]-dfS['Low'][length])
                BottomValue.append(dfS['Close'][length])
                ColorBar.append('g')
    
        #plt.plot(tw_time, close, color=(60/255,100/255,230/255))
        plt.bar(dfS.index, HighLowData, bottom=dfS['Low'], color='cyan', width=0.6)
        plt.bar(dfS.index, AllData, bottom=BottomValue, color=ColorBar, width=0.6)
        
        #Trend
        # dfT = df[i+TrainPeriod:i+TrainPeriod+Label_Trend]
        # dfT['date_id'] = ((dfT.index.date - dfT.index.date.min())).astype('timedelta64[D]')
        # dfT['date_id'] = dfT['date_id'].dt.days + 1
        # reg = linregress(dfT['date_id'], dfT['Close'])
        # '''LowLine = reg[0] * dfT['date_id'] + reg[1]
        # df_temp = dfT[dfT['Close'] < LowLine]
        # df_temp['date_id'] = ((df_temp.index.date - df_temp.index.date.min())).astype('timedelta64[D]')
        # df_temp['date_id'] = df_temp['date_id'].dt.days + 1
        # while len(df_temp) >= 5 :
        #     reg = linregress(df_temp['date_id'], df_temp['Close'])
        #     LowLine = reg[1] + reg[0] * df['date_id']
        #     df_temp = df[df['Close'] < LowLine]
        # plt.plot(dfT.index, reg[0] * dfT['date_id'] + reg[1], 'g')'''
        # if reg[0]<=-0.7:
        #     LabelNumber = 0
        # elif reg[0]>-0.7 and reg[0]<=-0.4:
        #     LabelNumber = 1
        # elif reg[0]>-0.4 and reg[0]<=-0.2:
        #     LabelNumber = 2
        # elif reg[0]>-0.2 and reg[0]<=0.05:
        #     LabelNumber = 3
        # elif reg[0]>0.05 and reg[0]<=0.2:
        #     LabelNumber = 4
        # elif reg[0]>0.2 and reg[0]<=0.35:
        #     LabelNumber = 5
        # elif reg[0]>0.35 and reg[0]<=0.5:
        #     LabelNumber = 6
        # elif reg[0]>0.5 and reg[0]<=0.6:
        #     LabelNumber = 7
        # elif reg[0]>0.6 and reg[0]<=0.7:
        #     LabelNumber = 8
        # elif reg[0]>0.7:
        #     LabelNumber = 9
        # fLabel.write(StockName + '_' + str(df.index[i+TrainPeriod].date()) + '.jpg'+' ' + str(LabelNumber) + '\n')
        
        #Williams' %R
        WilliamsR = talib.WILLR(dfS['High'], dfS['Low'], dfS['Close'], TrainPeriod)
        WR = WilliamsR[-1]*-1
        #print(WR)
        if WR<10 and df.Close[i+TrainPeriod-2]<df.Close[i+TrainPeriod-1] and df.Close[i+TrainPeriod]<df.Close[i+TrainPeriod-1]:
            LabelNumber = 0
        elif WR<35:
            LabelNumber = 1
        elif WR<65 and WR>=35:
            LabelNumber = 2
        elif WR<95 and WR>=65:
            LabelNumber = 3
        elif WR>95 and df.Close[i+TrainPeriod-2]>df.Close[i+TrainPeriod-1] and df.Close[i+TrainPeriod]>df.Close[i+TrainPeriod-1]:
            LabelNumber = 5
        elif WR>=95:
            LabelNumber = 4
        fLabel.write(StockName + '_' + str(df.index[i+TrainPeriod-1].date()) + '.jpg'+' ' + str(LabelNumber) + '\n')
        
        FolderName = workbook.sheet_names()[0]
        if not os.path.exists('Data2\\'):
            os.makedirs('Data2\\')
        plt.axis('off')
        plt.savefig('Data2\\' + StockName + '_' + str(df.index[i+TrainPeriod-1].date()) + '.jpg', dpi=40)
        plt.close()
    
    time.sleep(random.randrange(15)/10)

fLabel.close()