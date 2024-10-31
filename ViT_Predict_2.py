import requests, datetime, xlrd, os, random, time, talib, warnings
import numpy as np
import pandas as pd
import pandas_datareader as web
import tensorflow as tf
import yfinance as yf
from tensorflow import keras
from tensorflow.keras import layers
from PIL import Image
from scipy.stats import linregress
from scipy import signal
from sklearn import preprocessing
import tensorflow_addons as tfa
import matplotlib.pyplot as plt
from sklearn.metrics import precision_recall_curve

TimePeriod = '3mo' #1y 6mo 5d
excel_file = 'C:\\Users\\王庭軒\\Desktop\\RON_Note2023\\Python\\ViT_Stock\\ListPredict.xls'
TrainPeriod = 20

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
    
    for i in range(0, len(df)-TrainPeriod+1):
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

        FolderName = workbook.sheet_names()[0]
        if not os.path.exists('DataPredict\\'):
            os.makedirs('DataPredict\\')
        plt.axis('off')
        plt.savefig('DataPredict\\' + StockName + '_' + str(df.index[i+TrainPeriod-1].date()) + '.jpg', dpi=40)
        plt.close()
    
    time.sleep(random.randrange(15)/10)

#prepare the data
num_classes = 6
input_shape = (32, 32, 3)

#load all file name
path_images = 'C:/Users/王庭軒/Desktop/RON_Note2023/Python/ViT_Stock/DataPredict/'
all_file_name = []
workbook = xlrd.open_workbook(excel_file, formatting_info=True)
sheet1 = workbook.sheets()[0]
for s in range(0, sheet1.nrows):
    StockName = sheet1.cell(s, 0).value.strip()
    for names in os.listdir(path_images):
        if names.startswith(StockName + '_') and names.endswith(".jpg"):
            all_file_name.append(names)

#prepare the data
image_size = 32
images = []
for i in range(0, len(all_file_name)):
    image = keras.utils.load_img(path_images + all_file_name[i])
    img_size_ori = image.size
    image = image.resize((image_size, image_size), Image.BICUBIC)
    Resize_Ratio_w = image_size/img_size_ori[0]
    Resize_Ratio_h = image_size/img_size_ori[1]
    images.append(keras.utils.img_to_array(image))
ImgRatio = int(len(images) * 1)
x_test= np.asarray(images[:ImgRatio])

#Load the Model
#filename = 'ViT'
filename = 'ViT_SMOTE'
ViT_Model = tf.keras.models.load_model(filename)

preds = ViT_Model.predict(x_test)
Ans = []
for i in range(0, len(preds)):
    Ans.append(np.argmax(preds[i][:]))

#save/print the result here
workbook = xlrd.open_workbook(excel_file, formatting_info=True)
sheet1 = workbook.sheets()[0]
iSeries = 0
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
    plt.xticks(rotation=20)
    plt.ylabel('Price')
    plt.grid(True)
    
    for i in range(0, len(df)-TrainPeriod+1):
        if Ans[iSeries]==5: #or Ans[iSeries] == 4:
            plt.plot(df.index[i+TrainPeriod-1], df.Close[i+TrainPeriod-1], '*r')
        iSeries += 1
        
    FolderName = workbook.sheet_names()[0]
    if not os.path.exists('DataVerifyPredict\\'):
        os.makedirs('DataVerifyPredict\\')
    plt.savefig('DataVerifyPredict\\' + StockName + '_' + '.jpg', dpi=800)
    plt.close()
    
    time.sleep(random.randrange(15)/10)