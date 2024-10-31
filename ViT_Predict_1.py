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

TimePeriod = '2y' #1y 6mo 5d
TrainPeriod = 20
Label_Trend = 1

#prepare the data
num_classes = 6
input_shape = (32, 32, 3)

excel_file = 'C:\\Users\\王庭軒\\Desktop\\RON_Note2023\\Python\\ViT_Stock\\List2.xls'
path_images = 'C:/Users/王庭軒/Desktop/RON_Note2023/Python/ViT_Stock/Data2/'
Label_dir = 'C:/Users/王庭軒/Desktop/RON_Note2023/Python/ViT_Stock/Label2.txt'
Label_list = open(Label_dir, mode='r')

#load all file name
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
targets = []
for i in range(0, len(all_file_name)):
    image = keras.utils.load_img(path_images + all_file_name[i])
    img_size_ori = image.size
    image = image.resize((image_size, image_size), Image.BICUBIC)
    Resize_Ratio_w = image_size/img_size_ori[0]
    Resize_Ratio_h = image_size/img_size_ori[1]
    images.append(keras.utils.img_to_array(image))
    Axis = Label_list.readline().split(' ')
    targets.append(int(Axis[1]))
ImgRatio = int(len(images) * 1)
(x_test), (y_test) = ( np.asarray(images[:ImgRatio]), np.asarray(targets[:ImgRatio]) )

#Load the Model
#filename = 'ViT'
filename = 'ViT_SMOTE'
ViT_Model = tf.keras.models.load_model(filename)

_, accuracy, top_5_accuracy = ViT_Model.evaluate(x_test, y_test)
print(f"Test accuracy: {round(accuracy * 100, 2)}%")
print(f"Test top 5 accuracy: {round(top_5_accuracy * 100, 2)}%")

preds = ViT_Model.predict(x_test)
Ans = []
for i in range(0, len(preds)):
    Ans.append(np.argmax(preds[i][:]))

#PR Curve
PRy_test = np.zeros(len(Ans), dtype = int)
PRy_score = preds[:, 5]
for i in range(len(Ans)):
    PRy_test[i]=1 if Ans[i]==y_test[i] and y_test[i]==5 else 0
precision, recall, thresholds = precision_recall_curve(PRy_test, PRy_score)
fig, ax = plt.subplots()
ax.plot(recall, precision, color='g')
ax.set_title('Precision-Recall Curve')
ax.set_ylabel('Precision')
ax.set_xlabel('Recall')
plt.show()

#F1 score
from sklearn.metrics import f1_score
print(f1_score(y_test, Ans, average=None))

# #Count Label
# n0=n1=n2=n3=n4=n5=0
# for i in range(len(Ans)):
#     if Ans[i]==0:
#         n0+=1
#     elif Ans[i]==1:
#         n1+=1
#     elif Ans[i]==2:
#         n2+=1
#     elif Ans[i]==3:
#         n3+=1
#     elif Ans[i]==4:
#         n4+=1
#     elif Ans[i]==5:
#         n5+=1
# print('0: ' + str(n0))
# print('1: ' + str(n1))
# print('2: ' + str(n2))
# print('3: ' + str(n3))
# print('4: ' + str(n4))
# print('5: ' + str(n5) + '\n')

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
    
    for i in range(0, len(df)-TrainPeriod-Label_Trend+1):
        if int(y_test[iSeries]) == 5: #or int(y_test[iSeries]) == 4:
            plt.plot(df.index[i+TrainPeriod-1], df.Close[i+TrainPeriod-1], '*r')
        if Ans[iSeries]==5: #or Ans[iSeries] == 4:
            plt.plot(df.index[i+TrainPeriod-1], df.Close[i+TrainPeriod-1], '.b')
        iSeries += 1
        
    FolderName = workbook.sheet_names()[0]
    if not os.path.exists('DataVerify2\\'):
        os.makedirs('DataVerify2\\')
    plt.savefig('DataVerify2\\' + StockName + '_' + '.jpg', dpi=800)
    plt.close()
    
    time.sleep(random.randrange(15)/10)

Label_list.close()