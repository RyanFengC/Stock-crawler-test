# -*- coding: utf-8 -*-
"""
Created on Tue Jul 20 23:30:40 2021

@author: user
"""


import pandas_datareader as web
import matplotlib.pyplot as plt   # 資料視覺化套件
import pandas as pd
import datetime,time

# .............................下載股價資料................................................
def Stock_data_download(stock_ID,start_time,end_time,Sdata):
    df = web.DataReader(name=stock_ID +'.TW', data_source='yahoo', start=start_time, end=end_time) #name為股票代號名稱 start、end為資料下載期間
    df.insert(loc=0,column='Date',value=pd.to_datetime(df.index).dt.date)#將取得的資料日期存入表格中，並只存入日期，去除時間
    if df['Date']==start_time:
        print('NO New Data')
    else:
        df = Sdata.append(df)#新增最新資料
        df.to_excel(stock_ID +'TW.xlsx',index=False)
    return(df)
       

def Stock_plot(plotdata,plot_name):

    plt.style.use("ggplot")               # 使用ggplot主題樣式
    #畫第一條線，plt.plot(x, y, c)參數分別為x軸資料、y軸資料及線顏色 = 紅色
    plt.plot(plotdata['Date'], plotdata['Low'],c = "r")  
    #畫第二條線，plt.plot(x, y, c)參數分別為x軸資料、y軸資料、線顏色 = 綠色及線型式 = -.
    plt.plot(plotdata['Date'], plotdata['High'], "g-.")
    
    # 設定圖例，參數為標籤、位置
    plt.legend(labels=['Low','High'], loc = 'best')
    plt.xlabel("Date", fontweight = "bold")                # 設定x軸標題及粗體
    plt.ylabel("Price", fontweight = "bold")    # 設定y軸標題及粗體
    plt.title(plot_name+" Stock curve", fontsize = 15, fontweight = "bold", y = 1.1)   # 設定標題、文字大小、粗體及位置
    plt.xticks(rotation=45)   # 將x軸數字旋轉45度，避免文字重疊
    plt.show()
      

Focus_stock=pd.read_excel('Focus_stock.xlsx')#載入關注股票的EXCEL檔案
Focus_stock_name=Focus_stock['ID']
for row in range( len(Focus_stock_name)):
    filename = str(Focus_stock_name[row]).zfill(4)#zfill表示取出的代碼不足4碼的前面補滿四位數
    try:#判斷是否已經存在該股票資料
        stock_data = pd.read_excel(filename+'TW.xlsx')
        stock_data_latest_date=stock_data['Date'][len(stock_data)-1]#取出資料最後一筆之時間
        start_day=stock_data_latest_date + datetime.timedelta(days=1)#將最後一筆時間加一天做為起始天
        print(filename+"TW.xlsx 載入數據成功!")
    except:
        stock_data=pd.DataFrame()#建立空的資料空間，若不加則無法加入最新資料
        start_day=datetime.date(2001,1,1)
        print(filename+"TW.xlsx 無此資料!")  

    today=datetime.date.today()
    if  start_day-datetime.timedelta(days=1)!=today:
        #判斷數據裡的資料是否為今天，若不是則更新
        try:
            stock_data=Stock_data_download(filename,start_day,today,stock_data)
        except:
            print('沒有最新資料')
        print(filename+'TW :'+'sucessful update to latest data') 
        time.sleep(5)#休息怕程式抓太快被ban
    else:
        print(filename+'TW :'+'latest data')   
    
    try:
        Stock_plot(stock_data,filename)
    except:
        print(filename+"TW.xlsx 找不到此檔案!")

