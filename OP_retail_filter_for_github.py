#載入套件
import time
import datetime
import xlwings
import threading
import xlwt
from matplotlib import pyplot as plt
from websocket import create_connection
import json
import winsound

from linebot import LineBotApi
from linebot.models import TextSendMessage

## if test, test_flag = 1
test_flag = 0

#取得當天日期
today = datetime.datetime.today()
the_day = datetime.date.today()

#read excel file
#fill your excel file path
book = xlwings.Book(r'C:\xxxx\xxxx\xxxxx\xxxxx\xxxxx\retail_filter_dde_excel_for_github.xlsx')
book2 = xlwings.Book(r'C:\xxx\xxxxx\HTSAPI3.0_app_VBA_function-N_for_github.xls')

print("start")
sheet = book.sheets[0]
sheet_order = book2.sheets[0]



#散戶定義 < 5口
retail_threshold = 5
#幾秒抓一次data
time_threshold = 60
# 日盤 = 1, 夜盤 = 0
day_or_night = 1
#前日收盤價
price_temp = 9872


#line bot,use your line bot ID and pass word and user_id are line message reciever
channel_access_token = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
user_id1 = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" 
user_id2 = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" 

line_bot_api = LineBotApi(channel_access_token)


## ID imformation for data to cloud, it is optional function
#def CreateCredential():
#    secret = b"xxxxxxxx"
#    import hashlib, hmac
#
#    ts = str(int(time.time()*1000 +10000))
#    sig = hmac.new(secret, ts.encode("utf8"), hashlib.sha256).digest()
#
#    return "ws://xx.xxx.xxx.xx:xxxx?ts="+ts+"&sig="+sig.hex()

def get_time_in_sec(day_or_night):
    time_prev = sheet.range('B3').value
    h = int(time_prev[1:3])
    m = int(time_prev[4:6])
    s = int(time_prev[7:9])
    if(day_or_night == 1):
        time_current = (60*60*h+60*m+s) -(60*60*8+60*45)
    else:
        time_current = (60*60*h + 60*m + s) - (60*60*15)
    return time_current


def get_time_in_ws_format():
    t = sheet.range('B3').value[1:-1]
    h = int(t[0:2])
    if h > 23:
        t = '{:02d}{}'.format((h-24), t[2:])
    td = datetime.datetime.today()
    today_str = str(td.date()).replace('-', '/')
    return '{} {}'.format(today_str, t)



def api_order(order):
    if (order == "open buy"):
        #先將下單函數清掉
        sheet_order.range('E1').value = "none"
        # D4 建倉:O or 平倉:C
        sheet_order.range('D4').value = "O"
        # D5 買:B or 賣S
        sheet_order.range('D5').value = "B"
        # D8 price
        # sheet_order.range('D8').value = sheet.range('B24').value
        order_price = sheet.range('B4').value
        # 下單函數
        sheet_order.range('E1').value = "=FHTSOrder(D1,D2,D3,D4,D5,D6,D7,D8,D9,D10)"
    elif (order == "close buy"):
        # 先將下單函數清掉
        sheet_order.range('E1').value = "none"
        # D4 建倉:O or 平倉:C
        sheet_order.range('D4').value = "C"
        # D5 買:B or 賣S
        sheet_order.range('D5').value = "S"
        # D8 price
        # sheet_order.range('D8').value = sheet.range('B23').value
        order_price = sheet.range('B4').value
        # 下單函數
        sheet_order.range('E1').value = "=FHTSOrder(D1,D2,D3,D4,D5,D6,D7,D8,D9,D10)"
    if (order == "open sell"):
        #先將下單函數清掉
        sheet_order.range('E1').value = "none"
        # D4 建倉:O or 平倉:C
        sheet_order.range('D4').value = "O"
        # D5 買:B or 賣S
        sheet_order.range('D5').value = "S"
        # D8 price
        # sheet_order.range('D8').value = sheet.range('B23').value
        order_price = sheet.range('B4').value
        # 下單函數
        sheet_order.range('E1').value = "=FHTSOrder(D1,D2,D3,D4,D5,D6,D7,D8,D9,D10)"
    elif (order == "close sell"):
        # 先將下單函數清掉
        sheet_order.range('E1').value = "none"
        # D4 建倉:O or 平倉:C
        sheet_order.range('D4').value = "C"
        # D5 買:B or 賣S
        sheet_order.range('D5').value = "B"
        # D8 price
        # sheet_order.range('D8').value = sheet.range('B24').value
        order_price = sheet.range('B4').value
        # 下單函數
        sheet_order.range('E1').value = "=FHTSOrder(D1,D2,D3,D4,D5,D6,D7,D8,D9,D10)"
    return order_price



##define your own strategy
def get_strategy1_data():
####   your strategy start 


####   your strategy stop
    #order and record the order price
    order_price = api_order(strategy)
    # use line bot to have a reminder by sending a line message, for example:
    message = "平倉 價:%d :時間: %s" % (order_price, time_data_in_format[-1])
    line_bot_api.push_message(user_id1, TextSendMessage(text=message))
    line_bot_api.push_message(user_id2, TextSendMessage(text=message))
    return 0

##threading plot
def plot_data( time_data,price_data,day_or_night):
    plt.ion()
    plt.show()
    while(True):
        try:
            dog_data_mtx_p = [0] * (len(dog_data_mtx))
            dog_data_mtx_n = [0] * (len(dog_data_mtx))
            tiger_data_tx_p = [0] * (len(tiger_data_tx))
            tiger_data_tx_n = [0] * (len(tiger_data_tx))
            dog_data_mtx_raw_p = [0] * (len(dog_data_mtx_raw))
            dog_data_mtx_raw_n = [0] * (len(dog_data_mtx_raw))
            tiger_data_tx_raw_p = [0] * (len(tiger_data_tx_raw))
            tiger_data_tx_raw_n = [0] * (len(tiger_data_tx_raw))
            for i in range(0, len(dog_data_mtx)):
                if (dog_data_mtx[i] >= 0):
                    dog_data_mtx_p[i] = dog_data_mtx[i]
                    dog_data_mtx_n[i] = 0
                else:
                    dog_data_mtx_p[i] = 0
                    dog_data_mtx_n[i] = dog_data_mtx[i]
            for i in range(0, len(tiger_data_tx)):
                if (tiger_data_tx[i] >= 0):
                    tiger_data_tx_p[i] = tiger_data_tx[i]
                    tiger_data_tx_n[i] = 0
                else:
                    tiger_data_tx_p[i] = 0
                    tiger_data_tx_n[i] = tiger_data_tx[i]

            for i in range(0, len(dog_data_mtx_raw)):
                if (dog_data_mtx_raw[i] >= 0):
                    dog_data_mtx_raw_p[i] = dog_data_mtx_raw[i]
                    dog_data_mtx_raw_n[i] = 0
                else:
                    dog_data_mtx_raw_p[i] = 0
                    dog_data_mtx_raw_n[i] = -dog_data_mtx_raw[i]
            for i in range(0, len(tiger_data_tx_raw)):
                if (tiger_data_tx_raw[i] >= 0):
                    tiger_data_tx_raw_p[i] = tiger_data_tx_raw[i]
                    tiger_data_tx_raw_n[i] = 0
                else:
                    tiger_data_tx_raw_p[i] = 0
                    tiger_data_tx_raw_n[i] = -tiger_data_tx_raw[i]

            plt.figure(0)
            plt.clf()
            top1 = plt.subplot2grid((30, 18), (0, 0), rowspan=9, colspan=18)
            buttom1 = plt.subplot2grid((30, 18), (9, 0), rowspan=4, colspan=18)
            top2 = plt.subplot2grid((30, 18), (17, 0), rowspan=9, colspan=18)
            buttom2 = plt.subplot2grid((30, 18), (26, 0), rowspan=4, colspan=18)

            top1.bar(time_data, dog_data_mtx_p, fc='r')
            top1.bar(time_data, dog_data_mtx_n, fc='g')
            top1.xaxis.set_visible(False)
            top1.tick_params(axis='y', colors='r')
            if (day_or_night ==1):
                top1.set_xlim(0, 305)
                buttom1.set_xlim(0, 305)
            else:
                top1.set_xlim(0, 805)
                buttom1.set_xlim(0, 805)
            top_price1 = top1.twinx()
            top_price1.plot(time_data, price_data, 'k')
            top1.set_title("small")

            buttom1.bar(time_data, dog_data_mtx_raw_p, fc='r')
            buttom1.bar(time_data, dog_data_mtx_raw_n, fc='g')

            top2.bar(time_data, tiger_data_tx_p, fc='r')
            top2.bar(time_data, tiger_data_tx_n, fc='g')
            top2.xaxis.set_visible(False)
            top2.tick_params(axis='y', colors='r')
            if (day_or_night ==1):
                top2.set_xlim(0, 305)
                buttom2.set_xlim(0, 305)
            else:
                top2.set_xlim(0, 805)
                buttom2.set_xlim(0, 805)
            top_price2 = top2.twinx()
            top_price2.plot(time_data, price_data, 'k')
            top2.set_title("big")

            buttom2.bar(time_data, tiger_data_tx_raw_p, fc='r')
            buttom2.bar(time_data, tiger_data_tx_raw_n, fc='g')
            plt.pause(50)
            # plt.close()
        except ValueError:
            print("plot error")
            # print("len(time)=",len(time_data))
            # print("dog=", len(dog_data_mtx))
            # print("tiger=", len(tiger_data_tx))
            # print("price=", len((price_data)))
            # print("dog_raw=", len(dog_data_mtx_raw))
            # print("tiger_raw=", len(tiger_data_mtx_raw))


#初始化
#時間
time_data=[]
time_data_in_format=[]
time_data.append(0)
if (day_or_night==1):
    time_data_in_format.append("[09:00:00]")
else:
    time_data_in_format.append("[15:00:00]")
#價
price_data=[]
price_data.append(price_temp)

#小台
dog_data_mtx_raw=[]
dog_data_mtx_raw.append(0)
tiger_data_mtx_raw=[]
tiger_data_mtx_raw.append(0)
dog_data_mtx = []
dog_data_mtx.append(0) #散戶
tiger_data_mtx = []  #大戶
tiger_data_mtx.append(0)
dog_buy_mtx_temp = 0
tiger_buy_mtx_temp = 0
dog_sell_mtx_temp = 0
tiger_sell_mtx_temp = 0
dog_mtx_temp = 0
tiger_mtx_temp = 0
#大台
dog_data_tx_raw=[]
dog_data_tx_raw.append(0)
tiger_data_tx_raw=[]
tiger_data_tx_raw.append(0)
dog_data_tx = []  #  散戶
dog_data_tx.append(0)
tiger_data_tx = []  # 大戶
tiger_data_tx.append(0)
dog_buy_tx_temp = 0
tiger_buy_tx_temp = 0
dog_sell_tx_temp = 0
tiger_sell_tx_temp = 0
dog_tx_temp = 0
tiger_tx_temp = 0





#thread for plot
t = threading.Thread(target=plot_data, args=(time_data,price_data,day_or_night))
t.start()
#for cloud data
#ws = create_connection(CreateCredential(), subprotocols=["provider"])




#讀excel第一筆資料
#小台
buy_volume_mtx_prev = sheet.range('B17').value
sell_volume_mtx_prev = sheet.range('B18').value
#大台
buy_volume_tx_prev = sheet.range('E17').value
sell_volume_tx_prev = sheet.range('E18').value


if (day_or_night==1):
    time_data_in_format.append("[09:00:00]")
else:
    time_data_in_format.append("[15:00:00]")

time_prev = (get_time_in_sec(day_or_night)//time_threshold)

## for cloud data
#def sendData(data):
#    global ws
#    j = json.dumps(data)
#    print(j)
#    ws.send(j)

#def clearData():
#    global ws
#    ws.send("clear")

#if(test_flag ==0):
#    clearData()


while(True):
    #先讀時間及總量 判斷有無更新 有新資料才將資料塞進list
    time_temp = get_time_in_sec(day_or_night)
    time_temp_in_format = get_time_in_ws_format()
    if(time_temp >= 0):#開盤才更使更新
        #小台買賣量
        buy_volume_mtx = sheet.range('B17').value
        sell_volume_mtx = sheet.range('B18').value
        # 大台買賣量
        buy_volume_tx = sheet.range('E17').value
        sell_volume_tx = sheet.range('E18').value
        #小台買進或賣出成交量有變動 => 更新資料
        if ((buy_volume_mtx - buy_volume_mtx_prev) > 0 or (sell_volume_mtx - sell_volume_mtx_prev) > 0 or (buy_volume_tx - buy_volume_tx_prev) > 0 or (sell_volume_tx - sell_volume_tx_prev) > 0):
            price_temp = sheet.range('B4').value
            #時間增加一分鐘 才更新資料 否則一直累加
            if ( ((int) (time_temp / time_threshold) != time_prev)):
            #小台判斷買為散戶或大戶
                if ((sell_volume_mtx - sell_volume_mtx_prev) <= retail_threshold):
                    dog_buy_mtx_temp = dog_buy_mtx_temp + (sell_volume_mtx - sell_volume_mtx_prev)
                else:
                    tiger_buy_mtx_temp = tiger_buy_mtx_temp +  (sell_volume_mtx - sell_volume_mtx_prev)
                # 小台判斷賣為散戶或大戶
                if ((buy_volume_mtx - buy_volume_mtx_prev) <= retail_threshold):
                    dog_sell_mtx_temp = dog_sell_mtx_temp + (buy_volume_mtx - buy_volume_mtx_prev)
                else:
                    tiger_sell_mtx_temp = tiger_sell_mtx_temp +(buy_volume_mtx - buy_volume_mtx_prev)

                # 大台判斷買為散戶或大戶
                if ((sell_volume_tx - sell_volume_tx_prev) <= retail_threshold):
                    dog_buy_tx_temp = dog_buy_tx_temp + (sell_volume_tx - sell_volume_tx_prev)
                else:
                    tiger_buy_tx_temp = tiger_buy_tx_temp + (sell_volume_tx - sell_volume_tx_prev)
                # 大台判斷賣為散戶或大戶
                if ((buy_volume_tx - buy_volume_tx_prev) <= retail_threshold):
                    dog_sell_tx_temp = dog_sell_tx_temp + (buy_volume_tx - buy_volume_tx_prev)
                else:
                    tiger_sell_tx_temp = tiger_sell_tx_temp + (buy_volume_tx - buy_volume_tx_prev)

                time_data.append( time_temp // time_threshold)
                time_data_in_format.append(time_temp_in_format)
                price_data.append(price_temp)
                #小台累積資料
                dog_data_mtx.append(dog_data_mtx[-1] + dog_buy_mtx_temp - dog_sell_mtx_temp )  # 小台散戶累積資料
                tiger_data_mtx.append(tiger_data_mtx[-1] + tiger_buy_mtx_temp -tiger_sell_mtx_temp)  # 小台大戶累積資料
                dog_data_mtx_raw.append(dog_buy_mtx_temp - dog_sell_mtx_temp )
                tiger_data_mtx_raw.append(tiger_buy_mtx_temp -tiger_sell_mtx_temp)
                #大台累積資料
                dog_data_tx.append(dog_data_tx[-1] + dog_buy_tx_temp - dog_sell_tx_temp)  # 散戶累積資料
                tiger_data_tx.append(tiger_data_tx[-1] + tiger_buy_tx_temp - tiger_sell_tx_temp)  # 大戶累積資料
                dog_data_tx_raw.append(dog_buy_tx_temp - dog_sell_tx_temp)
                tiger_data_tx_raw.append(tiger_buy_tx_temp - tiger_sell_tx_temp)
                #策略訊號判斷並且下單
                get_strategy1_data()

                #data for internet, optional function
                #if (test_flag == 0):
                #    ws = create_connection(CreateCredential(), subprotocols=["provider"])
                #    internet_data = [time_temp_in_format,price_temp,dog_data_mtx[-1],tiger_data_tx[-1],dog_data_mtx_raw[-1],tiger_data_tx_raw[-1],strategy1_data[-1]]
                #    sendData(internet_data)
                #    ws.close()

                # print("time=", time_data)
                # print("dog_mtx=", dog_data_mtx)
                # print("tiger_tx=", tiger_data_tx)
                # print("dog_mtx_raw=", dog_data_mtx_raw)
                # print("tiger_tx_raw=", tiger_data_tx_raw)
                # print("price=", price_data)
                # print("strategy=",strategy1_data)

                tiger_sell_mtx_temp = 0
                tiger_buy_mtx_temp = 0
                dog_buy_mtx_temp = 0
                dog_sell_mtx_temp = 0

                tiger_sell_tx_temp = 0
                tiger_buy_tx_temp = 0
                dog_buy_tx_temp = 0
                dog_sell_tx_temp = 0

            else:
                #price_data[-1] = price_temp 前一筆的data不改
                #小台判斷買為散戶或大戶
                if ((sell_volume_mtx - sell_volume_mtx_prev) <= retail_threshold):
                    dog_buy_mtx_temp = dog_buy_mtx_temp + (sell_volume_mtx - sell_volume_mtx_prev)
                else:
                    tiger_buy_mtx_temp = tiger_buy_mtx_temp + (sell_volume_mtx - sell_volume_mtx_prev)
                # 小台判斷賣為散戶或大戶
                if ((buy_volume_mtx - buy_volume_mtx_prev) < retail_threshold):
                    dog_sell_mtx_temp = dog_sell_mtx_temp + (buy_volume_mtx - buy_volume_mtx_prev)
                else:
                    tiger_sell_mtx_temp = tiger_sell_mtx_temp + (buy_volume_mtx - buy_volume_mtx_prev)

                # 大台判斷買為散戶或大戶
                if ((sell_volume_tx - sell_volume_tx_prev) <= retail_threshold):
                    dog_buy_tx_temp = dog_buy_tx_temp + (sell_volume_tx - sell_volume_tx_prev)
                else:
                    tiger_buy_tx_temp = tiger_buy_tx_temp + (sell_volume_tx - sell_volume_tx_prev)
                # 大台判斷賣為散戶或大戶
                if ((buy_volume_tx - buy_volume_tx_prev) < retail_threshold):
                    dog_sell_tx_temp = dog_sell_tx_temp + (buy_volume_tx - buy_volume_tx_prev)
                else:
                    tiger_sell_tx_temp = tiger_sell_tx_temp + (buy_volume_tx - buy_volume_tx_prev)

            buy_volume_mtx_prev = buy_volume_mtx
            sell_volume_mtx_prev = sell_volume_mtx
            buy_volume_tx_prev = buy_volume_tx
            sell_volume_tx_prev = sell_volume_tx
            time_prev = (time_temp//time_threshold)

            # print("time=", time_data)
            # print("dog_mtx=", dog_data_mtx)
            # print("tiger_tx=", tiger_data_tx)
            # print("dog_mtx_raw=", dog_data_mtx_raw)
            # print("tiger_tx_raw=", tiger_data_tx_raw)
            # print("price=", price_data)
            # print("strategy=", strategy1_data)

    if (day_or_night==1):
        if (time_temp==17999):
            print("write")
            wb = xlwt.Workbook()
            sheet1 = wb.add_sheet('sheet1')
            sheet1.write(0,0,'time')
            sheet1.write(0,1,'price')
            sheet1.write(0, 2, 'retail_mtx')
            sheet1.write(0, 3, 'institution_mtx')
            sheet1.write(0, 4, 'retail_tx')
            sheet1.write(0, 5, 'institution_tx')
            sheet1.write(0, 6, 'retail_mtx_raw')
            sheet1.write(0, 7, 'institution_mtx_raw')
            sheet1.write(0, 8, 'retail_tx_raw')
            sheet1.write(0, 9, 'institution_tx_raw')
            for i in range(0,len(time_data)):
                sheet1.write(i + 1, 0 , time_data[i])
                sheet1.write(i + 1, 1 , price_data[i])
                sheet1.write(i + 1, 2 , dog_data_mtx[i])
                sheet1.write(i + 1, 3, tiger_data_mtx[i])
                sheet1.write(i + 1, 4, dog_data_tx[i])
                sheet1.write(i + 1, 5, tiger_data_tx[i])
                sheet1.write(i + 1, 6, dog_data_mtx_raw[i])
                sheet1.write(i + 1, 7, tiger_data_mtx_raw[i])
                sheet1.write(i + 1, 8, dog_data_tx_raw[i])
                sheet1.write(i + 1, 9, tiger_data_tx_raw[i])
            wb.save("retail_%s_day.xls" % the_day)
            print("finish")
            break
    else:
        if ( 50399-time_temp <= 20):
            print("write")
            wb = xlwt.Workbook()
            sheet1 = wb.add_sheet('sheet1')
            sheet1.write(0,0,'time')
            sheet1.write(0,1,'price')
            sheet1.write(0, 2, 'retail_mtx')
            sheet1.write(0, 3, 'institution_mtx')
            sheet1.write(0, 4, 'retail_tx')
            sheet1.write(0, 5, 'institution_tx')
            sheet1.write(0, 6, 'retail_mtx_raw')
            sheet1.write(0, 7, 'institution_mtx_raw')
            sheet1.write(0, 8, 'retail_tx_raw')
            sheet1.write(0, 9, 'institution_tx_raw')
            for i in range(0,len(time_data)):
                sheet1.write(i + 1, 0 , time_data[i])
                sheet1.write(i + 1, 1 , price_data[i])
                sheet1.write(i + 1, 2 , dog_data_mtx[i])
                sheet1.write(i + 1, 3, tiger_data_mtx[i])
                sheet1.write(i + 1, 4, dog_data_tx[i])
                sheet1.write(i + 1, 5, tiger_data_tx[i])
                sheet1.write(i + 1, 6, dog_data_mtx_raw[i])
                sheet1.write(i + 1, 7, tiger_data_mtx_raw[i])
                sheet1.write(i + 1, 8, dog_data_tx_raw[i])
                sheet1.write(i + 1, 9, tiger_data_tx_raw[i])
            wb.save("retail_%s_night.xls" % the_day)
            print("finish")
            break


