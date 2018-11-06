# Taiwan futures placing order tool and retail analyzer
## Instruction
It is a project that shows real-time trend of retail and institution trader.
Depend on this trend, we can define our own strategy and place an order automatically by programming.
Once an order was placed, it also provide a reminder that sent a message to the users by line bot.   

## DDE
We need DDE(dynamic data exchange) that securities firm provided so that we can use python to excess excel to get real time data. We drag DDE data to excel, the excel data and DDE data will be synchronize, then we can process data for trend.   
DDE(JihSun(日盛)) and excel:    
![Alt text](https://github.com/heyheychen/Taiwan_futures_placing_order_tool_and_retail_analyzer/blob/master/pic/DDE%20and%20DDE%20excel.jpg?raw=true)   

## Placing order API
We can download from JihSun official website and intall API.    
Open HTSAPI3.0_app_VBA_function-N_for_github.xls page.2 and fill out ID and password for logining into API.   
![Alt text](https://github.com/heyheychen/Taiwan_futures_placing_order_tool_and_retail_analyzer/blob/master/pic/API%20login.jpg?raw=true)   

After login API, we can use python to fill excel data to place an order.    
![Alt text](https://github.com/heyheychen/Taiwan_futures_placing_order_tool_and_retail_analyzer/blob/master/pic/API%20order.jpg?raw=true)   

## Result
I define MTX(小台) less than 5 lot an order as retial trader, TX(大台) bigger than 5 lot an order as institution trader.    
Raw data is one minute a count, the trend is the accumulated data of raw data.    
The trend of retail and institution trader is as followed:    
![Alt text](https://github.com/heyheychen/Taiwan_futures_placing_order_tool_and_retail_analyzer/blob/master/pic/result.png?raw=true)    
According to the data, we can define strategy to place a order, we don't need to reading the tape in front of a pc or cell phone.    
Line bot will also send a reminder if API placing a order:
![Alt text](https://github.com/heyheychen/Taiwan_futures_placing_order_tool_and_retail_analyzer/blob/master/pic/line%20message.jpg?raw=true)  

