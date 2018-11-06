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

