#!/usr/bin/env python
# coding: utf-8

# In[1]:


print('hello world')


# In[16]:


import win32com.client 
import time
import pymysql
import datetime
import re

instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr") 
mydb = pymysql.connect(
    user='root', 
    passwd='1234', 
    host='127.0.0.1', 
    db='mydb3', 
    charset='utf8'
)
cursor = mydb.cursor(pymysql.cursors.DictCursor)
CPE_MARKET_KIND = {0:'구분없음', 1: '거래소',2 :'코스닥',3: 'K-OTC', 4: 'KRX',5: 'KONEX'}
Data_Type = ['D','W','M']
for Type in Data_Type:
    for key,value in CPE_MARKET_KIND.items():
        codelist = instCpCodeMgr.GetStockListByMarket(key)
        now = datetime.datetime.now()
        for code in codelist:

            instStockChart.SetInputValue(0,code)
            try:
                sql = 'SELECT trade_date from dayprice where stock_code = %s order by trade_date DESC LIMIT 1;'
                code = re.sub(r'[^0-9]','',code)
                cursor.execute(sql,code)
                lastDate = cursor.fetchall()
                lastDate = lastDate[0]['trade_date']
                date_last = datetime.date(int(lastDate[0:4]),int(lastDate[4:6]),int(lastDate[6:8]))
                date_now = datetime.date(int(now.strftime('%Y')),int(now.strftime('%m')),int(now.strftime('%d')))
                delta_date = (date_now-date_last).days
            except:
                print('err')
                
            if Type == 'D':
                instStockChart.SetInputValue(1, ord('1')) 
                instStockChart.SetInputValue(2, now.strftime('%Y%m%d') )
                instStockChart.SetInputValue(3, lastDate)
                instStockChart.SetInputValue(6, ord('D'))
            elif Type == 'W' and now.weekday() == 5:
                instStockChart.SetInputValue(1, ord('2')) 
                instStockChart.SetInputValue(4, str(delta_date))
                instStockChart.SetInputValue(6, ord('W'))
            elif Type == 'M' and int(now.strftime('%d')) == 1:
                instStockChart.SetInputValue(1, ord('2')) 
                instStockChart.SetInputValue(4, str(delta_date))
                instStockChart.SetInputValue(6, ord('M'))
            
            instStockChart.SetInputValue(5, (0, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26)) 
            instStockChart.SetInputValue(9, ord('1'))

            rows = []
            while True:
                instStockChart.BlockRequest()
                time.sleep(0.25)
                numData = instStockChart.GetHeaderValue(3)
                print(numData)
                for i in range(numData):
                    row = []
                    code = re.sub(r'[^0-9]','',code)
                    row.append(instCpCodeMgr.CodeToName(code)) # Name
                    row.append(code) # Code
                    row.append(value) # Market Kind
                    trade_date = instStockChart.GetDataValue(0, i)
                    row.append(trade_date) # 날짜 date
                    row.append(instStockChart.GetDataValue(1, i)) # 시가 open
                    row.append(instStockChart.GetDataValue(2, i)) # 고가 high
                    row.append(instStockChart.GetDataValue(3, i)) # 저가 low
                    row.append(instStockChart.GetDataValue(4, i)) # 종가 close
                    row.append(instStockChart.GetDataValue(5, i)) # 전일대비 turnover
                    row.append(instStockChart.GetDataValue(6, i)) # 거래량 volume
                    row.append(instStockChart.GetDataValue(7, i)) # 거래대금 trading_value
                    row.append(instStockChart.GetDataValue(8, i)) # 누적체결매도수량 acc_executed_sell_shares
                    row.append(instStockChart.GetDataValue(9, i)) # 누적체결매수수량 acc_executed_buy_shares
                    row.append(instStockChart.GetDataValue(10, i)) # 상장주식수 num_listing_shares
                    row.append(instStockChart.GetDataValue(11, i)) # |시가총액 market_cap
                    row.append(instStockChart.GetDataValue(12, i)) # 외국인주문한도수량 foreigner_order_limit
                    row.append(instStockChart.GetDataValue(13, i)) # 외국인주문가능수량 foreigner_orderable_shares
                    row.append(instStockChart.GetDataValue(14, i)) # 외국인현보유수량 foreigner_holding_shares
                    row.append(instStockChart.GetDataValue(15, i)) # 외국인현보유비율 foreigner_holding_ratio
                    row.append(instStockChart.GetDataValue(16, i)) # 수정주가일자 adjusted_price_date
                    row.append(instStockChart.GetDataValue(17, i)) # 수정주가비율 adjusted_price_ratio
                    row.append(instStockChart.GetDataValue(18, i)) # 기관순매수량 institution_net_buy_shares
                    row.append(instStockChart.GetDataValue(19, i)) # 기관누적순매수량 institution_acc_net_buy_shares
                    row.append(instStockChart.GetDataValue(20, i)) # 등락주선 advanced_decline_line
                    row.append(instStockChart.GetDataValue(21, i)) # 등락비율 advanced_decline_ratio
                    row.append(instStockChart.GetDataValue(22, i)) # 예탁금 deposit
                    row.append(instStockChart.GetDataValue(23, i)) # 주식회전율 turnover_ratio
                    row.append(instStockChart.GetDataValue(24, i)) # 거래성립률 
                    if Type == 'D':
                        sql = """INSERT INTO dayprice (stock_name, stock_code,market_kind,trade_date,open_price,high_price,low_price,close_price,turnover,
                            volume,trading_value,acc_executed_sell_shares,acc_executed_buy_shares,num_listing_shares,market_cap,foreigner_order_limit,
                            foreigner_orderable_shares,foreigner_holding_shares,foreigner_holding_ratio,adjusted_price_date,adjusted_price_ratio,
                            institution_net_buy_shares,institution_acc_net_buy_shares,advanced_decline_line,advanced_decline_ratio,deposit,turnover_ratio,
                            transaction_success_ratio) SELECT %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s FROM DUAL
                            WHERE NOT EXISTS (SELECT * FROM dayprice WHERE stock_code = %s AND trade_date = %s);"""
                    if Type == 'W':
                        sql = """INSERT INTO weekprice (stock_name, stock_code,market_kind,trade_date,open_price,high_price,low_price,close_price,turnover,
                            volume,trading_value,acc_executed_sell_shares,acc_executed_buy_shares,num_listing_shares,market_cap,foreigner_order_limit,
                            foreigner_orderable_shares,foreigner_holding_shares,foreigner_holding_ratio,adjusted_price_date,adjusted_price_ratio,
                            institution_net_buy_shares,institution_acc_net_buy_shares,advanced_decline_line,advanced_decline_ratio,deposit,turnover_ratio,
                            transaction_success_ratio) SELECT %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s FROM DUAL
                            WHERE NOT EXISTS (SELECT * FROM weekprice WHERE stock_code = %s AND trade_date = %s);"""
                    if Type == 'M':
                        sql = """INSERT INTO monthprice (stock_name, stock_code,market_kind,trade_date,open_price,high_price,low_price,close_price,turnover,
                            volume,trading_value,acc_executed_sell_shares,acc_executed_buy_shares,num_listing_shares,market_cap,foreigner_order_limit,
                            foreigner_orderable_shares,foreigner_holding_shares,foreigner_holding_ratio,adjusted_price_date,adjusted_price_ratio,
                            institution_net_buy_shares,institution_acc_net_buy_shares,advanced_decline_line,advanced_decline_ratio,deposit,turnover_ratio,
                            transaction_success_ratio) SELECT %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s FROM DUAL
                            WHERE NOT EXISTS (SELECT * FROM monthprice WHERE stock_code = %s AND trade_date = %s);"""
                    row.append(code)
                    row.append(str(trade_date)) # 날짜
                    values = tuple((row))
                    cursor.execute(sql,values)
                    mydb.commit()
                    rows.append(row)

                if instStockChart.Continue == False:
                    print(code +' 완료')
                    break


# %%
