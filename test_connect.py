#coding:utf-8
'''
Created on 2017年5月29日 上午10:54:22

@author: caowei13622
'''

import pymysql
import xlrd
#import time
import datetime

data = xlrd.open_workbook("E:\\stock\\20170607.xlsx")
sheet = data.sheet_by_name("Sheet1")

#建立一个MySQL连接
conn = pymysql.connect(host="127.0.0.1", port=3306, user="root", passwd="nopasswd",db="mysql",charset="utf8")
# 获得游标对象, 用于逐行遍历数据库数据
cursor = conn.cursor()
# 创建插入SQL语句
insert_sql = """INSERT INTO t_GuPiaoData (GuPiaoCode, TodayOpenPrice, YesterdayClosePrice,\
TodayClosePrice, TodayHightPrice, TodayLowPrice, TurnoverQuantity, \
TurnoverDate, CreateDate) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)"""

#rows = sheet.nrows()         # 最大行数  
#columns = sheet.ncols()   # 最大列数  
#values_list = []
# 创建一个for循环迭代读取xls文件每行数据的, 从第二行开始是要跳过标题
for r in range(1, sheet.nrows):
    GuPiaoCode                  = str(sheet.cell(r,1).value)
    TodayOpenPrice              = float(sheet.cell(r,5).value)
    TodayClosePrice             = float(sheet.cell(r,4).value)
    YesterdayClosePrice         = float(0.00)
    TodayHightPrice             = float(sheet.cell(r,6).value)
    TodayLowPrice               = float(sheet.cell(r,7).value)
    TurnoverQuantity            = int(sheet.cell(r,9).value)
    TurnoverDate                = datetime.date.today()
    CreateDate                  = datetime.date.today()  #系统时间

    print(GuPiaoCode)
    print(TodayOpenPrice)
    print(YesterdayClosePrice)
    print(TodayClosePrice)
    print(TodayHightPrice)
    print(TodayLowPrice)
    print(TurnoverQuantity)
    print(TurnoverDate)
    print(CreateDate)
    #temp = GuPiaoCode,TodayOpenPrice,YesterdayClosePrice,TodayClosePrice,TodayHightPrice,TodayLowPrice,TurnoverQuantity,TurnoverDate,CreateDate
    #values_list.append(temp)
    
    
    values = (GuPiaoCode,TodayOpenPrice,YesterdayClosePrice,TodayClosePrice,TodayHightPrice,TodayLowPrice,TurnoverQuantity,TurnoverDate,CreateDate)
    print(values)
    # 执行sql语句
    #time.sleep(1)
    try:
        cursor.execute(insert_sql, values)
        conn.commit()
    except Exception as err:
        print('出错了！')
        cursor.close()
        conn.close()
    #cursor.execute(insert_sql, ('603999', 19.38, 20.0, 20.13, 19.32, 3068, datetime.date.today(), datetime.date.today()))
    #cursor.execute("INSERT INTO t_GuPiaoData (GuPiaoCode, TodayOpenPrice, YesterdayClosePrice, TodayClosePrice, TodayHightPrice, TodayLowPrice, TurnoverQuantity, TurnoverDate, CreateDate) VALUES ('603999', 19.38, 20.0, 0.00, 20.13, 19.32, 3067338, '20170607','20170607')")
    # 提交
    #conn.commit()
    #cursor.executemany(insert_sql, values)
    #if r == 1:
    #    break
#=======================================================================
#try:
#    cursor.executemany("""INSERT INTO t_GuPiaoData (GuPiaoCode, TodayOpenPrice, YesterdayClosePrice,TodayClosePrice, TodayHightPrice, TodayLowPrice, TurnoverQuantity,TurnoverDate, CreateDate) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)""",values_list)
#    conn.commit()
#except Exception as err:
#    print('失败')
    #cursor.close()
    #conn.close()
#=======================================================================
#cursor.executemany(insert_sql, values)    
#data = []  
#for rx in range(2, rows+1):  
#    for cx in range(1, columns+1):  
#        data.append(str(sheet.cell(row=rx, column=cx).value))  
#    cursor.execute(insert_sql, (data[1], data[5], 0.0, data[4], data[6], data[7], data[9], '20170607', '20170607'))  
#    data = []  
    
# 关闭游标
#cursor.close()
#print(conn)
#print(cursor)


print('提交')
cursor.close()
conn.close()
