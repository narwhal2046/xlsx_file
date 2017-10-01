#! /usr/bin/env python
# -*- coding:UTF-8 -*-
import xlrd
import pymysql
from datetime import date, datetime, timedelta

print ("begin connect to databases")
conn = pymysql.connect(host='127.0.0.1',
                        user='root',
                        db='mysite')
cursor = conn.cursor()
sql = 'show databases'
cursor.execute(sql)
print(cursor.fetchall())

print ("open xlsx file");
book = xlrd.open_workbook('./test/2.xlsx')
sheet = book.sheet_by_index(0);
print (sheet.name)

days_list = [1, 2, 3, 4, 7, 14, 28, 91, 182]
print (sheet.row_values(0))
for i in range(sheet.nrows):
    start_date, days, end_date, interest, _, fund, income, _, _, _, _, _, _, _, _ = sheet.row_values(i)
    if days not in days_list:
        continue
    if not int(income) > 0:
        continue
    print (start_date, days, end_date, interest, fund, income)
    end_date = xlrd.xldate_as_datetime(end_date, book.datemode)
    end_date = end_date.date()
    print (end_date)

    start_date = datetime.strptime(start_date, "%Y/%m/%d");
    start_date = start_date.date()
    print (start_date)

    sql = 'select id from reverse_repo_categories where days = %s'
    cursor.execute(sql, days)
    category_id, = cursor.fetchone()
    print(category_id)

    sql = 'insert into reverse_repo_record (start_date, end_date, interest, fund, profit, category_id) values (%s, %s, %s, %s, %s, %s)'
    cursor.execute(sql, (start_date, end_date, interest*100, fund, income, category_id))
    conn.commit()
