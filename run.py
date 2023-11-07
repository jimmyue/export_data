#!/usr/bin/python3
# -*- coding:utf-8 -*-
'''
Created on 2022年9月20日
@author: yuejing
'''
import pymysql
import cx_Oracle
from sqlalchemy import create_engine
import math
import pandas as pd

def ExportData():
	sql='''
	select 1,2 from dual
	'''
	#oracle,注意64位cx_Oracle安装使用问题
	con=create_engine('oracle+cx_oracle://username:password@ip:1521/database')
	#mysql
	#con=create_engine('mysql+pymysql://username:password@ip:3306/database')
	print('查询数据库中...')
	result=pd.read_sql(sql,con)
	rownum=len(result)
	n=math.ceil(rownum/1000000)
	print('需导出记录数：{}'.format(rownum))
	#导出EXCEL
	filename='数据导出.xlsx'
	if n==1:  #记录数小于100w则一个sheet导出
		with pd.ExcelWriter(filename) as writer:
			result.to_excel(writer, sheet_name='Sheet1',index=False)
			print('导出Excel成功！')
	else:     #记录数大于100w则多个sheet导出
		with pd.ExcelWriter(filename) as writer:
			for i in range(1,n+1):
				result.iloc[(i-1)*1000000:i*1000000].to_excel(writer,sheet_name='Sheet{}'.format(i),index=False)
				if i!=n:
					print('Sheet{}导出成功！ 行数：{}-{}'.format(i,(i-1)*1000000+1,i*1000000))
				else:
					print('Sheet{}导出成功！ 行数：{}-{}'.format(i,(i-1)*1000000+1,rownum))

if __name__ == "__main__":
	result=ExportData()
