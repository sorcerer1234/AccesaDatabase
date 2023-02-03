# -*- coding: utf-8 -*-


"""
@author: matthew pertici(c) 2023

Import data from a Microsoft Access table or query into a pandas dataframe.
Need to install pyodbc and pandas data libraries to work
"""

import pyodbc
import pandas as pd


def get_data(db_path,sql):
      conn = pyodbc.connect(db_path)
      cursor = conn.cursor()
      cursor.execute(sql)
      rows = cursor.fetchall()
      df = pd.DataFrame( [[ij for ij in i] for i in rows] )
      return df
      
#Select path of your access database
db_path=r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\my_db.accdb;'
#Enter your SQL statement as below
sql_text='SELECT field1, field2, field3, field4 From table'

#calling the function example below, storing the data in a pandas dataframe called new_table

new_table = get_data(db_path, sql_text)

#rename your columns as required, because import will not automatically name them for you
new_table.rename(columns={0: 'field1', 1: 'field2', 2: 'field3', 3: 'field4'}, inplace=True);


#YOUR CODE
print (new_table)