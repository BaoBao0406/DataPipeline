#! python3
# extract_sqlserver_data.py - 

import pyodbc
import pandas as pd
import datetime, os.path


table = pd.read_csv(os.path.abspath(os.getcwd()) + '\\tmp.csv')
number_of_booking = table.shape[0]


#conn = pyodbc.connect('Driver={SQL Server};'
#                      'Server=VOPPSCLDBN01\VOPPSCLDBI01;'
#                      'Database=SalesForce;'
#                      'Trusted_Connection=yes;')
#
#user = pd.read_sql('SELECT DISTINCT(Id), Name \
#                    FROM dbo.[User]', conn)
#user = user.set_index('Id')['Name'].to_dict()


# TODO: extract booking info

# TODO: extract event info

# TODO: extract roomnight_by_day info


# TODO: main function for extract_sqlserver_data
# def extract_sqlserver_data(data):

