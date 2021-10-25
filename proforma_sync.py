#! python3
# proforma_sync.py - 

#################################################
import pyodbc
import pandas as pd
import datetime, os.path


save_path = 'I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\DataPipeline\\'

table = pd.read_csv(os.path.abspath(os.getcwd()) + '\\tmp.csv')


col = table.iloc[0]['Booking ID']
BK_ID_no = str(int(col)).zfill(6)

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=VOPPSCLDBN01\VOPPSCLDBI01;'
                      'Database=SalesForce;'
                      'Trusted_Connection=yes;')


# Convert data to excel format
#def convert_to_excel(data, filename):
#    data.to_excel(save_path + filename + '.xlsx', sheet_name='Sheet1', index=False)


# FDC User ID and Name list
user = pd.read_sql('SELECT DISTINCT(Id), Name \
                    FROM dbo.[User]', conn)
user = user.set_index('Id')['Name'].to_dict()



# extract booking info
BK_tmp = pd.read_sql("SELECT BK.Id, BK.OwnerId, BK.Name, FORMAT(BK.nihrm__ArrivalDate__c, 'yyyy/MM/dd') AS ArrivalDate, FORMAT(BK.nihrm__DepartureDate__c, 'yyyy/MM/dd') AS DepartureDate, BK.nihrm__CommissionPercentage__c, BK.nihrm__Property__c \
                      FROM dbo.nihrm__Booking__c AS BK \
                      WHERE BK.Booking_ID_Number__c = " + BK_ID_no, conn)
BK_tmp['OwnerId'].replace(user, inplace=True)
BK_ID = BK_tmp.iloc[0]['Id']

# extract event info
Event_tmp = pd.read_sql("SELECT ET.Name, FR.Name, ET.nihrm__EventClassificationName__c, FORMAT(ET.nihrm__StartDate__c, 'MM/dd/yyyy') AS Start, ET.nihrm__AgreedEventAttendance__c, ET.nihrm__ForecastAverageCheck1__c, ET.nihrm__ForecastAverageCheck11__c, ET.nihrm__ForecastAverageCheck9__c, ET.nihrm__ForecastAverageCheckFactor9__c, ET.nihrm__ForecastAverageCheck2__c, ET.nihrm__ForecastAverageCheckFactor2__c, ET.nihrm__FunctionRoomRental__c, ET.nihrm__CurrentBlendedRevenue4__c \
                         FROM dbo.nihrm__BookingEvent__c AS ET \
                         INNER JOIN dbo.nihrm__FunctionRoom__c AS FR \
                             ON ET.nihrm__FunctionRoom__c = FR.Id \
                         WHERE ET.nihrm__Booking__c = '" + BK_ID + "'", conn)
Event_tmp.columns = ['Event name', 'Function Space', 'Event Classification', 'Start', 'Agreed', 'Food Check', 'Food Factor', 'Outlet Check', 'Outlet Factor', 'Beverage Check', 'Beverage Factor', 'Rental Revenue', 'AV Revenue']

#################################################


import win32com.client as win32
from win32com.client import constants


excel = win32.DispatchEx("Excel.Application")
wb = excel.Workbooks.Open(save_path + 'Booking Proforma Template_unprotected.xlsx', None, True)


# Sync data to Proforma Worksheet
ws_Proforma = wb.Worksheets('Proforma')

# Post As
ws_Proforma.Range("C3").Value = BK_tmp.iloc[0]['Name']
# Arrival and Departure
ws_Proforma.Range("C4").Value = BK_tmp.iloc[0]['ArrivalDate'] + ' - ' + BK_tmp.iloc[0]['DepartureDate']
# Venue
ws_Proforma.Range("C5").Value = BK_tmp.iloc[0]['nihrm__Property__c']
# Booking Owner
ws_Proforma.Range("C6").Value = BK_tmp.iloc[0]['OwnerId']


# Sync data to Entertainment Worksheet
ws_entertain = wb.Worksheets('D. Entertainment')

# Arena Rental
arena = (Event_tmp[Event_tmp['Function Space'].str.contains('Arena')])['Rental Revenue'].sum()
ws_entertain.Range("B3").Value = arena
# Venetian Theatre Rental
venetian_theatre = (Event_tmp[Event_tmp['Function Space'].str.contains('Venetian Theatre')])['Rental Revenue'].sum()
ws_entertain.Range("B4").Value = venetian_theatre
# Parisian Theatre Rental
parisian_theatre = (Event_tmp[Event_tmp['Function Space'].str.contains('Parisian Theatre')])['Rental Revenue'].sum()
ws_entertain.Range("B5").Value = parisian_theatre


# Sync data to C&E Worksheet
ws_CE = wb.Worksheets('C. C&E')

# Hall Rental
hall = (Event_tmp[Event_tmp['Function Space'].str.contains('Hall')])['Rental Revenue'].sum()
ws_CE.Range("B5").Value = hall
# AV Revenue
ws_CE.Range("B6").Value = Event_tmp['AV Revenue'].sum()
# Room Rental
ws_CE.Range("B4").Value = Event_tmp['Rental Revenue'].sum() - arena - venetian_theatre - parisian_theatre - hall


# Sync data to BQT Meal Worksheet
ws_BQT_meal = wb.Worksheets('B1. BQT Meal')
Event_wo_package = Event_tmp[~Event_tmp['Event Classification'].str.contains('Package')]

breakfast = Event_wo_package[Event_wo_package['Event Classification'].str.contains('Breakfast')]

lunch = Event_wo_package[Event_wo_package['Event Classification'].str.contains('Lunch')]

dinner = Event_wo_package[Event_wo_package['Event Classification'].str.contains('Dinner')]

wb.SaveAs(save_path + 'Testing1.xlsx')

wb.Close(True)
