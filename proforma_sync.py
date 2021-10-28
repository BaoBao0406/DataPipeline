#! python3
# proforma_sync.py - 

#################################################
import pyodbc
import pandas as pd
import datetime, os.path
import numpy as np


save_path = 'I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\DataPipeline\\'

table = pd.read_csv(os.path.abspath(os.getcwd()) + '\\tmp.csv')


# Convert data to excel format
def convert_to_excel(data, filename):
    data.to_excel(save_path + filename + '.xlsx', sheet_name='Sheet1')

col = table.iloc[0]['Booking ID']
BK_ID_no = str(int(col)).zfill(6)

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=VOPPSCLDBN01\VOPPSCLDBI01;'
                      'Database=SalesForce;'
                      'Trusted_Connection=yes;')


# FDC User ID and Name list
user = pd.read_sql('SELECT DISTINCT(Id), Name \
                    FROM dbo.[User]', conn)
user = user.set_index('Id')['Name'].to_dict()



# extract booking info
BK_tmp = pd.read_sql("SELECT BK.Id, BK.OwnerId, BK.Name, FORMAT(BK.nihrm__ArrivalDate__c, 'yyyy/MM/dd') AS ArrivalDate, FORMAT(BK.nihrm__DepartureDate__c, 'yyyy/MM/dd') AS DepartureDate, BK.nihrm__CommissionPercentage__c, BK.nihrm__Property__c, BK.nihrm__FoodBeverageMinimum__c \
                      FROM dbo.nihrm__Booking__c AS BK \
                      WHERE BK.Booking_ID_Number__c = " + BK_ID_no, conn)
BK_tmp['OwnerId'].replace(user, inplace=True)
BK_ID = BK_tmp.iloc[0]['Id']

# extract event info
Event_tmp = pd.read_sql("SELECT ET.Name, FR.Name, ET.nihrm__EventClassificationName__c, FORMAT(ET.nihrm__StartDate__c, 'yyyy/MM/dd') AS Start, ET.nihrm__AgreedEventAttendance__c, ET.nihrm__ForecastAverageCheck1__c, ET.nihrm__ForecastAverageCheck1__c, ET.nihrm__ForecastRevenue1__c, ET.nihrm__ForecastAverageCheck9__c, ET.nihrm__ForecastAverageCheckFactor9__c, ET.nihrm__ForecastRevenue9__c, ET.nihrm__ForecastAverageCheck2__c, ET.nihrm__ForecastAverageCheckFactor2__c, ET.nihrm__ForecastRevenue2__c, ET.nihrm__FunctionRoomRental__c, ET.nihrm__CurrentBlendedRevenue4__c \
                         FROM dbo.nihrm__BookingEvent__c AS ET \
                         INNER JOIN dbo.nihrm__FunctionRoom__c AS FR \
                             ON ET.nihrm__FunctionRoom__c = FR.Id \
                         WHERE ET.nihrm__Booking__c = '" + BK_ID + "'", conn)
Event_tmp.columns = ['Event name', 'Function Space', 'Event Classification', 'Start', 'Agreed', 'Food Check', 'Food Factor', 'Food Revenue', 'Outlet Check', 'Outlet Factor', 'Outlet Revenue', 'Beverage Check', 'Beverage Factor', 'Beverage Revenue', 'Rental Revenue', 'AV Revenue']
Event_tmp['Start'] = pd.to_datetime(Event_tmp['Start']).dt.date



RoomN_tmp = pd.read_sql("SELECT GS.nihrm__Property__c, GS.Name, FORMAT(RoomN.nihrm__PatternDate__c, 'yyyy/MM/dd') AS PatternDate, \
                             RoomN.nihrm__BlockedRooms1__c, RoomN.nihrm__BlockedRooms2__c, RoomN.nihrm__BlockedRooms3__c, RoomN.nihrm__BlockedRooms4__c, \
                             RoomN.nihrm__BlockedRate1__c, RoomN.nihrm__BlockedRate2__c, RoomN.nihrm__BlockedRate3__c, RoomN.nihrm__BlockedRate4__c \
                         FROM dbo.nihrm__BookingRoomNight__c AS RoomN \
                         INNER JOIN dbo.nihrm__GuestroomType__c AS GS \
                             ON RoomN.nihrm__GuestroomType__c = GS.Id \
                         WHERE RoomN.nihrm__Booking__c = '" + BK_ID + "'", conn)
RoomN_tmp.columns = ['Property', 'Room Type', 'Pattern Date', 'Room1', 'Room2', 'Room3', 'Room4', 'Rate1', 'Rate2', 'Rate3', 'Rate4']

# TODO: Find a better solution to fix this
# Melt Room Night number (4 Occupancy) to columns
Room_no = RoomN_tmp[['Property', 'Room Type', 'Pattern Date', 'Room1', 'Room2', 'Room3', 'Room4']]
Room_no = pd.melt(Room_no, id_vars=['Property', 'Room Type', 'Pattern Date'], value_name='Room')
Room_no['variable'].replace('Room', '', inplace=True, regex=True)
# Melt Room Rate (4 Occupancy) to columns
Room_rate = RoomN_tmp[['Property', 'Room Type', 'Pattern Date', 'Rate1', 'Rate2', 'Rate3', 'Rate4']]
Room_rate = pd.melt(Room_rate, id_vars=['Property', 'Room Type', 'Pattern Date'], value_name='Rate')
Room_rate['variable'].replace('Rate', '', inplace=True, regex=True)
# Join Room Night and Room Rate
RoomN_tmp = pd.merge(Room_no, Room_rate, on=['Property', 'Room Type', 'Pattern Date', 'variable'])
RoomN_tmp['Pattern Date'] = pd.to_datetime(RoomN_tmp['Pattern Date']).dt.date

#################################################


import win32com.client as win32
from win32com.client import constants


excel = win32.DispatchEx("Excel.Application")
wb = excel.Workbooks.Open(save_path + 'Booking Proforma Template_unprotected.xlsx', None, True)


# Calculate each type of meal (Breakfast, Lunch, Dinner) and groupby to find Revenue per pax and Agreed pax By day
def BQT_meal_table(meal_tmp):
    meal_tmp['Total Revenue'] = meal_tmp['Food Revenue'] + meal_tmp['Outlet Revenue']
    meal_tmp = meal_tmp.groupby('Start')[['Agreed', 'Total Revenue']].sum()
    meal_tmp['Revenue per pax'] = meal_tmp['Total Revenue'] / meal_tmp['Agreed']
    meal_tmp = meal_tmp[['Revenue per pax', 'Agreed']].T
    return meal_tmp

# Calculate beverage and groupby to find Revenue per pax and Agreed pax By day
def BQT_beverage_table(beverage_tmp):
    beverage_tmp = beverage_tmp.groupby('Start')[['Agreed', 'Beverage Revenue']].sum()
    beverage_tmp['Revenue per pax'] = beverage_tmp['Beverage Revenue'] / beverage_tmp['Agreed']
    beverage_tmp = beverage_tmp[['Revenue per pax', 'Agreed']].T
    return beverage_tmp


# Calculate room table to find number of room and Revenue per pax By day
def Room_type_table(RoomN_tb_tmp):
    RoomN_tb_tmp['Type'] = pd.np.where(RoomN_tb_tmp['Room Type'].str.contains("Royale"), "King",
                             pd.np.where(RoomN_tb_tmp['Room Type'].str.contains("Bella"), "Double", "Suite"))
    RoomN_tb_tmp = RoomN_tb_tmp[['Pattern Date', 'Type', 'Room', 'Rate']]
    # Capture all three room type for BP table in room (some bookings may only have one or two type)
    room_type = set(['Double', 'King', 'Suite'])
    room_type_inc = set(pd.unique(RoomN_tb_tmp['Type']))
    
    for rm in list(room_type-room_type_inc):
        add_row = [RoomN_tb_tmp.iloc[0]['Pattern Date'], str(rm), 0, 0]
    
    RoomN_tb_tmp = RoomN_tb_tmp.append(pd.DataFrame([add_row], columns=['Pattern Date', 'Type', 'Room', 'Rate']),ignore_index=True)
    RoomN_tb_tmp['Revenue'] = RoomN_tb_tmp['Room'] * RoomN_tb_tmp['Rate']
    RoomN_tb_tmp = RoomN_tb_tmp.groupby(['Pattern Date', 'Type'])['Room', 'Revenue'].sum().unstack(fill_value=0).stack()
    RoomN_tb_tmp['Daily Rate'] = (RoomN_tb_tmp['Revenue'] / RoomN_tb_tmp['Room']).fillna(0)
    RoomN_tb_tmp = RoomN_tb_tmp[['Room', 'Daily Rate']].T
    return RoomN_tb_tmp


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


# Sync data to Proforma Worksheet
ws_Room = wb.Worksheets('A. Room')

# Venetian
RoomN_venetian = RoomN_tmp[RoomN_tmp['Property'].str.contains('Venetian')]
if RoomN_venetian.empty is False:
    RoomN_venetian = Room_type_table(RoomN_venetian)
    ws_Room.Range(ws_Room.Cells(5,2), ws_Room.Cells(6, 2 + RoomN_venetian.shape[1] - 1)).Value = RoomN_venetian.values

# Parisian
RoomN_parisian = RoomN_tmp[RoomN_tmp['Property'].str.contains('Parisian')]
if RoomN_parisian.empty is False:
    RoomN_parisian = Room_type_table(RoomN_parisian)
    ws_Room.Range(ws_Room.Cells(11,2), ws_Room.Cells(12, 2 + RoomN_parisian.shape[1] - 1)).Value = RoomN_parisian.values


# Sync data to BQT Worksheet
ws_BQT = wb.Worksheets('B. BQT')

# F&B minimum
ws_BQT.Range("B7").Value = BK_tmp.iloc[0]['nihrm__FoodBeverageMinimum__c']
# TODO: Rebate


# Sync data to BQT Meal Worksheet
ws_BQT_meal = wb.Worksheets('B1. BQT Meal')
# exclude all package event
Event_wo_package = Event_tmp[~Event_tmp['Event Classification'].str.contains('Package')]
# Breakfast table
breakfast = Event_wo_package[Event_wo_package['Event Classification'].str.contains('Breakfast')]
if breakfast.empty is False:
    breakfast = BQT_meal_table(breakfast)
    ws_BQT_meal.Range(ws_BQT_meal.Cells(7,2), ws_BQT_meal.Cells(8, 2 + breakfast.shape[1] - 1)).Value = breakfast.values
# Lunch table
lunch = Event_wo_package[Event_wo_package['Event Classification'].str.contains('Lunch')]
if lunch.empty is False:
    lunch = BQT_meal_table(lunch)
    ws_BQT_meal.Range(ws_BQT_meal.Cells(22,2), ws_BQT_meal.Cells(23, 2 + lunch.shape[1] - 1)).Value = lunch.values
# Dinner table
dinner = Event_wo_package[Event_wo_package['Event Classification'].str.contains('Dinner')]
if dinner.empty is False:
    dinner = BQT_meal_table(dinner)
    ws_BQT_meal.Range(ws_BQT_meal.Cells(37,2), ws_BQT_meal.Cells(38, 2 + dinner.shape[1] - 1)).Value = dinner.values
# Beverage table
beverage = Event_wo_package
if beverage.empty is False:
    beverage = BQT_beverage_table(beverage)
    ws_BQT_meal.Range(ws_BQT_meal.Cells(51,2), ws_BQT_meal.Cells(52, 2 + beverage.shape[1] - 1)).Value = beverage.values


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






wb.SaveAs(save_path + 'Testing1.xlsx')

wb.Close(True)
