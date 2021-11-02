#! python3
# business_review_sync.py - 

#################################################
import pyodbc
import pandas as pd
import os.path
import numpy as np
from datetime import timedelta


save_path = 'I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\DataPipeline\\'

table = pd.read_csv(os.path.abspath(os.getcwd()) + '\\tmp.csv')


# Convert data to excel format
def convert_to_excel(data, filename):
    data.to_excel(save_path + filename + '.xlsx', sheet_name='Sheet1')

col = table.iloc[0]['Booking ID']
# Testing booking with BK_ID directly
# BK_ID_no = ''
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
BK_tmp = pd.read_sql("SELECT BK.Id, BK.Booking_ID_Number__c, BK.OwnerId, BK.Name, FORMAT(BK.nihrm__ArrivalDate__c, 'yyyy-MM-dd') AS ArrivalDate, FORMAT(BK.nihrm__DepartureDate__c, 'yyyy-MM-dd') AS DepartureDate, BK.nihrm__CommissionPercentage__c, BK.Percentage_of_Attrition__c, BK.nihrm__Property__c, BK.nihrm__FoodBeverageMinimum__c, ac.Name AS ACName, ag.Name AS AGName, BK.End_User_Region__c, BK.End_User_SIC__c, BK.nihrm__BookingTypeName__c, \
                             BK.RSO_Manager__c, BK.Non_Compete_Clause__c \
                      FROM dbo.nihrm__Booking__c AS BK \
                          LEFT JOIN dbo.Account AS ac \
                              ON BK.nihrm__Account__c = ac.Id \
                          LEFT JOIN dbo.Account AS ag \
                              ON BK.nihrm__Agency__c = ag.Id \
                      WHERE BK.Booking_ID_Number__c = " + BK_ID_no, conn)
BK_tmp['OwnerId'].replace(user, inplace=True)
BK_tmp['RSO_Manager__c'].replace(user, inplace=True)
BK_ID = BK_tmp.iloc[0]['Id']


# extract event info
Event_tmp = pd.read_sql("SELECT ET.nihrm__Property__c, ET.Name, FR.Name, ET.nihrm__EventClassificationName__c, FORMAT(ET.nihrm__StartDate__c, 'yyyy/MM/dd') AS Start, ET.nihrm__AgreedEventAttendance__c, ET.nihrm__ForecastAverageCheck1__c, ET.nihrm__ForecastAverageCheck1__c, ET.nihrm__ForecastRevenue1__c, ET.nihrm__ForecastAverageCheck9__c, ET.nihrm__ForecastAverageCheckFactor9__c, ET.nihrm__ForecastRevenue9__c, ET.nihrm__ForecastAverageCheck2__c, \
                                ET.nihrm__ForecastAverageCheckFactor2__c, ET.nihrm__ForecastRevenue2__c, ET.nihrm__FunctionRoomRental__c, ET.nihrm__CurrentBlendedRevenue4__c, ET.nihrm__StartTime24Hour__c, ET.nihrm__EndTime24Hour__c, ET.nihrm__FunctionRoomSetupName__c, ET.nihrm__FunctionRoomOption__c, FR.nihrm__Area__c \
                         FROM dbo.nihrm__BookingEvent__c AS ET \
                         INNER JOIN dbo.nihrm__FunctionRoom__c AS FR \
                             ON ET.nihrm__FunctionRoom__c = FR.Id \
                         WHERE ET.nihrm__Booking__c = '" + BK_ID + "'", conn)
Event_tmp.columns = ['Property', 'Event name', 'Function Space', 'Event Classification', 'Start', 'Agreed', 'Food Check', 'Food Factor', 'Food Revenue', 'Outlet Check', 'Outlet Factor', 'Outlet Revenue', 'Beverage Check', 'Beverage Factor', 'Beverage Revenue', 'Rental Revenue', 'AV Revenue', 'Start Time', 'End Time', 'Setup', 'Function Space Option', 'Area']
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
wb = excel.Workbooks.Open(save_path + 'BR Form_Macao_6.0.3.5.xlsm', None, True)

# Primary property of the Booking
primary_property = BK_tmp.iloc[0]['nihrm__Property__c']

# Rooms Worksheet
ws_Rooms = wb.Worksheets('Rooms')

# Rooms Worksheet Part 1
# Post As
ws_Rooms.Range("B2").Value = BK_tmp.iloc[0]['Name']
# Account Name
ws_Rooms.Range("B4").Value = BK_tmp.iloc[0]['ACName']
# Agency Name
ws_Rooms.Range("B5").Value = BK_tmp.iloc[0]['AGName']
# End User Region
ws_Rooms.Range("B6").Value = BK_tmp.iloc[0]['End_User_Region__c']
# Regional Manager
ws_Rooms.Range("J2").Value = BK_tmp.iloc[0]['RSO_Manager__c']
# Booking Owner
ws_Rooms.Range("J3").Value = BK_tmp.iloc[0]['OwnerId']
# Booking Type
ws_Rooms.Range("J4").Value = BK_tmp.iloc[0]['nihrm__BookingTypeName__c']
# End User Industry
ws_Rooms.Range("J6").Value = BK_tmp.iloc[0]['End_User_SIC__c']
# Non-Compete Clause
if BK_tmp.iloc[0]['End_User_SIC__c'] == 1:
    ws_Rooms.Range("J6").Value = 'Yes'
# Commission
if BK_tmp.iloc[0]['nihrm__CommissionPercentage__c'] != None:
    ws_Rooms.Range("O6").Value = BK_tmp.iloc[0]['nihrm__CommissionPercentage__c'] / 100
# Attrition
if BK_tmp.iloc[0]['Percentage_of_Attrition__c'] != None:
    ws_Rooms.Range("O7").Value = BK_tmp.iloc[0]['Percentage_of_Attrition__c'] / 100
# Booking ID
property_list = {'Venetian': '2', 'Conrad': '3', 'Londoner': '4', 'Parisian': '5'}
for d in property_list.keys():
    if d in BK_tmp.iloc[0]['nihrm__Property__c']:
        ws_Rooms.Range("O" + property_list[d]).Value = BK_tmp.iloc[0]['Booking_ID_Number__c']

# Rooms Worksheet Part 2
# Status
ws_Rooms.Range("B14").Value = 'Prospect'
# Arrival Date
ws_Rooms.Range("B15").Value = BK_tmp.iloc[0]['ArrivalDate']
# LOS - length of stay
ws_Rooms.Range("B16").Value = (pd.to_datetime(BK_tmp.iloc[0]['DepartureDate']) - pd.to_datetime(BK_tmp.iloc[0]['ArrivalDate'])).days
# Request Type
ws_Rooms.Range("B17").Value = 'New Group'

# F&B minimum (Londoner does not exist in BR yet)
property_list = {'Venetian': '28', 'Conrad': '38', 'Parisian': '46'}
for d in property_list.keys():
    if d in BK_tmp.iloc[0]['nihrm__Property__c']:
        ws_Rooms.Range("O" + property_list[d]).Value = BK_tmp.iloc[0]['nihrm__FoodBeverageMinimum__c']

# Food and Beverage Part
venetian_rest = ['Bambu', 'Jiang Nan', 'Imperial House', 'Golden Peacock', 'North', 'Portofino']

conrad_rest = ['Churchill', 'Southern Kitchen']

parisian_rest = ['Market Bistro', 'Le Buffet', 'Brasserie', 'Lotus Palace', 'La Chine']


# Meeting Space Worksheet
ws_Events = wb.Worksheets('Meeting Space')

# Event table
if Event_tmp.empty is False:
    Events_tb_tmp = Event_tmp[['Start', 'Start Time', 'End Time', 'Event Classification', 'Setup', 'Function Space', 'Rental Revenue', 'Agreed', 'Function Space Option']]
    Events_tb_tmp['Agreed'] = Events_tb_tmp['Agreed'].replace(np.NaN, 0, regex=True)
    Events_tb_tmp['Rental Revenue'] = Events_tb_tmp['Rental Revenue'].replace(np.NaN, 0, regex=True)
    Events_tb_tmp.sort_values(by='Start', inplace=True)
    Events_tb_tmp['Start'] = Events_tb_tmp['Start'].astype(str)
    # Transfer event table to BR
    ws_Events.Range(ws_Events.Cells(24,2), ws_Events.Cells(24 + Events_tb_tmp.shape[0] - 1, 10)).Value = Events_tb_tmp.values



excelfile_name = 'BR_' + BK_tmp.iloc[0]['ArrivalDate'] + '_' + BK_tmp.iloc[0]['Name'] + '.xlsm'

wb.SaveAs(save_path + excelfile_name)
wb.Close(True)
