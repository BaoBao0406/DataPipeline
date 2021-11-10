#! python3
# business_review_sync.py - 

import pandas as pd
import os.path, datetime
import numpy as np
from datetime import timedelta

import win32com.client as win32
from win32com.client import constants
import glob


# Add up the restaurant revenue for Rooms Worksheet
def event_rest_revenue(ws_Rooms, Event_tmp_rest, rest_et_list):
    for rest in rest_et_list.keys():
        tmp = Event_tmp_rest[Event_tmp_rest['Function Space'].str.contains(rest)]
        ws_Rooms.Range("B" + str(rest_et_list[rest])).Value = tmp['Total F&B Revenue'].sum()


# Transfer data to excel Rooms Worksheet
def rooms_info(wb, BK_tmp, Event_tmp):
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
    property_id_list = {'Venetian': '2', 'Conrad': '3', 'Londoner': '4', 'Parisian': '5'}
    for d in property_id_list.keys():
        # Loop for all property in property_id_list list
        if d in BK_tmp.iloc[0]['nihrm__Property__c']:
            # Primary property by ID
            ws_Rooms.Range("O" + property_id_list[d]).Value = BK_tmp.iloc[0]['Booking_ID_Number__c']
    
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
    property_FB_list = {'Venetian': '28', 'Conrad': '38', 'Parisian': '46'}
    for d in property_FB_list.keys():
        # Loop for all property in property_FB_list list
        if d in BK_tmp.iloc[0]['nihrm__Property__c']:
            # F&B Minimum by property
            ws_Rooms.Range("B" + property_FB_list[d]).Value = BK_tmp.iloc[0]['nihrm__FoodBeverageMinimum__c']
    
    # Food and Beverage Part
    if Event_tmp.empty is False:
        Event_tmp_rest = Event_tmp[['Function Space', 'Event Classification', 'Food Revenue', 'Outlet Revenue', 'Beverage Revenue']]
        Event_tmp_rest['Total F&B Revenue'] = Event_tmp_rest['Food Revenue'] + Event_tmp_rest['Outlet Revenue'] + Event_tmp_rest['Beverage Revenue']
        # Exclude Package and Breakfast classification type
        Event_tmp_rest = Event_tmp_rest[~Event_tmp_rest['Event Classification'].str.contains('Package|Breakfast')]
        # Venetian restaurant list and excel cell
        venetian_rest = {'Bambu': 29, 'Jiang Nan': 30, 'Imperial House': 31, 'Golden Peacock': 32, 'North': 33, 'Portofino': 34}
        # Run function for Venetian rest revenue
        event_rest_revenue(ws_Rooms, Event_tmp_rest, venetian_rest)
        # Conrad restaurant list and excel cell
        conrad_rest = {'Churchill': 39, 'Southern Kitchen': 42}
        # Run function for Conrad rest revenue
        event_rest_revenue(ws_Rooms, Event_tmp_rest, conrad_rest)
        # Conrad restaurant list and excel cell
        parisian_rest = {'Market Bistro': 47, 'Le Buffet': 48, 'Brasserie': 49, 'Lotus Palace': 50, 'La Chine': 51}
        # Run function for Parisian rest revenue
        event_rest_revenue(ws_Rooms, Event_tmp_rest, parisian_rest)


# Room and Rates part
def rooms_rates_info(wb, RoomN_tmp, start_bk):
    
    ws_Rooms = wb.Worksheets('Rooms')
    ws_Rates = wb.Worksheets('Daily Rates')
    
    # Room number
    property_rm_dict = {'The Venetian Macao': 'Venetian', 'Conrad Macao': 'Conrad', 
                        'The Londoner Macao Hotel': 'Londoner', 'The Parisian Macao': 'Parisian'}
    RoomN_tmp['Room'] = RoomN_tmp['Room'].replace(np.NaN, 0, regex=True)
    RoomN_rm_tmp = RoomN_tmp[RoomN_tmp['Room'] != 0]
    RoomN_rm_tmp['Property'] = RoomN_rm_tmp['Property'].replace(property_rm_dict)
    property_rm_list = property_rm_dict.values()
    RoomN_rm_tmp = RoomN_rm_tmp[RoomN_rm_tmp['Property'].isin(property_rm_list)]
    
    # Find Room Block start day
    start_rm = RoomN_rm_tmp['Pattern Date'].min()
    # Calculate Difference between Room Block day and Booking day
    date_diff = (pd.to_datetime(start_rm) - start_bk).days
    
    # TODO: replace BR Room type
    # df.set_index('id')['value'].to_dict()
    # TODO: if inlude bbf add breakfast to room rate
    
    RoomN_rm_tmp = RoomN_rm_tmp.groupby(['Room Block Name', 'Property', 'Room Type', 'variable', 'Pattern Date'])['Room', 'Rate'].sum()
    RoomN_rm_tmp = RoomN_rm_tmp.unstack()
    
    # Room table
    room_tmp = RoomN_rm_tmp['Room']
    # Rate table
    rate_tmp = RoomN_rm_tmp['Rate']
    # Calculate Ask Rate
    revenue_by_type = (room_tmp * rate_tmp).sum(axis=1)
    room_by_type = room_tmp.sum(axis=1)
    ask_rate = revenue_by_type / room_by_type

    # Paste Room Block name and Properties
    block_property = RoomN_rm_tmp.reset_index()[['Room Block Name', 'Property']]
    ws_Rooms.Range(ws_Rooms.Cells(70, 1), ws_Rooms.Cells(70 + block_property.shape[0] - 1, 2)).Value = block_property.values
    # Paste Room Type
    #room_type = RoomN_rm_tmp.reset_index()['Room Type']
    #ws_Rooms.Range(ws_Rooms.Cells(70, 3), ws_Rooms.Cells(70 + room_type.shape[0] - 1, 3)).Value = room_type.values
    
    # Paste Room Type and Ask rate
    ask_rate = ask_rate.reset_index()[['Room Type', 0]]
    convert_to_excel(ask_rate, 'rate')
    ws_Rooms.Range(ws_Rooms.Cells(70, 5), ws_Rooms.Cells(70 + ask_rate.shape[0] - 1, 5)).Value = ask_rate.values
    # Paste Room table
    ws_Rooms.Range(ws_Rooms.Cells(70, 7 + date_diff), ws_Rooms.Cells(70 + room_tmp.shape[0] - 1, 7 + date_diff + room_tmp.shape[1] - 1)).Value = room_tmp.values
    # Paste Rate table
    
    
# Transfer data to excel Meeting Space Worksheet
def meeting_space_info(wb, RoomN_tmp, Event_tmp):
    # Meeting Space Worksheet
    ws_Events = wb.Worksheets('Meeting Space')
    
    # Column index for excel input by properties
    property_et_list = {'Venetian': 8, 'Conrad': 9, 'Londoner': 9, 'Parisian': 10}
    
    # Peak Area and Peak day
    if Event_tmp.empty is False:
        Events_tb_tmp = Event_tmp[['Property', 'Start', 'Agreed', 'Area']]
        Events_tb_tmp.sort_values(by='Start', inplace=True)
        # Loop for all property in property_et_list list
        for d in property_et_list.keys():
            # Filter event by property
            Events_loop_tmp = Events_tb_tmp[Events_tb_tmp['Property'].str.contains(d)].reset_index(drop=True)
            if Events_loop_tmp.empty is False:
                # Find the row index number for max Area
                index = Events_loop_tmp['Area'].idxmax()
                # Peak meeting date by property
                ws_Events.Cells(18, property_et_list[d]).Value = str(Events_loop_tmp.iloc[index]['Start'])
                # Peak SQM by property
                ws_Events.Cells(19, property_et_list[d]).Value = Events_loop_tmp.iloc[index]['Area']
    
    # Peak Room day
    if RoomN_tmp.empty is False:
        Room_tb_tmp = RoomN_tmp[['Property', 'Room']]
        # Loop for all property in property_et_list list
        for d in property_et_list.keys():
            Room_loop_tmp = Room_tb_tmp[Room_tb_tmp['Property'].str.contains(d)].reset_index(drop=True)
            if Room_loop_tmp.empty is False:
                # Find the row index number for Room
                index = Room_loop_tmp['Room'].idxmax()
                # Peak room by property
                ws_Events.Cells(20, property_et_list[d]).Value = Room_loop_tmp.iloc[index]['Room']
                
    # Event table
    if Event_tmp.empty is False:
        Events_tb_tmp = Event_tmp[['Start', 'Start Time', 'End Time', 'Event Classification', 'Setup', 'Function Space', 'Rental Revenue', 'Agreed', 'Function Space Option']]
        # Replace NaN value for Agreed and Rental Revenue
        Events_tb_tmp['Agreed'] = Events_tb_tmp['Agreed'].replace(np.NaN, 0, regex=True)
        Events_tb_tmp['Rental Revenue'] = Events_tb_tmp['Rental Revenue'].replace(np.NaN, 0, regex=True)
        Events_tb_tmp.sort_values(by='Start', inplace=True)
        Events_tb_tmp['Start'] = Events_tb_tmp['Start'].astype(str)
        # Transfer event table to BR
        ws_Events.Range(ws_Events.Cells(24,2), ws_Events.Cells(24 + Events_tb_tmp.shape[0] - 1, 10)).Value = Events_tb_tmp.values


# main function for business_review_sync
def business_review_sync(BK_tmp, RoomN_tmp, Event_tmp):
    
    BR_folder = 'X:\\VML\\Sales\\Business_Review\\'
    
    excel = win32.DispatchEx("Excel.Application")
    
    BR_file = glob.glob(BR_folder + 'BR Form_Macao_*.xlsm')
    wb = excel.Workbooks.Open(BR_file[0], None, True)
    

    
    # Run rooms info function
    rooms_info(wb, BK_tmp, Event_tmp)
    
    
    # Find Booking start day
    start_bk = pd.to_datetime(BK_tmp.iloc[0]['ArrivalDate'])
    # Run rooms rates info function
    rooms_rates_info(wb, RoomN_tmp, start_bk)
    
    
    # Run meeting space info function
    meeting_space_info(wb, RoomN_tmp, Event_tmp)
    
    # excel filename format
    excelfile_name = 'BR_' + BK_tmp.iloc[0]['ArrivalDate'] + '_' + BK_tmp.iloc[0]['Name'] + '.xlsm'
    
#    # BR saving path
#    bk_year = pd.to_datetime(BK_tmp.iloc[0]['ArrivalDate']).year
#    bk_month_number = str(pd.to_datetime(BK_tmp.iloc[0]['ArrivalDate']).month)
#    bk_month_name = datetime.datetime.strptime(bk_month_number, "%m")
#    bk_month = bk_month_number + '-' + bk_month_name.strftime("%b")
#    BR_save_file = BR_folder + str(bk_year) + '\\' + bk_month
#    # if folder not exists create folder
#    if not os.path.exists(BR_save_file):
#        os.makedirs(BR_save_file)
#    # Save as excel in BR saving path
#    wb.SaveAs(BR_save_file + '\\' + excelfile_name)
#    wb.Close(True)


    save_path = 'I:\\10-Sales\\+Dept Admin (3Y, Internal)\\2021\\Personal Folders\\Patrick Leong\\Python Code\\DataPipeline\\Testing files\\'

    wb.SaveAs(save_path + excelfile_name)
    wb.Close(True)


#################################################
import pyodbc

save_path = 'I:\\10-Sales\\+Dept Admin (3Y, Internal)\\2021\\Personal Folders\\Patrick Leong\\Python Code\\DataPipeline\\Testing files\\'

#table = pd.read_csv(os.path.abspath(os.getcwd()) + '\\tmp.csv')


# Convert data to excel format
def convert_to_excel(data, filename):
    data.to_excel(save_path + filename + '.xlsx', sheet_name='Sheet1')

#col = table.iloc[0]['Booking ID']
# Testing booking with BK_ID directly
BK_ID_no = '014760'
#BK_ID_no = str(int(col)).zfill(6)


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
                                ET.nihrm__ForecastAverageCheckFactor2__c, ET.nihrm__ForecastRevenue2__c, ET.nihrm__FunctionRoomRental__c, ET.nihrm__CurrentBlendedRevenue4__c, ET.nihrm__StartTime24Hour__c, ET.nihrm__EndTime24Hour__c, ET.nihrm__FunctionRoomSetupName__c, FRO.Name, FR.nihrm__Area__c \
                         FROM dbo.nihrm__BookingEvent__c AS ET \
                         INNER JOIN dbo.nihrm__FunctionRoom__c AS FR \
                             ON ET.nihrm__FunctionRoom__c = FR.Id \
                         LEFT JOIN  dbo.nihrm__FunctionRoom__c AS FRO \
                             ON ET.nihrm__FunctionRoomOption__c = FRO.Id\
                         WHERE ET.nihrm__Booking__c = '" + BK_ID + "'", conn)
Event_tmp.columns = ['Property', 'Event name', 'Function Space', 'Event Classification', 'Start', 'Agreed', 'Food Check', 'Food Factor', 'Food Revenue', 'Outlet Check', 'Outlet Factor', 'Outlet Revenue', 'Beverage Check', 'Beverage Factor', 'Beverage Revenue', 'Rental Revenue', 'AV Revenue', 'Start Time', 'End Time', 'Setup', 'Function Space Option', 'Area']
Event_tmp['Start'] = pd.to_datetime(Event_tmp['Start']).dt.date



RoomN_tmp = pd.read_sql("SELECT GS.nihrm__Property__c, GS.Name, FORMAT(RoomN.nihrm__PatternDate__c, 'yyyy/MM/dd') AS PatternDate, RoomB.Name, \
                             RoomN.nihrm__BlockedRooms1__c, RoomN.nihrm__BlockedRooms2__c, RoomN.nihrm__BlockedRooms3__c, RoomN.nihrm__BlockedRooms4__c, \
                             RoomN.nihrm__BlockedRate1__c, RoomN.nihrm__BlockedRate2__c, RoomN.nihrm__BlockedRate3__c, RoomN.nihrm__BlockedRate4__c \
                         FROM dbo.nihrm__BookingRoomNight__c AS RoomN \
                         INNER JOIN dbo.nihrm__BookingRoomBlock__c AS RoomB \
                             ON RoomN.nihrm__RoomBlock__c = RoomB.Id \
                         INNER JOIN dbo.nihrm__GuestroomType__c AS GS \
                             ON RoomN.nihrm__GuestroomType__c = GS.Id \
                         WHERE RoomN.nihrm__Booking__c = '" + BK_ID + "'", conn)
RoomN_tmp.columns = ['Property', 'Room Type', 'Pattern Date', 'Room Block Name', 'Room1', 'Room2', 'Room3', 'Room4', 'Rate1', 'Rate2', 'Rate3', 'Rate4']

# TODO: Find a better solution to fix this
# Melt Room Night number (4 Occupancy) to columns
Room_no = RoomN_tmp[['Property', 'Room Type', 'Pattern Date', 'Room Block Name', 'Room1', 'Room2', 'Room3', 'Room4']]
Room_no = pd.melt(Room_no, id_vars=['Property', 'Room Type', 'Pattern Date', 'Room Block Name'], value_name='Room')
Room_no['variable'].replace('Room', '', inplace=True, regex=True)
# Melt Room Rate (4 Occupancy) to columns
Room_rate = RoomN_tmp[['Property', 'Room Type', 'Pattern Date', 'Room Block Name','Rate1', 'Rate2', 'Rate3', 'Rate4']]
Room_rate = pd.melt(Room_rate, id_vars=['Property', 'Room Type', 'Pattern Date', 'Room Block Name'], value_name='Rate')
Room_rate['variable'].replace('Rate', '', inplace=True, regex=True)
# Join Room Night and Room Rate
RoomN_tmp = pd.merge(Room_no, Room_rate, on=['Property', 'Room Type', 'Pattern Date', 'Room Block Name', 'variable'])
RoomN_tmp['Pattern Date'] = pd.to_datetime(RoomN_tmp['Pattern Date']).dt.date


#################################################


business_review_sync(BK_tmp, RoomN_tmp, Event_tmp)
