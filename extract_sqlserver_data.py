#! python3
# extract_sqlserver_data.py - 

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
def convert_to_excel(data, filename):
    data.to_excel(save_path + filename + '.xlsx', sheet_name='Sheet1', index=False)


# FDC User ID and Name list
user = pd.read_sql('SELECT DISTINCT(Id), Name \
                    FROM dbo.[User]', conn)
user = user.set_index('Id')['Name'].to_dict()



# extract booking info
BK_tmp = pd.read_sql("SELECT BK.Id, BK.OwnerId, BK.Name, FORMAT(BK.nihrm__ArrivalDate__c, 'MM/dd/yyyy') AS ArrivalDate, FORMAT(BK.nihrm__DepartureDate__c, 'MM/dd/yyyy') AS DepartureDate \
                      FROM dbo.nihrm__Booking__c AS BK \
                      WHERE BK.Booking_ID_Number__c = " + BK_ID_no, conn)
BK_tmp['OwnerId'].replace(user, inplace=True)
BK_ID = BK_tmp.iloc[0]['Id']


# extract roomnight_by_day info
RoomN_tmp = pd.read_sql("SELECT GS.nihrm__Property__c, GS.Name, FORMAT(RoomN.nihrm__PatternDate__c, 'MM/dd/yyyy') AS PatternDate, RoomN.nihrm__AgreedRoomsTotal__c \
                         FROM dbo.nihrm__BookingRoomNight__c AS RoomN \
                         INNER JOIN dbo.nihrm__GuestroomType__c AS GS \
                             ON RoomN.nihrm__GuestroomType__c = GS.Id \
                         WHERE RoomN.nihrm__Booking__c = '" + BK_ID + "'", conn)
RoomN = 'RoomN'
convert_to_excel(RoomN_tmp, RoomN)


# extract event info
Event_tmp = pd.read_sql("SELECT ET.Name, FR.Name, ET.nihrm__EventClassificationName__c, ET.nihrm__StartDate__c, nihrm__AgreedEventAttendance__c \
                         FROM dbo.nihrm__BookingEvent__c AS ET \
                         INNER JOIN dbo.nihrm__FunctionRoom__c AS FR \
                             ON ET.nihrm__FunctionRoom__c = FR.Id \
                         WHERE ET.nihrm__Booking__c = '" + BK_ID + "'", conn)

EventT = 'EventT'
convert_to_excel(Event_tmp, EventT)


# TODO: main function for extract_sqlserver_data
# def extract_sqlserver_data(data):

