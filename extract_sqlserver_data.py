#! python3
# extract_sqlserver_data.py - 

import pyodbc
import pandas as pd
import datetime, os.path

save_path = 'I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\DataPipeline\\'


# Convert data to excel format
def convert_to_excel(data, filename):
    data.to_excel(save_path + filename + '.xlsx', sheet_name='Sheet1', index=False)


# extract booking info
def BK_sql_data(BK_ID_no, user, conn):
    # Booking SQL
    BK_tmp = pd.read_sql("SELECT BK.Id, BK.OwnerId, BK.Name, FORMAT(BK.nihrm__ArrivalDate__c, 'yyyy-MM-dd') AS ArrivalDate, FORMAT(BK.nihrm__DepartureDate__c, 'yyyy-MM-dd') AS DepartureDate, BK.nihrm__CommissionPercentage__c, BK.Percentage_of_Attrition__c, BK.nihrm__Property__c, BK.nihrm__FoodBeverageMinimum__c, ac.Name AS ACName, ag.Name AS AGName, BK.End_User_Region__c, BK.End_User_SIC__c, BK.nihrm__BookingTypeName__c \
                          FROM dbo.nihrm__Booking__c AS BK \
                              LEFT JOIN dbo.Account AS ac \
                                  ON BK.nihrm__Account__c = ac.Id \
                              LEFT JOIN dbo.Account AS ag \
                                  ON BK.nihrm__Agency__c = ag.Id \
                          WHERE BK.Booking_ID_Number__c = " + BK_ID_no, conn)
    BK_tmp['OwnerId'].replace(user, inplace=True)
    # pull the actual booking ID in FDC
    BK_ID = BK_tmp.iloc[0]['Id']
    
    return BK_tmp, BK_ID


# extract roomnight_by_day info function
def RoomN_sql_data(BK_ID, conn):
    # Room SQL
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
    
    return RoomN_tmp


# extract event info function
def Event_sql_data(BK_ID, conn):
    # Event SQL
    Event_tmp = pd.read_sql("SELECT ET.Name, FR.Name, ET.nihrm__EventClassificationName__c, FORMAT(ET.nihrm__StartDate__c, 'yyyy/MM/dd') AS Start, ET.nihrm__AgreedEventAttendance__c, ET.nihrm__ForecastAverageCheck1__c, ET.nihrm__ForecastAverageCheck1__c, ET.nihrm__ForecastRevenue1__c, ET.nihrm__ForecastAverageCheck9__c, ET.nihrm__ForecastAverageCheckFactor9__c, ET.nihrm__ForecastRevenue9__c, ET.nihrm__ForecastAverageCheck2__c, ET.nihrm__ForecastAverageCheckFactor2__c, ET.nihrm__ForecastRevenue2__c, ET.nihrm__FunctionRoomRental__c, ET.nihrm__CurrentBlendedRevenue4__c \
                             FROM dbo.nihrm__BookingEvent__c AS ET \
                             INNER JOIN dbo.nihrm__FunctionRoom__c AS FR \
                                 ON ET.nihrm__FunctionRoom__c = FR.Id \
                             WHERE ET.nihrm__Booking__c = '" + BK_ID + "'", conn)
    Event_tmp.columns = ['Event name', 'Function Space', 'Event Classification', 'Start', 'Agreed', 'Food Check', 'Food Factor', 'Food Revenue', 'Outlet Check', 'Outlet Factor', 'Outlet Revenue', 'Beverage Check', 'Beverage Factor', 'Beverage Revenue', 'Rental Revenue', 'AV Revenue']
    Event_tmp['Start'] = pd.to_datetime(Event_tmp['Start']).dt.date
    
    return Event_tmp


# main function for extract_sqlserver_data
def extract_sqlserver_data(table):
    
    # for testing purpose
#    table = pd.read_csv(os.path.abspath(os.getcwd()) + '\\tmp.csv')
    
    col = table['Booking ID']
    BK_ID_no = str(int(col)).zfill(6)
    

    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=VOPPSCLDBN01\VOPPSCLDBI01;'
                          'Database=SalesForce;'
                          'Trusted_Connection=yes;')
    
    
    # FDC User ID and Name list
    user = pd.read_sql('SELECT DISTINCT(Id), Name \
                        FROM dbo.[User]', conn)
    user = user.set_index('Id')['Name'].to_dict()
    
    
    # load Booking info function
    BK_tmp, BK_ID = BK_sql_data(BK_ID_no, user, conn)
    # load Room info function
    RoomN_tmp = RoomN_sql_data(BK_ID, conn)
    # load Event info function
    Event_tmp = Event_sql_data(BK_ID, conn)

    return BK_tmp, RoomN_tmp, Event_tmp



# Data Checking
#BK_tmp, RoomN_tmp, Event_tmp = extract_sqlserver_data()
#
#BK = 'BKinfo'
#convert_to_excel(BK_tmp, BK)
#
#Room = 'RoomN'
#convert_to_excel(RoomN_tmp, Room)
#
#Event = 'EventT'
#convert_to_excel(Event_tmp, Event)
