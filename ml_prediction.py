import pyodbc
import pandas as pd
import datetime, os.path


# extract booking info
def BK_sql_data(BK_ID_no, user, conn):
    # Booking SQL
    BK_tmp = pd.read_sql("SELECT BK.Id, BK.Booking_ID_Number__c, BK.OwnerId, BK.Name, FORMAT(BK.nihrm__ArrivalDate__c, 'yyyy-MM-dd') AS ArrivalDate, FORMAT(BK.nihrm__DepartureDate__c, 'yyyy-MM-dd') AS DepartureDate, BK.nihrm__CommissionPercentage__c, BK.Percentage_of_Attrition__c, BK.nihrm__Property__c, BK.nihrm__FoodBeverageMinimum__c, ac.Name AS ACName, ag.Name AS AGName, BK.End_User_Region__c, BK.End_User_SIC__c, BK.nihrm__BookingTypeName__c, \
                                 BK.RSO_Manager__c, BK.Non_Compete_Clause__c, ac.nihrm__RegionName__c, ac.Industry, BK.nihrm__CurrentBlendedRoomnightsTotal__c, BK.nihrm__BlendedGuestroomRevenueTotal__c, BK.VCL_Blended_F_B_Revenue__c, BK.nihrm__CurrentBlendedEventRevenue7__c, BK.nihrm__CurrentBlendedEventRevenue4__c, BK.nihrm__BookingMarketSegmentName__c, BK.Promotion__c, BK.nihrm__CurrentBlendedADR__c, BK.nihrm__PeakRoomnightsBlocked__c, \
                                 FORMAT(BK.nihrm__BookedDate__c, 'yyyy-MM-dd') AS BookedDate, FORMAT(BK.nihrm__LastStatusDate__c, 'yyyy-MM-dd') AS LastStatusDate \
                          FROM dbo.nihrm__Booking__c AS BK \
                              LEFT JOIN dbo.Account AS ac \
                                  ON BK.nihrm__Account__c = ac.Id \
                              LEFT JOIN dbo.Account AS ag \
                                  ON BK.nihrm__Agency__c = ag.Id \
                          WHERE BK.Booking_ID_Number__c = " + BK_ID_no, conn)
    BK_tmp['OwnerId'].replace(user, inplace=True)
    BK_tmp['RSO_Manager__c'].replace(user, inplace=True)
    BK_ID = BK_tmp.iloc[0]['Id']
    
    return BK_tmp, BK_ID


def RoomN_sql_data(BK_ID, conn):
    # Room SQL
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
    
    return RoomN_tmp



def Event_sql_data(BK_ID, conn):
    # Event SQL
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
    
    return Event_tmp

# TODO: RoomB sql for rates
def RoomB_sql_data(BK_ID, conn):
    RoomB_tmp = pd.read_sql("SELECT RoomB.nihrm__Location__c, ac.nihrm__RegionName__c, ac.Industry, ag.Name, ag.Industry, BK.End_User_Region__c, BK.End_User_SIC__c, BK.nihrm__BookingTypeName__c, BK.nihrm__CurrentBlendedRoomnightsTotal__c, BK.VCL_Blended_F_B_Revenue__c, BK.nihrm__CurrentBlendedEventRevenue7__c, BK.nihrm__CurrentBlendedEventRevenue4__c, BK.RSO_Manager__c, BK.nihrm__BookingMarketSegmentName__c, BK.Promotion__c, RoomB.nihrm__PeakRoomnightsAgreed__c, \
                                 RoomB.nihrm__CurrentBlendedADR__c,  \
                             FROM dbo.nihrm__Booking__c AS BK \
                                 INNER JOIN dbo.nihrm__RoomBlock__c AS RoomB \
                                     ON BK.nihrm__Booking__c = RoomB.nihrm__Booking__c \
                                 LEFT JOIN dbo.Account AS ac \
                                     ON BK.nihrm__Account__c = ac.Id \
                                 LEFT JOIN dbo.Account AS ag \
                                     ON BK.nihrm__Agency__c = ag.Id \
                             WHERE BK.nihrm__Booking__c = " + BK_ID, conn)
    
    return RoomB_tmp

# TODO: Use Event_tmp for max Attendance for booking and room block BK & RoomBlock
attendance = Event_tmp['Agreed'].max()

# TODO: booking info for materization percentage
BK_mat_percent = BK_tmp[['Property', 'Account: Region', 'Account: Industry', 'Agency', 'End User Region', 'End User SIC', 'Booking Type', 'Blended Roomnights',
                        'Blended Guestroom Revenue Total', 'Blended F&B Revenue', 'Blended Rental Revenue', 'Blended AV Revenue', 'Attendance',
                        'RSO Manager', 'Market Segment', 'Promotion', 'Blended ADR', 'Peak Roomnights Blocked', 'ArrivalDate', 'DepartureDate', 
                        'BookedDate', 'LastStatusDate']]
# Maximum attendance
BK_mat_percent['Attendance'] = int(attendance)
# calculate Inhouse day (Departure - Arrival)    
BK_mat_percent['Inhouse day'] = pd.to_datetime(BK_mat_percent['DepartureDate']).dt.date - pd.to_datetime(BK_mat_percent['ArrivalDate']).dt.date
# calculate Lead day (Arrival - Booked) 
BK_mat_percent['Lead day'] = pd.to_datetime(BK_mat_percent['ArrivalDate']).dt.date - pd.to_datetime(BK_mat_percent['BookedDate']).dt.date
# calculate Decision day (Last Status date - Booked) 
BK_mat_percent['Decision day'] = pd.to_datetime(BK_mat_percent['LastStatusDate']).dt.date - pd.to_datetime(BK_mat_percent['BookedDate']).dt.date


# TODO: room block info for ADR rate
RB_adr_rate = RoomB_tmp[['']]
# Maximum attendance
RB_adr_rate['Attendance'] = int(attendance)
# calculate Inhouse day (Departure - Arrival) - via RoomN_tmp
RB_adr_rate['Inhouse day'] = RoomN_tmp['Pattern Date'].max() - RoomN_tmp['Pattern Date'].min()
# calculate Lead day (Arrival - Booked) - via RoomN_tmp
RB_adr_rate['Lead day'] = RoomN_tmp['Pattern Date'].min() - pd.to_datetime(BK_mat_percent['BookedDate']).dt.date
# TODO: calculate Decision day (Last Status date - Booked) - via RoomN_tmp
RB_adr_rate['Decision day'] = pd.to_datetime(BK_mat_percent['LastStatusDate']).dt.date - pd.to_datetime(BK_mat_percent['BookedDate']).dt.date
# TODO: take out Blended Revenue Total in RoomBlock


['Arrival Month', 'Booked Month', 'Last Status Month']

