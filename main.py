#! python3
# main.py - 

import pandas as pd
import datetime, os.path
import outlook_trigger, extract_sqlserver_data, proforma_sync, business_review_sync
    
#tmp_path = 'I:\\10-Sales\\+Dept Admin (3Y, Internal)\\2021\\Personal Folders\\Patrick Leong\\Python Code\\DataPipeline\\'

# Run outlook_trigger function
outlook_trigger.outlook_trigger()

# Run extract_sqlserver_data function if there is any data save in csv
if os.path.exists(os.getcwd() + '\\tmp.csv'):

    table = pd.read_csv(os.path.abspath(os.getcwd()) + '\\tmp.csv')
    number_of_booking = table.shape[0]
    
    # Run every booking in each row and pass Booking ID to extract_sqlserver_data Function
    for i in range(number_of_booking):
        col = table.iloc[i, :]
        BK_tmp, RoomN_tmp, Event_tmp = extract_sqlserver_data.extract_sqlserver_data(col)
        # Run proforma main function
        proforma_sync.proforma_sync(BK_tmp, RoomN_tmp, Event_tmp)
        # Run BR main function
        business_review_sync.business_review_sync(BK_tmp, RoomN_tmp, Event_tmp)
        
#        BK = 'BKinfo' + str(i)
#        convert_to_excel(BK_tmp, BK)                                                                                                                                                   
#        
#        Room = 'RoomN'+ str(i)
#        convert_to_excel(RoomN_tmp, Room)
#        
#        Event = 'EventT' + str(i)
#        convert_to_excel(Event_tmp, Event)

        
# TODO: Convert text for BookingComment

# TODO: Using model to
