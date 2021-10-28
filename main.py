#! python3
# main.py - 

import pandas as pd
import datetime, os.path
#import outlook_trigger, extract_sqlserver_data, proforma_sync
import extract_sqlserver_data


# Convert data to excel format
def convert_to_excel(data, filename):
    data.to_excel(save_path + filename + '.xlsx', sheet_name='Sheet1', index=False)
    
save_path = 'I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\DataPipeline\\'

# Run outlook_trigger function
#outlook_trigger.outlook_trigger()

# Run extract_sqlserver_data function if there is any data save in csv
if os.path.exists(os.getcwd() + '\\tmp.csv'):
# TODO: fix the bugs
#    table = pd.read_csv(os.path.abspath(os.getcwd()) + 'tmp.csv')
    table = pd.read_csv(save_path + 'tmp.csv')
    number_of_booking = table.shape[0]
    
    # Run every booking in each row and pass Booking ID to extract_sqlserver_data Function
    for i in range(number_of_booking):
        col = table.iloc[i, :]
        BK_tmp, RoomN_tmp, Event_tmp = extract_sqlserver_data.sqlserver_data(col)
        
        
        
#        BK = 'BKinfo' + str(i)
#        convert_to_excel(BK_tmp, BK)
#        
#        Room = 'RoomN'+ str(i)
#        convert_to_excel(RoomN_tmp, Room)
#        
#        Event = 'EventT' + str(i)
#        convert_to_excel(Event_tmp, Event)

        # TODO: transform data to Proforma 
        
# TODO: transform data to BR
        
# TODO: Convert text for BookingComment

# TODO: Using model to





