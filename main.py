#! python3
# main.py - 

import pandas as pd
import datetime, os.path, os
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
        # get row of Booking info and pass to sql to run data
        bk_row = table.iloc[i, :]
        BK_tmp, RoomN_tmp, Event_tmp = extract_sqlserver_data.extract_sqlserver_data(bk_row)
        
        # proforma main function
        run_BP = str(bk_row['Proforma']).lower()
        if run_BP == 'yes':
            # Run proforma main function
            proforma_sync.proforma_sync(BK_tmp, RoomN_tmp, Event_tmp)
            
        # BR main function
        run_BR = str(bk_row['Business Review']).lower()
        if run_BR == 'yes':
            # Check if this is bbf inclusive
            bbf_inc = str(bk_row['Breakfast inclusive']).lower()
            # Run BR main function
            business_review_sync.business_review_sync(BK_tmp, RoomN_tmp, Event_tmp, bbf_inc)
    
    # remove tmp file
    os.remove(os.getcwd() + '\\tmp.csv')
    
        
# TODO: Convert text for BookingComment

# TODO: Using model to
