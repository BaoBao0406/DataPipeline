#! python3
# main.py - 

import pandas as pd
import datetime, os.path, os, logging
import outlook_trigger, extract_sqlserver_data, proforma_sync, business_review_sync

# logging error into file
log_file = os.path.abspath(os.getcwd()) + '\\error_log.log'
FORMAT = '%(asctime)s %(levelname)s: %(message)s'
logging.basicConfig(level=logging.ERROR, filename=log_file, filemode='w', format=FORMAT)    


# Run outlook_trigger function
outlook_trigger.outlook_trigger()

# Run extract_sqlserver_data function if there is any data save in csv
if os.path.exists(os.getcwd() + '\\tmp.csv'):

    table = pd.read_csv(os.path.abspath(os.getcwd()) + '\\tmp.csv')
    table.reindex(columns=[*table.columns.tolist(), 'BP_file_path', 'BR_file_path'])
    number_of_booking = table.shape[0]
    
    # Run every booking in each row and pass Booking ID to extract_sqlserver_data Function
    for i in range(number_of_booking):
        # get row of Booking info and pass to sql to run data
        bk_row = table.iloc[i, :]
        
        try:
            # extract sql from sqlserver main function
            BK_tmp, RoomN_tmp, Event_tmp = extract_sqlserver_data.extract_sqlserver_data(bk_row)
        
            # proforma main function
            run_BP = str(bk_row['Proforma']).lower()
            # Set default value for BP_file_path
            bk_row['BP_file_path'] = 'Not Available'
            if run_BP == 'yes':
                # Run proforma main function
                BP_file_path = proforma_sync.proforma_sync(BK_tmp, RoomN_tmp, Event_tmp)
                # Save path to table
                bk_row['BP_file_path'] = BP_file_path
            
            # BR main function
            run_BR = str(bk_row['Business Review']).lower()
            # Set default value for BR_file_path
            bk_row['BR_file_path'] = 'Not Available'
            # Boolean for event size if above 260 rows
            oversize_event_table = False
            if run_BR == 'yes':
                # Check if this is bbf inclusive
                bbf_inc = str(bk_row['Breakfast inclusive']).lower()
                # Run BR main function
                BR_file_path, oversize_event_table = business_review_sync.business_review_sync(BK_tmp, RoomN_tmp, Event_tmp, bbf_inc, oversize_event_table)
                # Save path to table
                bk_row['BR_file_path'] = BR_file_path
        
            # send email to reply with BR and BP path link
            outlook_trigger.reply_notification(bk_row, oversize_event_table)
        
        except Exception as e:
            # log the error to log file
            logging.exception('Booking ID: ' + str(bk_row['Booking ID']))
            # Remove msg file if error occurs
            os.remove(str(bk_row['msg_path']))
            # function to send notification for failing to run
            outlook_trigger.error_notification(bk_row, log_file)
            pass
            
    # remove tmp file
    os.remove(os.getcwd() + '\\tmp.csv')

      
# TODO: Convert text for BookingComment

# TODO: Using model to predict
