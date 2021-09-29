#! python3
# main.py - 

import pandas as pd
import datetime, os.path
import outlook_trigger, extract_sqlserver_data


# Run outlook_trigger function
#outlook_trigger.outlook_trigger()

# Run extract_sqlserver_data function if there is any data save in csv
if os.path.exists(os.getcwd() + '\\tmp.csv'):
    table = pd.read_csv(os.path.abspath(os.getcwd()) + '\\tmp.csv')
    number_of_booking = table.shape[0]

    #for i in range(number_of_booking):
        #extract_sqlserver_data()

# TODO: transform data to Proforma 
        
# TODO: transform data to BR
        
# TODO: Convert text for BookingComment

# TODO: Using model to predict
