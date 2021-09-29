#! python3
# main.py - 

from win32com.client import Dispatch
import datetime, os.path, re, io
import pandas as pd
import numpy as np


# Path for the Attachment to be saved
AttachPath = 'I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\DataPipeline\\Testing\\'

# Keyword to search for email subject
msgKeyWord = re.compile(r'booking check')


# Loop for all email within the Date Range with the keyword
MsgToMove = []
def search_all_mail(filename_list):
    for msg in msgs:
        # Search for keywords in email subject
        msgSearch = msgKeyWord.search((msg.Subject).lower())
        if (msgSearch == None) is False:
            msg_name = msg.Subject + '.msg'
            if msg_name not in filename_list:
                MsgToMove.append(msg)
                msgSearch = 'None'
                
# Function to get email body information in table
def extract_email_table(msg):
        # Using read_html to get email table, then use table[0] to convert to DataFrame format
        table_tmp = pd.read_html(msg.HTMLBody, header=0, index_col=0)[0].T
        # Convert to csv to remove the index column then convert back to DatdFrame
        table_tmp = pd.read_csv(io.StringIO(table_tmp.to_csv(index=False)), sep=",")
        return table_tmp
    
# Function to save the unprocess booking email to folder
def save_email(msg):
    # replace all the special character in email Subject
    msg_filename = re.sub('[^a-zA-Z0-9 \n\.]', '', msg.Subject) + '.msg'
    # Save the email to folder
    msg.SaveAs(AttachPath + msg_filename)


# Function to Move email to specific folder and save to csv file
def process_save_email_2_csv():
    table = pd.DataFrame()
    for msg in MsgToMove:
        # run function "save_email"
        save_email(msg)
        # run function "extract_email_table"
        table_tmp = extract_email_table(msg)
        # Check if table is empty, if not merge old to new table
        if table.empty:
            table = table_tmp
        else:
            table = pd.concat([table, table_tmp])
    table = table.reset_index().drop(['index'], axis=1)
    # Export table to csv file for later process
    table.to_csv(os.path.abspath(os.getcwd()) + '\\tmp.csv')


if __name__ == '__main__':
    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder("6")
    msgs = inbox.Items
    
    # Date Range from last three days
    d = (datetime.date.today() - datetime.timedelta (days=5)).strftime("%d-%m-%y")
    
    # Search in inbox for last three days
    msgs = msgs.Restrict("[ReceivedTime] >= '" + d +"'")
    
    if msgs:
        # Get all the previous filename save (already process emails)
        filename_list = os.listdir(AttachPath)
        search_all_mail(filename_list)
        
        if MsgToMove:
            process_save_email_2_csv()
