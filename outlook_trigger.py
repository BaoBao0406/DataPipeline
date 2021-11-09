#! python3
# outlook_trigger.py - search for keyword in outlook email Subject, and extract table in email Body if criteria is met. 
# and export to csv for later process  


from win32com.client import Dispatch
import datetime, os.path, re, io
import pandas as pd


# Keyword to search for email subject
msgKeyWord = re.compile(r'booking dataflow')


# Loop for all email within the Date Range with the keyword
MsgToMove = []
def search_all_mail(filename_list, msgs):
    for msg in msgs:
        # Search for keywords in email subject
        msgSearch = msgKeyWord.search((msg.Subject).lower())
        if (msgSearch == None) is False:
            msg_name = re.sub('[^a-zA-Z0-9 \n\.]', '', msg.Subject) + '.msg'
            if msg_name not in filename_list:
                MsgToMove.append(msg)
                msgSearch = 'None'
                
# Function to get email body information in table
def extract_email_table(msg):
        # Using read_html to get email table, then use table[0] to convert to DataFrame format
        table_tmp = pd.read_html(msg.HTMLBody, header=0, index_col=0)[0].T
        # Convert to csv to remove the index column then convert back to DatdFrame
        table_tmp = pd.read_csv(io.StringIO(table_tmp.to_csv(index=False)), sep=",")
        # TODO: Add function to debug
        
        return table_tmp
    
# Function to save the unprocess booking email to folder
def save_email(msg):
    # replace all the special character in email Subject
    msg_filename = re.sub('[^a-zA-Z0-9 \n\.]', '', msg.Subject) + '.msg'
    # Save the email to folder
    msg.SaveAs(os.path.abspath(os.getcwd()) + '\\Email\\' + msg_filename)


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


# Main function in outlook_trigger
def outlook_trigger():
    
    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder("6")
    msgs = inbox.Items
    
    # Date Range from last three days
    d = (datetime.date.today() - datetime.timedelta (days=5)).strftime("%d-%m-%y")
    
    # Search in inbox for last three days
    msgs = msgs.Restrict("[ReceivedTime] >= '" + d +"'")
    if msgs:
        # Get all the previous filename save (already process emails)
        filename_list = os.listdir(os.path.abspath(os.getcwd()) + '\\Email\\')
        search_all_mail(filename_list, msgs)
        
        if MsgToMove:
            process_save_email_2_csv()


# Testing purpose
outlook_trigger()
