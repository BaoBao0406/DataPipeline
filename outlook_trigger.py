#! python3
# outlook_trigger.py - search for keyword in outlook email Subject, and extract table in email Body if criteria is met. 
# and export to csv for later process  

import win32com.client as win32
from win32com.client import Dispatch
import datetime, os.path, re, io, os
import pandas as pd


# Keyword to search for email subject
msgKeyWord = re.compile(r'^booking dataflow')


# Function to save the unprocess booking email to folder
def save_email(msg):
    # replace all the special character in email Subject
    msg_filename = re.sub('[^a-zA-Z0-9 \n\.]', '', msg.Subject) + '.msg'
    # Save the email to folder
    msg_path = os.path.abspath(os.getcwd()) + '\\Email\\' + msg_filename
    msg.SaveAs(msg_path)

    return msg_path

# Function to get email body information in table
def extract_email_table(msg, msg_path):
    # Using read_html to get email table, then use table[0] to convert to DataFrame format
    table_tmp = pd.read_html(msg.HTMLBody, header=0, index_col=0)[0].T
    # Convert to csv to remove the index column then convert back to DatdFrame
    table_tmp = pd.read_csv(io.StringIO(table_tmp.to_csv(index=False)), sep=",")
    # Add sender email to table for reply email send
    if msg.SenderEmailType=='EX':
        table_tmp['sender'] = msg.Sender.GetExchangeUser().PrimarySmtpAddress
    else:
        table_tmp['sender'] = msg.SenderEmailAddress
    # Add msg path for process email
    table_tmp['msg_path'] = msg_path
    
    return table_tmp
    
    
# Loop for all email within the Date Range with the keyword
def search_all_mail(filename_list, msgs, MsgToMove):
    for msg in msgs:
        # Search for keywords in email subject
        msgSearch = msgKeyWord.search((msg.Subject).lower())
        if (msgSearch == None) is False:
            # replace all the special character in email Subject
            msg_name = re.sub('[^a-zA-Z0-9 \n\.]', '', msg.Subject) + '.msg'
            # if msg file not exist in filename_list list, append to MsgToMove
            if msg_name not in filename_list:
                MsgToMove.append(msg)
                msgSearch = 'None'
                
    return MsgToMove


# Function to Move email to specific folder and save to csv file
def process_save_email_2_csv(MsgToMove):
    table = pd.DataFrame()
    for msg in MsgToMove:
        # run function "save_email"
        msg_path = save_email(msg)
        # run function "extract_email_table"
        table_tmp = extract_email_table(msg, msg_path)
        
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
    MsgToMove = []
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
        MsgToMove = search_all_mail(filename_list, msgs, MsgToMove)
        
        if MsgToMove:
            process_save_email_2_csv(MsgToMove)


# Function to reply notification email with BR and BP path
def reply_notification(bk_row, oversize_event_table):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = bk_row['sender']
    mail.Subject = 'Notification for successfully create BP/BR'
    mail.GetInspector
    MessageBody = "<p>BP file path :&nbsp;<a href='" + str(bk_row['BP_file_path']) + "'>" + str(bk_row['BP_file_path']) + "</a></p><p>BR file path :&nbsp;<a href='" + str(bk_row['BR_file_path']) + "'>"  + str(bk_row['BR_file_path']) + "</a></p>"
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    # if event row above 260 rows will send event table as attachment in reply email as BR cannot paste more than 260 rows
    if oversize_event_table:
        mail.Attachments.Add(os.path.abspath(os.getcwd()) + '\\Documents\\event_table.xlsx')
        mail.send
        # remove tmp event table
        os.remove(os.path.abspath(os.getcwd()) + '\\Documents\\event_table.xlsx')
    else:
        mail.send
        

# Function to send out error notification email
def error_notification(bk_row, log_file):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = bk_row['sender']
    mail.CC = list(['patrick.leong@sands.com.mo', 'joyce.ieong@sands.com.mo', 'tony.ho.lucas@sands.com.mo'])
    mail.Subject = 'Error Notification for Booking ID :' + str(bk_row['Booking ID']) + ' in Booking Dataflow'
    mail.GetInspector
    MessageBody = "<p>Noted that an error occurs in your Booking, and due to this error Profoma and Business Review cannot be created.</p> <p>Please double check the below information in your Booking.&nbsp;</p> <p>- Do you enter the correct Booking ID in Booking Dataflow email?</p> <p>- Do you input roomnight and room rate in your Room Block?</p> <p>- Do you enter all necessary information in the Booking?</p> <p><br></p> <p>Systems Team,</p> <p>Access the below link for Error log for this Booking:</p> <a href='" + log_file + "'>Error Log</a></p>"
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    mail.send


# Testing purpose
#outlook_trigger()
