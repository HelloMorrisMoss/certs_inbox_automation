"""This uses Outlook, to extract information about the e-mails contained in certain folders."""

import csv

from log_setup import lg

import win32com.client


# Connect to Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Define account name or shared folder name
account_name = "SB-certs"

# Define start and end date range
start_date = '17-Feb-2023'
end_date = '17-Mar-2023'

# Define output CSV file path
output_file_path = r"C:\Users\lmcglaughlin\OneDrive - NITTO DENKO CORPORATION\Documents\Projects\Cert " \
                   r"Automation\Sample files\from outlook\certs_in_inbox_for_counts\raw_data.csv"

# Define folder path to start search from
start_folder_paths = [r'\\SB-certs\1-CERTS Inbox', r'\\SB-certs\Inbox', r'\\SB-certs\Sent Items']

# Get shared Inbox folder
try:
    recipient = outlook.CreateRecipient(account_name)
    recipient.Resolve()
    shared_inbox = outlook.GetSharedDefaultFolder(recipient, 6)
except Exception as e:
    lg.debug(f"Error: {e}")
    quit()

# Open CSV file for writing
with open(output_file_path, 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)

    # Write header row
    writer.writerow(['Folder', 'Date', 'Subject'])

    # Define function to recursively loop through folders
    def process_folder(olFolder):
        items = olFolder.Items.Restrict("[ReceivedTime] >= '" + start_date + "' AND [ReceivedTime] <= '" + end_date + "'")
        for item in items:
            try:
                date = item.ReceivedTime.strftime('%m/%d/%Y %H:%M:%S')
                subject = item.Subject
                writer.writerow([olFolder.FolderPath, date, subject])
            except Exception as e:
                lg.debug(f"Received date error: {e}")
                try:
                    date = item.SentOn.strftime('%m/%d/%Y %H:%M:%S')
                except Exception as e:
                    lg.debug(f"Sent date error: {e}")
                    # date = "N/A"

        for sub_folder in olFolder.Folders:
            process_folder(sub_folder)

    # Find start folders and process them and their subfolders recursively
    for olStore in outlook.Stores:
        for olFolder in olStore.GetRootFolder().Folders:
            if olFolder.FolderPath in start_folder_paths:
                process_folder(olFolder)
