from helpers.outlook_helpers import find_folders_in_outlook
from outlook_interface import wc_outlook
from untracked_config.accounts_and_folder_paths import acct_path_dct

ol_folders = wc_outlook.get_outlook_folders()
account_name = acct_path_dct['account_name']
production_inbox_folders = acct_path_dct['inbox_folders']
found_folders_dict: dict = find_folders_in_outlook(ol_folders, account_name, production_inbox_folders)


# class EventHandler:
#     def OnNewMailEx(self, receivedItemsIDs):
#         for ID in receivedItemsIDs.split(","):
#             mail = wc_outlook.get_outlook().Session.GetItemFromID(ID)
#             lg.debug(mail.Subject)
#             mail.Categories = 'Blue Category'
#
# # create the event handler object
# handler = EventHandler()
#
# fp = '\\\\SB-certs\\1-CERTS Inbox\\Automation Testing\\active_files\\Inbox'
# inbox = found_folders_dict[fp]
# # subscribe to the inbox events
# lg.debug('adding handler')
# inboxItems = inbox.Items
# pythoncom.PumpMessages()
# inboxItems.ItemAdd += handler.OnNewMailEx
# lg.debug('event handler added')