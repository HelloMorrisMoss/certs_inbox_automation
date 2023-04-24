from helpers.paths import double_slash_paths
from untracked_config.development_node import ON_DEV_NODE


acct_path_dct = {}

if ON_DEV_NODE:
    development_inbox_folders = [r'\\account_name\Specific Inbox\Automation Testing\active_files\Inbox']
    development_inbox_folders = double_slash_paths(development_inbox_folders)

    acct_path_dct = {
        "account_name": "account_name",
        "sent_items_folder": r'\\account_name\Sent Items',
        "inbox_folders": development_inbox_folders,
        "target_folder_path": r'\\account_name\Inbox\Foam Duplicate Lots',
        }
else:
    production_inbox_folders = [r'\\account_name\1-Specific Inbox', r'\\account_name\Inbox']  # live folders
    production_inbox_folders = double_slash_paths(production_inbox_folders)

    acct_path_dct = {
        "account_name": "account_name",
        "sent_items_folder": r'\\account_name\Sent Items',
        "inbox_folders": production_inbox_folders,
        "target_folder_path": r'\\account_name\Inbox\Foam Duplicate Lots',
        }
