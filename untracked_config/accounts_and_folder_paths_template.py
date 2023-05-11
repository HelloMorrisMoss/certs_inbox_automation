from typing import List

from helpers.paths import double_slash_paths
from untracked_config.development_node import ON_DEV_NODE, UNIT_TESTING

if ON_DEV_NODE:
    development_inbox_folders: List[str] = [r'\\account\Specific Inbox\Automation Testing\active_files\Inbox']
    development_inbox_folders: List[str] = double_slash_paths(development_inbox_folders)

    acct_path_dct = {
        "account_name": "account",
        "origin_email_address": "source_email@address.com",
        "sent_items_folder": r'\\account\Sent Items',
        "inbox_folders": development_inbox_folders,
        "target_folder_path": r'\\account\Inbox\Foam Duplicate Lots',
        "local_save_folder_path": "dev/local/files",
        }
    if UNIT_TESTING:
        acct_path_dct.update({
            'known_good_final_state_inbox_folder': r'\\account\unit_test_files\known_good_final_state_inbox_folder',
            'known_good_final_state_move_to_folder': r'\\account\unit_test_files\known_good_state Moved Items',
            'inbox_folders': [r'\\account\unit_test_files\Inbox'],
            "target_folder_path": r'\\account\unit_test_files\Moved Items',
            "test_file_origin": r'\\account\unit_test_files\Inbox - test original files',
            })
else:
    production_inbox_folders: List[str] = [r'\\account\1-Specific Inbox', r'\\account\Inbox']  # live folders
    production_inbox_folders: List[str] = double_slash_paths(production_inbox_folders)

    acct_path_dct = {
        "account_name": "account",
        "origin_email_address": "source_email@address.com",
        "sent_items_folder": r'\\account\Sent Items',
        "inbox_folders": production_inbox_folders,
        "target_folder_path": r'\\account\Inbox\Foam Duplicate Lots',
        "local_save_folder_path": "./local_files/",
        }
