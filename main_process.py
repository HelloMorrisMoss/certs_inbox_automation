"""This is the main entry point for the program. It will connect to an Outlook application on the local system, find
the relevant inboxes/folders, and make the desired modifications to the items therein.

modifications:
* put a follow-up flag on items from prioritized customer shipments
* WIP: move duplicate foam certs out of the main inbox
"""
import datetime
import traceback

import pandas as pd

from helpers.json_help import df_json_handler
from helpers.outlook_helpers import find_folders_in_outlook, valid_colors
from log_setup import lg
from outlook_interface import wc_outlook
from tasks.clean_foam_inbox import get_process_folders_dfs, process_foam_groups
from tasks.mark_priority_emails import set_priority_customer_category
from untracked_config.accounts_and_folder_paths import acct_path_dct
from untracked_config.auto_dedupe_cust_ids import dedupe_cnums
from untracked_config.development_node import DEV_TEST_MODE, ON_DEV_NODE
from untracked_config.priority_shipment_customers import priority_flag_dict


def main_process_function(found_folders_dict, production_inbox_folders):
    if ON_DEV_NODE:
        lg.debug('Running on the development system.')
        # pandas display settings for development
        pd.set_option('display.max_rows', 100)
        pd.set_option('display.max_columns', 100)
        pd.set_option('display.width', 1000)

        # a summary debug info dictionary
        smry = dict(checked_folders={}, skipped_folders=[], all_subj_lines=[], matched=[], missing_a_match=[],
                    non_regex_matching_emails=[], testing_colors_move=['grey'], valid_colors=valid_colors)
    else:
        smry = dict()
        lg.info('Running on a PRODUCTION system.')

    # config data
    pfdfs: list = get_process_folders_dfs(production_inbox_folders, found_folders_dict)
    unmatched_foam_rows = []  # for checking for unmatched items
    found_folders_keys = found_folders_dict.keys()
    move_folder_com = found_folders_dict[acct_path_dct['target_folder_path']]

    lg.info('Folders found: %s', found_folders_keys if found_folders_keys else None)
    # process mail items
    for df, this_folder_path in pfdfs:
        lg.info('Processing %s', this_folder_path)
        if this_folder_path in found_folders_keys:
            lg.info('Setting follow up flags on priority customer items.')
            set_priority_customer_category(df, priority_flag_dict, True)
            process_foam_groups(df[df.c_number.isin(dedupe_cnums)], this_folder_path,
                                move_folder_com, smry)
        else:
            lg.warn(f'Missing {this_folder_path} in checked folders!')

    if ON_DEV_NODE:  # write the smry dictionary to a file to make it easier to look at
        import json
        with open('./last_smry.json', 'w') as jf:
            json.dump(smry, jf, indent=4, default=df_json_handler)
    return found_folders_dict, smry


def get_process_ol_folders(wc_outlook):
    account_name = acct_path_dct['account_name']
    production_inbox_folders = acct_path_dct['inbox_folders']
    # get current folder data
    find_folder_keys = ['target_folder_path']
    if DEV_TEST_MODE:
        find_folder_keys += ['known_good_final_state_inbox_folder', 'known_good_final_state_inbox_folder',
                             'test_file_origin']
    test_keys = [acct_path_dct[k] for k in find_folder_keys]
    must_find_folders = production_inbox_folders + test_keys
    ol_folders = wc_outlook.get_outlook_folders()
    found_folders_dict: dict = find_folders_in_outlook(ol_folders,
                                                       account_name, must_find_folders)
    return found_folders_dict, production_inbox_folders


if __name__ == '__main__':
    # ### some items in this section are for development and demonstration only ###
    now = datetime.datetime.now()
    lg.debug(f'Starting at {now}')
    try:
        found_folders_dict, production_inbox_folders = get_process_ol_folders(wc_outlook)
        main_process_function(found_folders_dict, production_inbox_folders)

    # log and alert on unhandled exceptions
    except Exception as err:
        stack_trace_str = traceback.format_exc()
        lg.error(stack_trace_str)
        if not ON_DEV_NODE:
            try:
                from development_files.email_alert import send_alert

                send_alert(subject='Certs_inbox_automation has encountered an unhandled error!', body=stack_trace_str)
            except Exception as em_exc:
                lg.error(traceback.format_exc())
    finally:
        lg.debug('Deleting Outlook com instance.')
        del (wc_outlook)
# TODO: complete unit tests; next: a test confirming that the inbox looks like it does after "# color the groups"

pass  # for breakpoint
