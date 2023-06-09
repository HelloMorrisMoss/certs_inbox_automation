"""This is the main entry point for the program. It will connect to an Outlook application on the local system, find
the relevant inboxes/folders, and make the desired modifications to the items therein.

modifications:
* put a follow-up flag on items from prioritized customer shipments
* move duplicate foam certs out of the main inbox
"""
import datetime
import traceback
from typing import Any, Dict, List, Tuple

import pandas as pd

from helpers.json_help import df_json_handler
from helpers.outlook_helpers import find_folders_in_outlook, valid_colors
from log_setup import lg
from outlook_interface import OutlookSingleton, wc_outlook
from tasks.clean_foam_inbox import get_process_folders_dfs, process_foam_groups
from tasks.mark_priority_emails import set_priority_customer_category
from untracked_config.accounts_and_folder_paths import acct_path_dct
from untracked_config.auto_dedupe_cust_ids import dedupe_cnums
from untracked_config.development_node import ON_DEV_NODE, UNIT_TESTING
from untracked_config.priority_shipment_customers import priority_flag_dict

if ON_DEV_NODE:
    # pandas display settings for development
    pd.set_option('display.max_rows', 100)
    pd.set_option('display.max_columns', 100)
    pd.set_option('display.width', 1000)


def main_process_function(found_folders_dict: Dict[str, Any], production_inbox_folders: List[str]) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """Perform the main processing of mail items.

    This function performs the main processing of mail items based on the provided `found_folders_dict` and
    `production_inbox_folders`. It processes mail items in each folder, sets follow-up flags on priority customer items,
    and performs additional processing tasks. It also generates a summary dictionary containing debug information.

    Args:
        found_folders_dict (Dict[str, Any]): A dictionary containing the found folders.
        production_inbox_folders (List[str]): A list of production inbox folders.

    Returns:
        Tuple[Dict[str, Any], Dict[str, Any]]: A tuple containing the updated `found_folders_dict` and the summary dictionary.
    """
    if ON_DEV_NODE:
        # a summary debug info dictionary
        smry = dict(checked_folders={}, skipped_folders=[], all_subj_lines=[], matched=[], missing_a_match=[],
                    non_regex_matching_emails=[], testing_colors_move=['grey'], valid_colors=valid_colors)
    else:
        smry = dict()
        lg.info('Running on a PRODUCTION system.')

    # config data
    pfdfs: List[Tuple[Any, str]] = get_process_folders_dfs(production_inbox_folders, found_folders_dict)
    found_folders_keys = found_folders_dict.keys()
    move_folder_com = found_folders_dict[acct_path_dct['target_folder_path']]

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


def get_process_ol_folders(wc_outlook: OutlookSingleton) -> Tuple[Dict[str, Any], List[str]]:
    """Retrieve Outlook folders for processing.

    Retrieves the Outlook folders for processing based on the provided `wc_outlook` instance. It gets the current folder
    data, including the target folder path and other relevant information. It then searches for the required folders in
    the Outlook folders using the `find_folders_in_outlook` function.

    Args:
        wc_outlook (OutlookSingleton): An instance of the `OutlookSingleton` class representing the Outlook application.

    Returns:
        Tuple[Dict[str, Any], List[str]]: A tuple containing a dictionary of found folders and a list of production inbox folders.
    """
    account_name = acct_path_dct['account_name']
    inbox_folders = acct_path_dct['inbox_folders']
    # get current folder data
    find_folder_keys = ['target_folder_path']
    if UNIT_TESTING:
        find_folder_keys += ['known_good_final_state_inbox_folder', 'known_good_final_state_inbox_folder',
                             'test_file_origin']
    test_keys = [acct_path_dct[k] for k in find_folder_keys]
    must_find_folders = inbox_folders + test_keys
    ol_folders = wc_outlook.get_outlook_folders()
    found_folders: Dict[str, Any] = find_folders_in_outlook(ol_folders, account_name, must_find_folders)
    return found_folders, inbox_folders


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
        if not ON_DEV_NODE and not UNIT_TESTING:
            try:
                from development_files.email_alert import send_alert
                send_alert(subject='Certs_inbox_automation has encountered an unhandled error!', body=stack_trace_str)
            except Exception as em_exc:
                lg.error(traceback.format_exc())
    finally:
        lg.debug('Deleting Outlook com instance.')
        del (wc_outlook)

pass  # for breakpoint
