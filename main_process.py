"""This is the main entry point for the program. It will connect to an Outlook application on the local system, find
the relevant inboxes/folders, and make the desired modifications to the items therein.

modifications:
* put a follow-up flag on items from prioritized customer shipments
* WIP: move duplicate foam certs out of the main inbox
"""
import traceback

import pandas as pd

from helpers.json_help import df_json_handler
from helpers.outlook_helpers import find_folders_in_outlook, reset_testing_mods
from log_setup import lg
from tasks.clean_foam_inbox import wc_outlook, get_process_folders_dfs
from tasks.mark_priority_emails import set_priority_customer_category
from untracked_config.accounts_and_folder_paths import acct_path_dct
from untracked_config.development_node import ON_DEV_NODE
from untracked_config.priority_shipment_customers import priority_flag_dict

if __name__ == '__main__':
        # ### some items in this section are for development and demonstration only ###
    try:
        if ON_DEV_NODE:
            lg.debug('Running on the development system.')
            # pandas display settings for development
            pd.set_option('display.max_rows', 100)
            pd.set_option('display.max_columns', 100)
            pd.set_option('display.width', 1000)

            # a summary debug info dictionary
            smry = dict(checked_folders={}, skipped_folders=[], all_subj_lines=[], matched=[], missing_a_match=[],
                        non_regex_matching_emails=[])
            testing_colors_move = ['grey']
        else:
            lg.info('Running on a PRODUCTION system.')

        # config data
        account_name = acct_path_dct['account_name']
        production_inbox_folders = acct_path_dct['inbox_folders']

        # get current folder data
        found_folders_dict: dict = find_folders_in_outlook(wc_outlook, account_name, production_inbox_folders)
        pfdfs: list = get_process_folders_dfs(account_name, production_inbox_folders, found_folders_dict)
        unmatched_foam_rows = []  # for checking for unmatched items
        found_folders_keys = found_folders_dict.keys()

        lg.info('Folders found: %s', found_folders_keys if found_folders_keys else None)
        # process mail items
        for df, folder_path in pfdfs:
            lg.info('Processing %s', folder_path)
            if folder_path in found_folders_keys:
                lg.info('Setting follow up flags on priority customer items.')
                set_priority_customer_category(df, priority_flag_dict, True)
                if ON_DEV_NODE:
                    reset_testing_mods(df['o_item'])
                # todo: tests for priority category customers

                # unmatched_foam_rows = process_foam_groups(df, folder_path, unmatched_foam_rows,
                #                                           testing_colors_move, valid_colors,
                #                                           found_folders_dict[''], smry)

            else:
                lg.warn(f'Missing {folder_path} in checked folders!')

        if unmatched_foam_rows:
            lg.warn('UNMATCHED ROWS FOR FOAM DUPLICATES!!')
            lg.debug(unmatched_foam_rows)
        pass  # for breakpoint

        # write the smry dictionary to a file to make it easier to look at
        if ON_DEV_NODE:
            import json
            with open('./last_smry.json', 'w') as jf:
                json.dump(smry, jf, indent=4, default=df_json_handler)

    # log and alert on unhandled exceptions
    except Exception as err:
        stack_trace_str = traceback.format_exc()
        lg.error(stack_trace_str)
        try:
            from development_files.email_alert import send_alert

            send_alert(subject='Certs_inbox_automation has encountered an unhandled error!', body=stack_trace_str)
        except Exception as em_exc:
            lg.error(traceback.format_exc())

# TODO: complete unit tests; next: a test confirming that the inbox looks like it does after "# color the groups"

pass  # for breakpoint
