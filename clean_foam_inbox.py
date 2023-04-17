"""Check Outlook inboxes for redundant cert e-mails."""

import datetime
from typing import Dict, List, Optional

import pandas as pd
import pythoncom
from win32com import client as wclient

from log_setup import lg
from untracked_config.accounts_and_folder_paths import acct_path_dct
from untracked_config.auto_dedupe_cust_ids import dedupe_cnums
from untracked_config.foam_clean_product_names import product_names
from untracked_config.subject_regex import subject_pattern

now = datetime.datetime.now()
lg.debug(f'Starting at {now}')

# Connect to Outlook application
outlook = wclient.Dispatch("Outlook.Application", pythoncom.CoInitialize()).GetNamespace("MAPI")


# Process inbox folders
def process_folder(olFolder, folder_path):
    items = olFolder.Items
    results = process_mail_items(folder_path, items)

    if results:
        df = pd.DataFrame(results).sort_values('received_time', axis=0, ascending=True).reset_index()
        df['lot8'] = df['lot_number'].str[:8]
        # Calculate time differences between consecutive rows
        time_diffs = df['received_time'].diff().fillna(pd.Timedelta(seconds=0))
        df['_time_diffs'] = time_diffs  # for visibility during development

        # Group rows into clusters based on time differences and time deltas within clusters
        # new_group_bools = ((time_diffs > pd.Timedelta(minutes=15)) | (time_diffs.shift(-1) > pd.Timedelta(
        # minutes=15)))
        new_group_bools = (time_diffs > pd.Timedelta(minutes=15))
        df['_ngbools'] = new_group_bools  # for visibility during development
        cluster_ids = new_group_bools.cumsum()

        # Add cluster IDs as a new column to the DataFrame
        df['cluster'] = cluster_ids
        return df
    else:
        return pd.DataFrame()


def process_mail_items(folder_path, items):
    results = []
    all_subj = []
    matched_sub = []
    for item in items:
        subject = item.Subject
        all_subj.append(subject)
        match = subject_pattern.match(subject)
        if match:
            matched_sub.append(subject)
            subj_dict: dict = match.groupdict()
            if subj_dict['product_number'] in product_names:

                if subj_dict['c_number'] not in dedupe_cnums:  # only certain customers
                    continue
                # pandas needs datetime.datetime not pywintypes.datetime
                received_time = datetime.datetime.fromtimestamp(item.ReceivedTime.timestamp() + 14400)  # it's in UTC
                if received_time is None:
                    lg.debug(f'No rec time on {subject}')
                    continue
                row = {"received_time": received_time, "subject": subject, 'o_item': item} | subj_dict
                results.append(row)

    smry['checked_folders'][folder_path]['all_subj_lines'] += all_subj
    smry['checked_folders'][folder_path]['matched'] += matched_sub
    return results


def find_folders(store_name_filter: str, must_find_list: List[str] = '', map_all=False) -> Dict[str, any]:
    """
    Recursively searches for all folders within Outlook stores whose display names contain the specified
    store_name_filter string, and returns a dictionary where the keys are the folder paths and the values
    are the corresponding olFolder objects. Raises a custom exception if any of the folders in must_find_list
    are not found.
    """
    must_find_list = must_find_list if (must_find_list and not map_all) else ''
    folders_dict = {}
    target_store = get_store_by_name(store_name_filter)
    parent_folder = target_store.GetRootFolder()
    map_folder_structure_to_flat_dict(folders_dict, parent_folder, must_find_list)
    for folder in must_find_list:
        if folder not in folders_dict.keys():
            raise Exception(f"Required folder '{folder}' not found!")
    return folders_dict


def get_store_by_name(store_name_filter: str) -> Optional[object]:
    """Searches for an Outlook store with a display name that contains the given filter string and returns the first
    store that matches. If no matching store is found, returns None.

    :param store_name_filter: A string to search for in the display names of the Outlook stores.
    :return: The first Outlook store that matches the search filter, or None if no match is found.
    """
    for olStore in outlook.Stores:
        if store_name_filter not in olStore.DisplayName:
            continue
        target_store = olStore
        return target_store
    return None


def map_folder_structure_to_flat_dict(folders_dict: Dict[str, any], parent_folder: wclient.CDispatch,
                                      must_find_list: List[str]) -> None:
    """Iteratively searches for all folders within the specified parent_folder object and updates the
    folders_dict dictionary with the folder paths and olFolder objects. Stops searching as soon as all
    folders in must_find_list have been found.

    :param folders_dict: A dictionary to store the folder paths and olFolder objects.
    :type folders_dict: Dict[str, any]
    :param parent_folder: The parent folder object to search within.
    :type parent_folder: any
    :param must_find_list: A list of folder paths that must be found. If not all are found, continue searching.
    :type must_find_list: List[str]
    :return: None
    :rtype: None
    """
    folders_stack = [parent_folder]
    while folders_stack:
        current_folder = folders_stack.pop()
        for olFolder in current_folder.Folders:
            folder_path = olFolder.FolderPath
            folders_dict[folder_path] = olFolder
            if must_find_list:
                if all(mfitem in folders_dict.keys() for mfitem in must_find_list):
                    return
            folders_stack.append(olFolder)

def series_to_df(srs):
    return pd.DataFrame.from_dict({k: [v] for k, v in srs.to_dict().items()})


def main_folders_process():

    # get a dictionary of folders from the account
    found_folders_dict = find_folders(account_name, production_inbox_folders, map_all=True)
    for folder_path in production_inbox_folders:
        olFolder = found_folders_dict.get(folder_path)
        if olFolder is None:
            lg.debug(f'{folder_path} was not found and will not be processed!')
            continue
        lg.debug(f'Processing folder: {folder_path}')
        smry['checked_folders'][folder_path] = {'all_subj_lines': [], 'matched': []}
        ibdf = process_folder(olFolder, folder_path)

        if ibdf.empty:
            lg.debug(f'No results in {folder_path}')
            continue

        dfg = ibdf.groupby(['product_number', 'so_number', 'lot8', 'c_number'])
        keep_item_rows = []
        move_item_rows = []
        for grp in dfg:
            keep_item_rows.append(grp[1].iloc[0])
            move_item_rows.append([row for row in grp[1].iloc[1:].iterrows()])
        smry['checked_folders'][folder_path]['ibdf'] = ibdf
        smry['checked_folders'][folder_path]['dfg'] = ibdf
        smry['checked_folders'][folder_path]['keep_item_rows'] = keep_item_rows
        smry['checked_folders'][folder_path]['move_item_rows'] = move_item_rows


if __name__ == '__main__':
    # display settings for development
    pd.set_option('display.max_rows', 100)
    pd.set_option('display.max_columns', 100)
    pd.set_option('display.width', 1000)

    # write the smry dictionary to a file to make it easier to look at
    import json


    def df_json_handler(df):
        try:
            return df.to_dict()
        except AttributeError as aer:
            if isinstance(df, pd.Timestamp) or isinstance(df, pd.Timedelta):
                return df.isoformat()
            else:
                return repr(df)


    # global production_inbox_folders, smry
    account_name = acct_path_dct['account_name']
    production_inbox_folders = acct_path_dct['inbox_folders']
    # a summary debug info dictionary
    smry = dict(checked_folders={}, skipped_folders=[], all_subj_lines=[])

    main_folders_process()

    with open('./last_smry.json', 'w') as jf:
        json.dump(smry, jf, indent=4, default=df_json_handler)

    # check that the move rows all have keep rows to match
    unmatched = []
    for fp in production_inbox_folders:
        if fp not in smry['checked_folders'].keys():
            continue  # skip the ones that didn't have matches
        if smry['checked_folders'][fp].get('move_item_rows') is None:
            lg.debug(f'No move items for {fp}')
            kirs = smry['checked_folders'][fp].get('keep_item_rows')
            if kirs is not None:
                lg.debug(f'{kirs=}')
            continue
        kirs = smry['checked_folders'][fp]['keep_item_rows']
        mirs = smry['checked_folders'][fp]['move_item_rows']
        for mirow in mirs:
            if not mirow:
                continue  # skip empty lists
            for idx, mirowrow in mirow:
                matched = False
                compare_columns = ('product_number', 'so_number', 'lot8', 'c_number')
                mrdf = series_to_df(mirowrow)
                match_cols = mrdf.loc[:, compare_columns]

                for kir in kirs:
                    compare_kir = series_to_df(kir).loc[:, compare_columns]
                    if not pd.merge(mrdf, series_to_df(kir), on=compare_columns, how='inner').empty:
                        matched = True
                        break
                if not matched:
                    unmatched.append(mirowrow)
                    lg.debug(f'No match found for {mirowrow}')
    if unmatched:
        lg.debug('UNMATCHED!!')
        lg.debug(unmatched)
    else:
        lg.debug('All e-mails to be moved have a matched kept e-mail.')

pass
