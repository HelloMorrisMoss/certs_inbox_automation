"""Check Outlook inboxes for redundant cert e-mails."""

import datetime
from pprint import pprint
from typing import List, Dict, Any

import pandas as pd
import win32com

from helpers.json_help import df_json_handler
from helpers.outlook_helpers import add_categories_to_mail, find_folders_in_outlook, remove_categories_from_mail
from log_setup import lg
from outlook_interface import wc_outlook
from untracked_config.accounts_and_folder_paths import acct_path_dct
from untracked_config.subject_regex import subject_pattern

now = datetime.datetime.now()
lg.debug(f'Starting at {now}')

wc_outlook = wc_outlook.get_outlook_folders()


def get_mail_items_from_inbox(olFolder: win32com.client.CDispatch) -> List[win32com.client.CDispatch]:
    """
    Returns a list of mail items in the specified Outlook folder.

    :param olFolder: The Outlook folder to retrieve mail items from.
    :return: A list of win32com CDispatch objects representing the mail items in the folder.
    """
    return olFolder.Items


# def sort_mail_items_by_subject(mail_items: List[win32com.client.CDispatch], subject_regex: subject_pattern) -> Tuple[
#     List[Dict[str, Any]], List[str]]:
#     """
#     Sorts a list of mail items into two lists based on whether their subject line matches the specified regex.
#
#     :param mail_items: A list of win32com CDispatch objects representing the mail items to sort.
#     :param subject_regex: A compiled regular expression pattern to match against the mail item subject lines.
#     :return: A tuple containing a list of dictionaries representing the matched mail items, and a list of all subject lines.
#     """
#     matched_results = []
#     all_subject_lines = []
#     for item in mail_items:
#         subject = item.Subject
#         all_subject_lines.append(subject)
#         match = subject_regex.match(subject)
#         if match:
#             subj_dict: dict = match.groupdict()
#             if subj_dict['product_number'] in product_names and subj_dict['c_number'] in dedupe_cnums:
#                 # pandas needs datetime.datetime not pywintypes.datetime
#                 received_time = datetime.datetime.fromtimestamp(item.ReceivedTime.timestamp() + 14400)  # it's in UTC
#                 if received_time is None:
#                     lg.debug(f'No rec time on {subject}')
#                     continue
#                 row = {"received_time": received_time, "subject": subject, 'o_item': item} | subj_dict
#                 matched_results.append(row)
#
#     return matched_results, all_subject_lines


# def process_mail_items(mail_items: List[win32com.client.CDispatch], subject_regex: subject_pattern) -> List[
#     Dict[str, Any]]:
#     """
#     Processes a list of mail items by filtering and extracting relevant information.
#
#     :param mail_items: A list of win32com CDispatch objects representing the mail items to process.
#     :param subject_regex: A compiled regular expression pattern to match against the mail item subject lines.
#     :return: A list of dictionaries representing the processed mail items.
#     """
#     results = []  # list of
#     for item in mail_items:
#         subject = item.Subject
#         match = subject_regex.match(subject)
#         if match:
#             subj_dict: dict = match.groupdict()
#             # if subj_dict['product_number'] in product_names and subj_dict['c_number'] in dedupe_cnums:
#             # pandas needs datetime.datetime not pywintypes.datetime
#             received_time = datetime.datetime.fromtimestamp(item.ReceivedTime.timestamp() + 14400)  # it's in UTC
#             if received_time is None:
#                 lg.debug(f'No rec time on {subject}')
#                 continue
#             row = {"received_time": received_time, "subject": subject, 'o_item': item} | subj_dict
#             results.append(row)
#
#     return results


def process_mail_items(mail_items: list) -> list:
    """Processes the given mail items, extracting relevant information and returning a list of dictionaries.

    :param mail_items: A list of win32com CDispatch objects representing the mail items.
    :return: A list of dictionaries representing the mail items, with keys for 'received_time', 'subject', and other
        extracted information.
    """
    results: List[Dict[str, Any]] = []
    all_subj: List[str] = []
    matched_sub: List[str] = []
    for item in mail_items:
        subject = item.Subject
        all_subj.append(subject)
        match = subject_pattern.match(subject)
        if match:
            matched_sub.append(subject)
            subj_dict: dict = match.groupdict()
            # if subj_dict['product_number'] in product_names:
            #
            #     if subj_dict['c_number'] not in dedupe_cnums:  # only certain customers
            #         continue
                # pandas needs datetime.datetime not pywintypes.datetime
            received_time = datetime.datetime.fromtimestamp(item.ReceivedTime.timestamp() + 14400)  # it's in UTC
            if received_time is None:
                lg.debug(f'No rec time on {subject}')
                continue
            row = {"received_time": received_time, "subject": subject, 'o_item': item} | subj_dict
            results.append(row)

    smry['all_subj_lines'] += all_subj
    smry['matched'] += matched_sub
    return results


def sort_mail_items_to_dataframes(items, subject_pattern):
    # results = []
    # all_subj = []
    # matched_sub = []
    # for item in items:
    #     subject = item.Subject
    #     all_subj.append(subject)
    #     match = subject_pattern.match(subject)
    #     if match:
    #         results.append(row)
    #         matched_sub.append(subject)
    #         subj_dict: dict = match.groupdict()
    #         if subj_dict['product_number'] in product_names:
    #             received_time = datetime.datetime.fromtimestamp(item.ReceivedTime.timestamp() + 14400)  # it's in UTC
    #             if received_time is None:
    #                 lg.debug(f'No rec time on {subject}')
    #                 continue
    #             row = {"received_time": received_time, "subject": subject, 'o_item': item} | subj_dict
    #
    # smry['all_subj_lines'] += all_subj
    # smry['matched'] += matched_sub
    return pd.DataFrame(items).sort_values('received_time', axis=0, ascending=True).reset_index()


def get_process_folders_dfs(acct_name: str, proc_folders: list, folders_dict=None):
    pf_dfs = []
    # get a dictionary of folders from the account
    for folder_path in proc_folders:
        olFolder = folders_dict.get(folder_path)
        if olFolder is None:
            lg.debug(f'{folder_path} was not found and will not be processed!')
            continue
        lg.debug(f'Processing folder: {folder_path}')
        smry['checked_folders'][folder_path] = {'all_subj_lines': [], 'matched': []}
        items = olFolder.Items
        results = process_mail_items(items)
        if results:
            df = sort_mail_items_to_dataframes(results, subject_pattern)

            if df.empty:
                lg.debug(f'No results in {folder_path}')
                continue
            smry['checked_folders'][folder_path]['dfs'] = df
            df['lot8'] = df['lot_number'].str[:8]
            pf_dfs.append((df, folder_path))
        else:
            lg.debug(f'No results in {folder_path}')
            continue
    return pf_dfs

# def main_folders_process(acct_name: str, proc_folders: list, folders_dict=None):
#     # get a dictionary of folders from the account
#     for folder_path in proc_folders:
#         olFolder = folders_dict.get(folder_path)
#         if olFolder is None:
#             lg.debug(f'{folder_path} was not found and will not be processed!')
#             continue
#         lg.debug(f'Processing folder: {folder_path}')
#         smry['checked_folders'][folder_path] = {'all_subj_lines': [], 'matched': []}
#         items = get_mail_items_from_inbox(olFolder)
#         results = process_mail_items(items)
#         if results:
#             dfs = sort_mail_items_to_dataframes(results)
#             smry['checked_folders'][folder_path]['dfs'] = dfs
#         else:
#             lg.debug(f'No results in {folder_path}')
#             continue

def group_foam_mail(df, folder_path):
    dfg = df.groupby(['product_number', 'so_number', 'lot8', 'c_number'])
    keep_item_rows = []
    move_item_rows = []
    for grp in dfg:
        keep_item_rows.append([item_row for item_row in grp[1].iloc[:1].iterrows()])  # .append([grp[1].iloc[0:1]])
        move_item_rows.append([item_row for item_row in grp[1].iloc[1:].iterrows()])
    smry['checked_folders'][folder_path]['ibdf'] = df
    smry['checked_folders'][folder_path]['dfg'] = dfg
    smry['checked_folders'][folder_path]['keep_item_rows'] = keep_item_rows
    smry['checked_folders'][folder_path]['move_item_rows'] = move_item_rows


def series_to_df(srs):
    return pd.DataFrame.from_dict({k: [v] for k, v in srs.to_dict().items()})


def colorize_series(mail_items: list, color: str):
    """Add the color category to all of the mail items in the list.

    :param mail_items:
    :param color:
    """
    for mail_item in mail_items:
        add_categories_to_mail(mail_item, color)


def get_mail_items_from_results(list_of_series) -> list:
    """Extract the w32com.CDispatch.client Outlook mail item from the list of Pandas' series."""
    mail_items = []
    for p_row in list_of_series:
        if len(p_row) > 1:
            lg.warning(f'mirs row list had more than a single item.')
            pprint(p_row)
        for rlist in p_row:
            mail_items.append(rlist[1]['o_item'])
    return mail_items


def clear_testing_colors(testing_series: pd.Series, testing_colors: list) -> None:
    """Remove the color categories from the mail items used in testing.

    :param testing_series: series, the series that includes the mail items.
    :param testing_colors: list, the colors to remove from the mail items.
    """
    mitems = get_mail_items_from_results(testing_series)
    for mi in mitems:
        remove_categories_from_mail(mi, testing_colors)


if __name__ == '__main__':
    # pandas display settings for development
    pd.set_option('display.max_rows', 100)
    pd.set_option('display.max_columns', 100)
    pd.set_option('display.width', 1000)

    # write the smry dictionary to a file to make it easier to look at
    import json

    account_name = acct_path_dct['account_name']
    production_inbox_folders = acct_path_dct['inbox_folders']
    # a summary debug info dictionary
    smry = dict(checked_folders={}, skipped_folders=[], all_subj_lines=[], matched=[], missing_a_match=[])

    found_folders_dict = find_folders_in_outlook(wc_outlook, account_name, production_inbox_folders)
    # main_folders_process(acct_name=account_name, proc_folders=production_inbox_folders, folders_dict=found_folders_dict)
    pfdfs = get_process_folders_dfs(account_name, production_inbox_folders, found_folders_dict)
    for df, folder_path in pfdfs:
        group_foam_mail(df, folder_path)

    # ### this section is for development and demonstration only ###
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
                    if len(kir) > 1:
                        lg.warning(f'kir longer than 1: {kir}')
                    active_kir = kir[0][1]
                    compare_kir = series_to_df(active_kir).loc[:, compare_columns]
                    missing_a_match_mail_df = pd.merge(mrdf, series_to_df(active_kir), on=compare_columns, how='inner')
                    df_is_empty = missing_a_match_mail_df.empty  # if it is empty, then there are no missing rows

                    if not df_is_empty:
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

        testing_colors_move = ['grey']
        # testing_colors_keep = ['pink']
        #
        move_items = get_mail_items_from_results(mirs)
        colorize_series(move_items, testing_colors_move)
        # colorize_series(kirs, testing_colors_keep)
        #
        # if 'y' in input('clear testing colors?'):
        #     # clear testing colors
        #     clear_testing_colors(mirs, testing_colors_move)
        #     clear_testing_colors(kirs, testing_colors_keep)
        from helpers.outlook_helpers import valid_colors

        # color the groups
        dfg = smry['checked_folders']['\\\\SB-certs\\1-CERTS Inbox\\Automation Testing\\active_files\\Inbox']['dfg']
        colorize_series(move_items, testing_colors_move)
        for color, (group_name, group_df) in zip(valid_colors, dfg):
            print(f'{color=}, {len(group_df)}')
            for _, row in group_df.iterrows():
                o_item = row['o_item']

                print(f'{o_item.Subject=}, {o_item.Categories=}')
                add_categories_to_mail(o_item, color)

        # clear the category color groups
        pass
        for color, (group_name, group_df) in zip(valid_colors, dfg):
            for _, row in group_df.iterrows():
                o_item = row['o_item']
                o_item.Categories = ''
                o_item.Save()

# TODO: complete unit tests; next: a test confirming that the inbox looks like it does after "# color the groups"

pass
