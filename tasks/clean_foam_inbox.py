"""Check Outlook inboxes for redundant cert e-mails."""

import datetime
from pprint import pprint
from typing import Any, Dict, List

import pandas as pd
import win32com

from helpers.outlook_helpers import add_categories_to_mail, colorize_outlook_email_list, move_mail_items_to_folder, \
    remove_categories_from_mail
from log_setup import lg
from untracked_config.subject_regex import subject_pattern


def process_mail_items(mail_items: list, summary_dict=None) -> list:
    """Processes the given mail items, extracting relevant information and returning a list of dictionaries.

    :param mail_items: A list of win32com CDispatch objects representing the mail items.
    :param summary_dict: dict, a dictionary for storing development/debugging information from the process.
    :return: A list of dictionaries representing the mail items, with keys for 'received_time', 'subject', and other
        extracted information.
    """
    results: List[Dict[str, Any]] = []
    all_subj: List[str] = []
    matched_sub: List[str] = []
    non_regex_matching_emails = []
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
        else:
            non_regex_matching_emails.append((subject, item))
    if summary_dict is not None:
        summary_dict['all_subj_lines'] += all_subj
        summary_dict['matched'] += matched_sub
        summary_dict['non_regex_matching_emails'] = non_regex_matching_emails
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


def get_process_folders_dfs(acct_name: str, proc_folders: list, folders_dict=None, summary_dict=None):
    pf_dfs = []
    # get a dictionary of folders from the account
    for folder_path in proc_folders:
        olFolder = folders_dict.get(folder_path)
        if olFolder is None:
            lg.debug(f'{folder_path} was not found and will not be processed!')
            continue
        lg.debug(f'Processing folder: {folder_path}')
        items = olFolder.Items
        results = process_mail_items(items)
        if results:
            df = sort_mail_items_to_dataframes(results, subject_pattern)

            if not df.empty:
                df['lot8'] = df['lot_number'].str[:8]
                pf_dfs.append((df, folder_path))
            else:
                lg.debug(f'No results in {folder_path}')
            if summary_dict is not None:
                summary_dict['checked_folders'][folder_path] = {'all_subj_lines': [], 'matched': [], 'dfs': df}
        else:
            lg.debug(f'No results in {folder_path}')
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

def group_foam_mail(df, folder_path, summary_dict=None):
    dfg = df.groupby(['product_number', 'so_number', 'lot8', 'c_number'])
    keep_item_rows = []
    move_item_rows = []
    for grp in dfg:
        keep_item_rows.append([item_row for item_row in grp[1].iloc[:1].iterrows()])  # .append([grp[1].iloc[0:1]])
        move_item_rows.append([item_row for item_row in grp[1].iloc[1:].iterrows()])
    if summary_dict:
        if summary_dict['checked_folders'].get(folder_path) is None:
            summary_dict['checked_folders'][folder_path] = {}
        summary_dict['checked_folders'][folder_path]['ibdf'] = df
        summary_dict['checked_folders'][folder_path]['dfg'] = dfg
        summary_dict['checked_folders'][folder_path]['keep_item_rows'] = keep_item_rows
        summary_dict['checked_folders'][folder_path]['move_item_rows'] = move_item_rows
    return move_item_rows, keep_item_rows, dfg

def series_to_df(srs):
    return pd.DataFrame.from_dict({k: [v] for k, v in srs.to_dict().items()})


def get_mail_items_from_results(list_of_series, o_item_col='o_item') -> list:
    """Extract the w32com.CDispatch.client Outlook mail item from the list of Pandas' series."""
    mail_items = []
    for p_row in list_of_series:
        if len(p_row) > 1:
            lg.warning(f'mirs row list had more than a single item.')
            pprint(p_row)
        for rlist in p_row:
            mail_items.append(rlist[1][o_item_col])
    return mail_items


def clear_testing_colors(testing_series: pd.Series, testing_colors: list) -> None:
    """Remove the color categories from the mail items used in testing.

    :param testing_series: series, the series that includes the mail items.
    :param testing_colors: list, the colors to remove from the mail items.
    """
    mitems = get_mail_items_from_results(testing_series)
    for mi in mitems:
        remove_categories_from_mail(mi, testing_colors)


def compare_keep_and_move(mirs, kirs, unmatched):
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
                # compare_kir = series_to_df(active_kir).loc[:, compare_columns]
                missing_a_match_mail_df = pd.merge(mrdf, series_to_df(active_kir), on=compare_columns, how='inner')
                df_is_empty = missing_a_match_mail_df.empty  # if it is empty, then there are no missing rows

                if not df_is_empty:
                    matched = True
                    break
            if not matched:
                unmatched.append(mirowrow)
                lg.debug(f'No match found for {mirowrow}')
    return unmatched


def color_foam_groups(dfg, move_items, move_item_color, valid_colors):
    colorize_outlook_email_list(move_items, move_item_color)
    for color, (group_name, group_df) in zip(valid_colors, dfg):
        print(f'{color=}, {len(group_df)}')
        for _, row in group_df.iterrows():
            o_item = row['o_item']

            print(f'{o_item.Subject=}, {o_item.Categories=}')
            add_categories_to_mail(o_item, color)


def process_foam_groups(df, current_folder_path: str, unmatched_foam_rows, testing_colors_move,
                        valid_colors, destination_folder: win32com.client.Dispatch, smry=None):
    # get lists of mail to move and leave and a pandas.DataFrame.GroupBy
    item_rows_to_move, item_rows_to_keep, dfg = group_foam_mail(df, current_folder_path, smry)

    # check for move mail without a keep
    unmatched_foam_rows: list = compare_keep_and_move(item_rows_to_move, item_rows_to_keep, unmatched_foam_rows)
    if unmatched_foam_rows:
        lg.warn('Unmatched rows: %s', unmatched_foam_rows)
        raise RuntimeError('Unmatched rows found in ')

    # get the mail items from the dataframe
    move_items: list = get_mail_items_from_results(item_rows_to_move)
    # for development, color code the groups and items to move
    # color_foam_groups(dfg, move_items, move_item_color=testing_colors_move, valid_colors=valid_colors)

    # move the duplicates
    move_mail_items_to_folder(move_items, destination_folder)
    pass
    # clear_all_category_colors_foam(dfg)
    return unmatched_foam_rows
