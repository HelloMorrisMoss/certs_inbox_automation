"""Check Outlook inboxes for redundant cert e-mails."""

import datetime
import re
from typing import Any, List, Optional, Tuple, Union

import pandas as pd
import win32com
from win32com import client as wclient

from helpers.outlook_helpers import add_categories_to_mail, colorize_outlook_email_list, \
    move_mail_items_to_folder, \
    remove_categories_from_mail
from log_setup import lg
from untracked_config.auto_dedupe_cust_ids import dedupe_columns
from untracked_config.development_node import ON_DEV_NODE, UNIT_TESTING
from untracked_config.subject_regex import subject_pattern


def process_mail_items(mail_items: list, summary_dict=None) -> tuple[List[dict[str, Any]], List[dict[str, Any]]]:
    """Processes the given mail items, extracting relevant information and returning a list of dictionaries.

    :param mail_items: A list of win32com CDispatch objects representing the mail items.
    :param summary_dict: dict, a dictionary for storing development/debugging information from the process.
    :return: List[dict], A list of dictionaries representing the mail items, with keys for 'received_time',
    'subject', and other
        extracted information.
    """
    # lists to populate
    results: List[dict[str, Any]] = []
    all_subj = []
    matched_sub = []
    non_regex_matching_emails = []

    # check if the subject line matches a generate cert mail
    for item in mail_items:
        subject: str = item.Subject
        all_subj.append(subject)
        match: re.Match = subject_pattern.match(subject)

        # pandas needs datetime.datetime not pywintypes.datetime; it's in UTC, thus the adjustment
        # todo: DST-proof this
        received_time = datetime.datetime.fromtimestamp(item.ReceivedTime.timestamp() + 14400)  # it's in UTC
        if received_time is None:
            lg.debug(f'No received time on {subject}')
            continue
        initial_row = {"received_time": received_time, "subject": subject, 'o_item': item}

        if match:
            matched_sub.append(subject)
            subj_dict = match.groupdict()

            row = initial_row | subj_dict
            results.append(row)
        else:
            non_regex_matching_emails.append(initial_row)

    if summary_dict is not None:  # recording for development
        summary_dict['all_subj_lines'] += all_subj
        summary_dict['matched'] += matched_sub
        summary_dict['non_regex_matching_emails']: List[Tuple[str, wclient.CDispatch]] = non_regex_matching_emails
    return results, non_regex_matching_emails


def sort_mail_items_to_dataframes(items: List[dict[str, Any]]) -> pd.DataFrame:
    """Get a dataframe sorted by received_time from a list of mail item dictionaries.

    :param items: List[dict[str, Any]], A list of dictionaries, each representing a mail item.
    :return: A pandas DataFrame containing the sorted mail items.
    """
    return pd.DataFrame(items).sort_values('received_time', axis=0, ascending=True).reset_index(drop=True)


def get_process_folders_dfs(proc_folders: List[str], folders_dict: dict = None,
                            summary_dict: dict = None) -> List[Tuple[pd.DataFrame, str]]:
    """Process mail items in a list of folders and returns a list of tuples, each containing a DataFrame with the mail
    items and the path of the folder it came from.

    :param proc_folders: List[str], the list of folder paths to process.
    :param folders_dict: dict, a dictionary containing the folders to process, indexed by their path.
    :param summary_dict: dict, a dictionary to store summary information about the mail items processed.
    :return: List[Tuple[pd.DataFrame, str]], a list of tuples, each containing a DataFrame with the mail items and
        the path of the folder it came from.
    """
    pf_dfs: List = []
    # get a dictionary of folders from the account
    for folder_path in proc_folders:
        olFolder = folders_dict.get(folder_path)
        if olFolder is None:
            lg.debug(f'{folder_path} was not found and will not be processed!')
            continue
        lg.debug(f'Processing folder: {folder_path}')

        items = olFolder.Items.Restrict('[FlagRequest] <> \'Follow up\'')  # exclude those already flagged
        if not ON_DEV_NODE:  # don't need a year's worth of e-mails each time in production, but test files will lag
            five_days_ago = datetime.datetime.now() - datetime.timedelta(days=5)
            date_filter = five_days_ago.strftime('%m/%d/%Y')
            filter_string = f'[ReceivedTime] >= \'{date_filter}\''
            items: List[wclient.CDispatch] = items.Restrict(filter_string)
        results, other_emails = process_mail_items(items)
        if results:
            df = sort_mail_items_to_dataframes(results)
            dfc = df.columns
            # in case there are no other emails, just use an empty dataframe
            other_emails_df = sort_mail_items_to_dataframes(other_emails) if other_emails else pd.DataFrame(columns=dfc)

            if not df.empty:
                df['lot8'] = df['lot_number'].str[:8]
                pf_dfs.append((folder_path, df, other_emails_df))
            else:
                lg.debug(f'No results in {folder_path}')
            if summary_dict is not None:
                summary_dict['checked_folders'][folder_path] = {'all_subj_lines': [], 'matched': [], 'dfs': df}
        else:
            lg.debug(f'No results in {folder_path}')
    return pf_dfs


def group_foam_mail(df: pd.DataFrame, folder_path: str, summary_dict: dict = None) -> \
        Tuple[List[Tuple[pd.Series]], List[Tuple[pd.Series]], pd.core.groupby.generic.DataFrameGroupBy]:
    """Groups mail items in the DataFrame by the specified columns and returns the groups, as well as the items to
    move and the items to keep for each group.

    :param df: pd.DataFrame, The DataFrame containing the mail items to group.
    :param folder_path: str, The path of the folder being processed.
    :param summary_dict: dict, A dictionary to which summary information will be added for each folder, defaults to
        None.
    :return: tuple, contains a tuple containing the items to move, the items to keep, and the DataFrameGroupBy object
        containing the mail items grouped by the specified columns.
    :rtype: Tuple[List[Tuple[pd.Series]], List[Tuple[pd.Series]], pd.core.groupby.generic.DataFrameGroupBy]
    """
    dfg: pd.DataFrame.groupby = df.groupby(dedupe_columns)
    keep_item_rows: list = []  # rows to keep in the mailbox
    move_item_rows: list = []  # rows to move from the mailbox

    name: tuple;
    grp: pd.DataFrame  # type hinting for the loop
    for name, grp in dfg:
        grp.sort_values(axis=0, by='cert_number', ascending=True, inplace=True)
        keep_item_rows.append([item_row for item_row in grp.iloc[:1].iterrows()])  # the first row (mail)
        move_item_rows.append([item_row for item_row in grp.iloc[1:].iterrows()])  # the rest of the rows

    if summary_dict:  # if working on development, store results for later examination
        if summary_dict['checked_folders'].get(folder_path) is None:
            summary_dict['checked_folders'][folder_path]: dict = {}
        summary_dict['checked_folders'][folder_path]['ibdf']: pd.DataFrame = df
        summary_dict['checked_folders'][folder_path]['dfg']: pd.DataFrame.groupby = dfg
        summary_dict['checked_folders'][folder_path]['keep_item_rows']: list = keep_item_rows
        summary_dict['checked_folders'][folder_path]['move_item_rows']: list = move_item_rows
    return move_item_rows, keep_item_rows, dfg


def series_to_df(srs: pd.Series) -> pd.DataFrame:
    """Convert a pandas series to a single-row DataFrame.

    :param srs: The pandas series to be converted.
    :return: A single-row DataFrame with the same data as the input series.
    """
    pass
    frm1 = pd.DataFrame.from_dict({k: [v] for k, v in srs.to_dict().items()})
    frm2 = srs.to_frame()
    assert frm1.equals(frm2)
    return frm1


def get_mail_items_from_results(list_of_series, o_item_col='o_item') -> list:
    """Extract the w32com.CDispatch.client Outlook mail item from the list of Pandas' series."""
    mail_items = []
    for p_row in list_of_series:
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


def compare_keep_and_move(mirs, kirs):
    unmatched: list = []
    for mirow in mirs:
        if not mirow:
            continue  # skip empty lists
        for idx, mirowrow in mirow:
            matched = False
            compare_columns = ('product_number', 'so_number', 'lot8', 'c_number')
            mrdf = mirowrow.to_frame().T
            # match_cols = mrdf.loc[:, compare_columns]

            for kir in kirs:
                if len(kir) > 1:
                    lg.warning(f'kir longer than 1: {kir}')
                active_kir = kir[0][1]
                missing_a_match_mail_df = pd.merge(mrdf, active_kir.to_frame().T, on=compare_columns, how='inner')
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
        for _, row in group_df.iterrows():
            o_item = row['o_item']
            add_categories_to_mail(o_item, color)


def process_foam_groups(df: pd.DataFrame, current_folder_path: str,
                        destination_folder: win32com.client.Dispatch, smry: Optional[dict] = None) -> None:
    """Move duplicate emails within a dataframe to a destination folder.

    This function groups the emails and identifies duplicates
    as those with the same subject, sender, and body. It then moves all but the first instance
    of each group to the destination folder.

    :param df: The DataFrame containing the emails to process.
    :param current_folder_path: The path of the folder to process.
    :param destination_folder: The destination folder to which duplicates will be moved.
    :param smry: A dictionary containing additional information for development purposes.
                 If provided, the function will color code the groups and items to move.
    :return: None.
    """

    # get lists of mail to move and leave and a pandas.DataFrame.GroupBy
    item_rows_to_move, item_rows_to_keep, dfg = group_foam_mail(df, current_folder_path, smry)

    # check for move mail without a keep
    unmatched_foam_rows: list = compare_keep_and_move(item_rows_to_move, item_rows_to_keep)
    if unmatched_foam_rows:
        lg.warn('Unmatched rows: %s', unmatched_foam_rows)
        raise RuntimeError(f'Unmatched rows found in {current_folder_path}')

    # get the mail items from the dataframe
    items_to_move: list = get_mail_items_from_results(item_rows_to_move)

    # for development, color code the groups and items to move
    if ON_DEV_NODE and not UNIT_TESTING:  # unit testing will put a copy in the unit test directory
        color_foam_groups(dfg, items_to_move, move_item_color=smry['testing_colors_move'],
                          valid_colors=smry['valid_colors'])
    # move the duplicates
    move_mail_items_to_folder(items_to_move, destination_folder)
    pass
