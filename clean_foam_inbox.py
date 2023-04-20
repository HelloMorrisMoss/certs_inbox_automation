"""Check Outlook inboxes for redundant cert e-mails."""

import datetime
from pprint import pprint

import pandas as pd
import pythoncom
from win32com import client as wclient

from helpers.json_help import df_json_handler
from helpers.outlook_helpers import add_categories_to_mail, find_folders_in_outlook, remove_categories_from_mail
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
def cluster_mail_items_by_time(olFolder, folder_path):
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
    smry['checked_folders'][folder_path]['results'] = results
    return results


def series_to_df(srs):
    return pd.DataFrame.from_dict({k: [v] for k, v in srs.to_dict().items()})


def main_folders_process(acct_name: str, proc_folders: list):

    # get a dictionary of folders from the account
    found_folders_dict = find_folders_in_outlook(outlook, acct_name, proc_folders)
    for folder_path in proc_folders:
        olFolder = found_folders_dict.get(folder_path)
        if olFolder is None:
            lg.debug(f'{folder_path} was not found and will not be processed!')
            continue
        lg.debug(f'Processing folder: {folder_path}')
        smry['checked_folders'][folder_path] = {'all_subj_lines': [], 'matched': []}
        ibdf = cluster_mail_items_by_time(olFolder, folder_path)

        if ibdf.empty:
            lg.debug(f'No results in {folder_path}')
            continue

        dfg = ibdf.groupby(['product_number', 'so_number', 'lot8', 'c_number'])
        keep_item_rows = []
        move_item_rows = []
        for grp in dfg:
            keep_item_rows.append([item_row for item_row in grp[1].iloc[:1].iterrows()])  #.append([grp[1].iloc[0:1]])
            move_item_rows.append([item_row for item_row in grp[1].iloc[1:].iterrows()])
        smry['checked_folders'][folder_path]['ibdf'] = ibdf
        smry['checked_folders'][folder_path]['dfg'] = dfg
        smry['checked_folders'][folder_path]['keep_item_rows'] = keep_item_rows
        smry['checked_folders'][folder_path]['move_item_rows'] = move_item_rows


def colorize_series(mail_items, color):
    for mail_item in mail_items:
        add_categories_to_mail(mail_item, color)


def get_mail_items_from_results(list_of_series):
    """Extract the w32com.CDispatch.client Outlook mail item from the list of Pandas' series."""
    mail_items = []
    for row in list_of_series:
        if len(row) > 1:
            lg.warning(f'mirs row list had more than a single item.')
            pprint(row)
        for rlist in row:
            mail_items.append(rlist[1]['o_item'])
    return mail_items


def clear_testing_colors(testing_series, testing_colors):
    mitems = get_mail_items_from_results(testing_series)
    for mi in mitems:
        remove_categories_from_mail(mi, testing_colors)


if __name__ == '__main__':
    # display settings for development
    pd.set_option('display.max_rows', 100)
    pd.set_option('display.max_columns', 100)
    pd.set_option('display.width', 1000)

    # write the smry dictionary to a file to make it easier to look at
    import json

    # global production_inbox_folders, smry
    account_name = acct_path_dct['account_name']
    production_inbox_folders = acct_path_dct['inbox_folders']
    # a summary debug info dictionary
    smry = dict(checked_folders={}, skipped_folders=[], all_subj_lines=[])

    main_folders_process(account_name, production_inbox_folders)

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
                    if not pd.merge(mrdf, series_to_df(active_kir), on=compare_columns, how='inner').empty:
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
