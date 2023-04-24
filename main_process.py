import pandas as pd

from clean_foam_inbox import wc_outlook, get_process_folders_dfs, group_foam_mail, compare_keep_and_move, \
    get_mail_items_from_results, color_foam_groups, clear_test_foam_group_colors
from helpers.json_help import df_json_handler
from helpers.outlook_helpers import find_folders_in_outlook
from log_setup import lg
from untracked_config.accounts_and_folder_paths import acct_path_dct

if __name__ == '__main__':
    # ### this section is for development and demonstration only ###

    # pandas display settings for development
    pd.set_option('display.max_rows', 100)
    pd.set_option('display.max_columns', 100)
    pd.set_option('display.width', 1000)

    import json
    from helpers.outlook_helpers import valid_colors

    # a summary debug info dictionary
    smry = dict(checked_folders={}, skipped_folders=[], all_subj_lines=[], matched=[], missing_a_match=[])

    account_name = acct_path_dct['account_name']
    production_inbox_folders = acct_path_dct['inbox_folders']

    found_folders_dict = find_folders_in_outlook(wc_outlook, account_name, production_inbox_folders)
    pfdfs = get_process_folders_dfs(account_name, production_inbox_folders, found_folders_dict)
    # unmatched = []  # for checking for unmatched items
    testing_colors_move = ['grey']
    found_folders_keys = found_folders_dict.keys()

    for df, folder_path in pfdfs:
        if folder_path in found_folders_keys:
            move_item_rows, keep_item_rows, dfg = group_foam_mail(df, folder_path, smry)
            unmatched = compare_keep_and_move(move_item_rows, keep_item_rows, unmatched)
            move_items = get_mail_items_from_results(move_item_rows)
            color_foam_groups(dfg, move_items, move_item_color=testing_colors_move, valid_colors=valid_colors)
            pass
            clear_test_foam_group_colors(dfg, test_colors=valid_colors)
        else:
            lg.warn(f'Missing {folder_path} in checked folders!')

    if unmatched:
        lg.debug('UNMATCHED!!')
        lg.debug(unmatched)
    pass  # for breakpoint

    # write the smry dictionary to a file to make it easier to look at
    with open('./last_smry.json', 'w') as jf:
        json.dump(smry, jf, indent=4, default=df_json_handler)

# TODO: complete unit tests; next: a test confirming that the inbox looks like it does after "# color the groups"

pass  # for breakpoint
