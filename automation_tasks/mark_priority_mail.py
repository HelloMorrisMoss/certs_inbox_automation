from untracked_config.accounts_and_folder_paths import acct_path_dct

account_name = acct_path_dct['account_name']
production_inbox_folders = acct_path_dct['inbox_folders']

found_folders_dict = find_folders_in_outlook(outlook, acct_name, proc_folders)