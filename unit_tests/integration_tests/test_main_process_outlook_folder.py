import unittest

from main_process import get_process_ol_folders, main_process_function, wc_outlook
from untracked_config.accounts_and_folder_paths import acct_path_dct


def test_email_properties(control_folder, test_folder):
    """Tests that emails in the control and test folders have identical properties."""
    control_emails = control_folder.Items
    test_emails = test_folder.Items

    # Create a list of dictionaries containing email properties
    control_email_props = []
    for email in control_emails:
        props = {
            "subject": email.Subject,
            "follow_up_state": email.FlagRequest,
            "follow_up_status": email.FlagStatus,
            }
        control_email_props.append(props)

    test_email_props = []
    for email in test_emails:
        props = {
            "subject": email.Subject,
            "follow_up_state": email.FlagRequest,
            "follow_up_status": email.FlagStatus,
            }
        test_email_props.append(props)

    # Match emails based on unique properties (subject, sender, and creation time)
    for control_props in control_email_props:
        matching_test_props = None
        for test_props in test_email_props:
            if (all([control_props[prop_key] == test_props[prop_key] for prop_key in control_props.keys()])
            ):
                matching_test_props = test_props
                break

        assert matching_test_props is not None, f"Email not matched in test folder: {control_props}"

        # Compare properties of matching emails  - TODO: is this part redundant?
        for property_key in control_props.keys():
            assert control_props[property_key] == matching_test_props[property_key], f'Mismatched property ' \
                                                                                     f'{property_key} for ' \
                                                                                     f'{control_props}'

        # Remove test props from list
        test_email_props.remove(matching_test_props)

    # Check that all test emails have been used
    assert len(test_email_props) == 0, f"{len(test_email_props)} test emails not found in control folder"


def test_final_state_folders():
    found_folders_dict, production_inbox_folders = get_process_ol_folders(wc_outlook)

    # test folders
    test_inbox_path = acct_path_dct["inbox_folders"][0]
    tst_path = r'automation testing\unit_test_files'
    # to help protect against accidentally running against production folder
    assert tst_path in test_inbox_path.lower(), f'''Inbox doesn't match test folder pattern: {test_inbox_path}'''

    inbox_folder = found_folders_dict[test_inbox_path]
    target_folder = found_folders_dict[acct_path_dct['target_folder_path']]

    # test file storage folder
    test_file_origin = found_folders_dict[acct_path_dct["test_file_origin"]]

    # Retrieve final state of folders
    known_good_final_state_inbox_folder = found_folders_dict[acct_path_dct['known_good_final_state_inbox_folder']]
    known_good_final_state_move_to_folder = found_folders_dict[acct_path_dct['known_good_final_state_move_to_folder']]

    # clear old items
    for t_folder in [inbox_folder, target_folder]:
        tries = 10  # it looks like duplicated mail (due to interrupted testing) doesn't all delete on the first try
        while tries and t_folder.Items:
            tries -= 1
            for item in t_folder.Items:
                item.Delete()

    for item in test_file_origin.Items:
        new_copy = item.Copy()
        new_copy.Move(inbox_folder)

    # Run program
    main_process_function(found_folders_dict, production_inbox_folders)

    # Compare before and after folders
    test_email_properties(known_good_final_state_inbox_folder, inbox_folder)
    test_email_properties(known_good_final_state_move_to_folder, target_folder)


class TestFolderFinalStates(unittest.TestCase):
    def test_final_states(self):
        test_final_state_folders()


if __name__ == '__main__':
    unittest.main()
