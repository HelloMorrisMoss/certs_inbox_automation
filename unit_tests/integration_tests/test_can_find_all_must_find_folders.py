import unittest
from unittest import TestCase
from unittest.mock import patch


from outlook_interface import wc_outlook


class test_finding_folders(TestCase):
    @patch('untracked_config.development_node.ON_DEV_NODE', False)  # test for the actual folders
    def test_final_state_folders(self):
        from untracked_config.accounts_and_folder_paths import acct_path_dct
        import main_process  # this must be imported within the patched scope
        found_folders_dict, production_inbox_folders = main_process.get_process_ol_folders(wc_outlook)
        must_find_folders_tst = main_process.get_must_find_folders()
        mff_summary_tst = {}
        ff_list_tst = list(found_folders_dict.keys())
        for mff in must_find_folders_tst:
            mff_summary_tst[mff] = mff in ff_list_tst
        assert all(mff_summary_tst.values()), 'Not all must_find_folders were found: ' + str(mff_summary_tst)


if __name__ == '__main__':
    unittest.main()
