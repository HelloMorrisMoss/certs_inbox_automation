"""Tests to ensure the config files and template files are not out of sync."""

import json
import unittest

import untracked_config.accounts_and_folder_paths as afp
import untracked_config.accounts_and_folder_paths_template as afp_t
import untracked_config.auto_dedupe_cust_ids as cnums
import untracked_config.auto_dedupe_cust_ids_template as cnums_t
import untracked_config.development_node as odn
import untracked_config.development_node_template as odn_t
import untracked_config.foam_clean_product_names as fcpn
import untracked_config.foam_clean_product_names_template as fcpn_t
import untracked_config.priority_shipment_customers as psc
import untracked_config.priority_shipment_customers_template as psc_t
import untracked_config.scheduling_data as schd
import untracked_config.scheduling_data_template as schd_t
import untracked_config.subject_regex as sbjre
import untracked_config.subject_regex_template as sbjre_t


def test_sync(config_module, template_module):
    # Check that all variables in config module exist in template module
    for var_name in dir(config_module):
        if var_name.startswith("__"):  # don't test the built-ins
            continue

        # Check that variable exists in template module
        assert hasattr(template_module, var_name), f"Variable {var_name} is missing in template module"

        var_config = getattr(config_module, var_name)
        var_template = getattr(template_module, var_name)

        # Check that types match
        assert type(var_config) == type(var_template), f"Type of {var_name} is different in config and template module"

        # Check if variable is a dictionary
        if isinstance(var_config, dict):
            test_dict_sync(var_config, var_template, var_name)


def test_dict_sync(var_config, var_template, var_name=''):
    assert isinstance(var_template, dict), f"Type of {var_name} is different in config and template module"
    assert var_config.keys() == var_template.keys(), f"Keys of dictionary {var_name} don't match in config" \
                                                     f" and template module"
    for key in var_config.keys():
        assert type(var_config[key]) == type(var_template[
                                                 key]), f"Type of value for key {key} in dictionary {var_name} is" \
                                                        f" different in config and template module"


class TestSynchronization(unittest.TestCase):

    def test_accounts_paths(self):
        test_sync(afp, afp_t)

    def test_dedupe_cnums(self):
        test_sync(cnums, cnums_t)

    def test_on_dev_node(self):
        test_sync(odn, odn_t)

    def test_foam_clean_product_names(self):
        test_sync(fcpn, fcpn_t)

    def test_priority_shipment_customers(self):
        test_sync(psc, psc_t)

    def test_scheduling_data(self):
        test_sync(schd, schd_t)

    def test_subject_regex(self):
        test_sync(sbjre, sbjre_t)

    def test_email_json(self):
        prefix = '../'
        settings_json_filepath = f'{prefix}untracked_config/email_settings.json'
        settings_json_template_filepath = f'{prefix}untracked_config/email_settings_template.json'
        with open(settings_json_filepath, 'r') as settings_json_file, open(settings_json_template_filepath, 'r') as settings_json_template_file:
            emj = json.load(settings_json_file)
            emj_t = json.load(settings_json_template_file)
            test_dict_sync(emj, emj_t, 'email_settings_dict')


if __name__ == '__main__':
    unittest.main()
