from typing import Dict, List, Union

import numpy as np
import pandas as pd
import pypdf

program_performance_results_dict = {'unparsed_count': 0}


def page_text_to_coordinate_dataframe(rdr_page) -> Dict[str, pd.DataFrame]:
    """Converts the text and coordinate information of a PDF page into a pandas DataFrame.

    :param rdr_page: The PDF page object.
    :return: dict, A dictionary containing the resulting pandas DataFrame.
    """
    vd = {}
    vd['vdf'] = pd.DataFrame.from_dict({prm: [] for prm in ['text', 'cm', 'tm', 'font_dict', 'font_size']})

    def visitor_body(text, cm, tm, font_dict, font_size):
        vd['vdf'] = pd.concat([vd['vdf'],
                               pd.DataFrame.from_dict({'text': [text],
                                                       'cm': [cm],
                                                       'tm': [tm],
                                                       'font_dict': [font_dict],
                                                       'font_size': font_size,
                                                       })])

    rdr_page.extract_text(visitor_text=visitor_body)
    vd['vdf'][['tm_0', 'tm_1', 'tm_2', 'tm_3', 'tm_x', 'tm_y']] = vd['vdf']['tm'].apply(pd.Series)
    vd['vdf'][['cm_0', 'cm_1', 'cm_2', 'cm_3', 'cm_x', 'cm_y']] = vd['vdf']['cm'].apply(pd.Series)
    vd['vdf'][['base_font', 'encoding', 'subtype', 'type']] = vd['vdf']['font_dict'].apply(series_default_obj)
    vd['vdf'] = vd['vdf'].drop(['tm', 'cm'], axis=1)  # drop tm and cm columns, no longer needed
    vd['vdf'] = vd['vdf'].loc[~(vd['vdf']['text'] == '\n')]  # drop the rows that only contain new lines
    return vd


def series_default_obj(input_dict: Dict) -> pd.Series:
    """Create a pandas Series from the input dictionary with the dtype set as 'object'.

    :param input_dict: The input dictionary.
    :return: pd.Series, The resulting pandas Series.
    """
    return pd.Series(input_dict, dtype='object')


def get_row_by_left_header(coords_df: pd.DataFrame, left_header: str, tolerance: int = 1) -> Union[pd.DataFrame, None]:
    """Get rows from a DataFrame based on the left header value.

    :param coords_df: The DataFrame containing the rows.
    :param left_header: The left header value to search for.
    :param tolerance: The tolerance value used for comparison (default: 1).
    :return: pd.DataFrame or None, The rows matching the left header value or None if no rows match the condition.
    """
    y_coordinate = coords_df[coords_df['text'] == left_header].iloc[0, :]['tm_y']
    lh_row = get_tolerance_rows(coords_df, 'tm_y', y_coordinate, tolerance)
    return lh_row


def get_tolerance_rows(rows_df: pd.DataFrame, col_header: str, target_value: float, tolerance: float = 1) -> \
        pd.DataFrame:
    """Get rows where the numeric column is within the tolerance of the target value.

    :param rows_df: The DataFrame containing the rows to filter.
    :param col_header: The name of the column to compare.
    :param target_value: The target value to compare against.
    :param tolerance: The tolerance value used for comparison (default: 1).
    :return: pd.DataFrame or None, The filtered rows DataFrame or None if no rows match the condition.
    """
    lteq = rows_df[col_header] <= (target_value + tolerance)
    gteq = rows_df[col_header] >= (target_value - tolerance)
    lh_row: pd.DataFrame = rows_df[lteq & gteq]
    if isinstance(lh_row, tuple):  # todo: this check may no longer be needed
        lh_row = pd.DataFrame(lh_row[1])
    return lh_row


def get_test_results_dict_from_page(coords_df: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, str]]]:
    """Extracts test results information from a DataFrame and returns a dictionary containing the results.

    :param coords_df: The DataFrame containing the coordinates data.
    :return: dict, A dictionary containing the extracted test results information.
    """
    # todo: split this up
    test_results_dict: Dict[str, Dict[str, Dict[str, str]]] = {}
    coords_df = coords_df.sort_values('tm_y', ascending=False).copy().reindex()  # sort by vertical position

    # get the rows for date of manufacture, the results below their y value above any other belong to that lot
    dom_left_header = 'Date of Manufacturing (DOM):'
    mfr_dates = coords_df[coords_df['text'].str.contains(dom_left_header, regex=False)]

    coords_df['dom_y'] = np.nan  # initialize 'empty'
    coords_df.loc[coords_df['text'].str.contains(dom_left_header, regex=False), 'dom_y'] = coords_df.loc[
        coords_df['text'].str.contains(dom_left_header, regex=False), 'tm_y']  # set the dom rows' to their own y
    coords_df.loc[:, 'dom_y'] = coords_df.loc[:, 'dom_y'].fillna(method='ffill')  # fill the rest from the dom rows

    results_headers_left_header = 'Characteristic'
    result_header_rows = get_row_by_left_header(coords_df, results_headers_left_header)
    result_header_coords_df = result_header_rows[['text', 'tm_y', 'tm_x']]
    results_column_top_headers = ['Unit', 'Value', 'Lower Limit', 'Upper Limit']

    # loop through the rows that have test names
    for index, mfr_date in mfr_dates.iterrows():
        dom_n = mfr_date['text'].replace(dom_left_header, '').strip()  # the date of manufacture/lot number
        dom_y = mfr_date['tm_y']  # the y coordinate

        # filter the results for this dom
        y_mask = coords_df['dom_y'] == dom_y
        result_Mask = coords_df['text'].str.contains('initial')
        test_name_rows_df: pd.DataFrame = coords_df[result_Mask & y_mask]

        test_results_dict[dom_n] = {}  # add an entry for this dom

        for df_index, result_row_left_header_row in test_name_rows_df.iterrows():
            test_header = result_row_left_header_row['text']  # test name
            result_column_header_rows = get_row_by_left_header(coords_df.loc[y_mask],
                                                               test_header)  # row results column headers
            result_header_dict = {k: None for k in results_column_top_headers}

            # loop through the column headers
            for col_header in results_column_top_headers:
                # get values that match the x coordinate for this header
                _, header_y, header_x = result_header_coords_df[result_header_coords_df['text'] == col_header].values[0]
                result_value_rows = get_tolerance_rows(result_column_header_rows, 'tm_x', header_x, 1)
                value_row_count = len(result_value_rows)  # this hasn't been a problem yet
                if value_row_count == 1:  # only one result (should be this mostly)
                    result_value = result_value_rows.loc[0, 'text']
                elif value_row_count > 0:  # multiple results (shouldn't happen)
                    result_value = [value for value in result_value_rows['text']]
                    print(f'Multiple result rows found: {col_header}: {result_value=}')
                else:  # some don't have one, such as no upper/lower bound on some test ranges
                    result_value = 'None'
                result_header_dict[col_header] = result_value  # add this result to the dict
            test_results_dict[dom_n][test_header] = result_header_dict  # add this result dict to the dict under the dom
    return test_results_dict


def get_lot_info_dict(lot_info_page: pypdf.PageObject) -> Dict[str, str]:
    """Get a dictionary of lot information from the lot page of an NBE test report.

    :param lot_info_page: The lot_info_page object.
    :return: dict, A dictionary containing the extracted lot information.
    """
    # todo: move these config bits to untracked_config
    lot_keys: List[str] = ['Purchase Order / date:', 'Delivery / date:', 'Order / date: ', 'Customer number:',
                           'Material our / your reference:', 'Commercial Name:', 'Judgement :']
    lot_text_dict: Dict[str, str] = get_left_header_dict_from_page(lot_info_page, lot_keys)

    # use this to split combo lines into multiple dictionary entries; avoids a bunch of case-by-case if/then
    # tuples have a new keys tuple and then the old key: (('new', 'keys'), 'old key')
    split_guide = (
        (('po_number_nbe', 'po_date_nbe'), 'Purchase Order / date:'),
        (('order_number_nbe', 'order_date_nbe'), 'Order / date: '),
        (('product_number_nbe', 'tabcode_lw'), 'Material our / your reference:'),
        (('product_name',), 'Commercial Name:'),
        (('customer_number_nbe',), 'Customer number:'),
        (('delivery_number_nbe', 'delivery_date_nbe'), 'Delivery / date:'),
        (('judgement_nbe',), 'Judgement :')
        )

    # split into a new dictionary with the combo values on their own and (maybe) better names
    lot_info_dict: Dict[str, str] = {}
    for (new_keys, txt_key) in split_guide:
        old_value: str = lot_text_dict[txt_key]
        new_values: List[str] = old_value.rsplit('/', 1) if old_value is not None else ['not parsed'] * len(new_keys)
        if old_value is None:
            program_performance_results_dict['unparsed_count'] += 1
        for (nk, nv) in zip(new_keys, new_values):
            lot_info_dict[nk] = nv.strip()

    return lot_info_dict


def get_left_header_dict_from_page(lot_info_page, lot_keys: List[str]) -> Dict[str, str]:
    """Extracts information from a lot_info_page and returns a dictionary containing the left header values.

    :param lot_info_page: The lot_info_page object.
    :param lot_keys: A list of keys to search for in the page.
    :return: dict, A dictionary containing the extracted left header values.
    """
    lot_info_text: str = lot_info_page.extract_text()
    lot_text_dict = {k: None for k in lot_keys}
    split_text = lot_info_text.split('\n')
    while split_text:
        this_line = split_text.pop()
        for this_key in lot_keys:
            if this_key in this_line:  # remove extra chuff before adding to the results
                lot_text_dict[this_key] = this_line.replace(this_key, '').replace('\n', '').strip()
                break  # stop looking for this key
    return lot_text_dict


def get_test_results(page):
    visitor_dict: dict = page_text_to_coordinate_dataframe(page)
    vdf = visitor_dict['vdf']
    test_results: dict = get_test_results_dict_from_page(vdf)
    return test_results


def extract_nbe_report_data(reader: pypdf.PdfReader) -> Dict[str, dict]:
    """Extracts data from an NBE test report certificate PDF using PyPDF.

    example:
        reader = pypdf.PdfReader(path)
        report_data = extract_nbe_report_data(reader)
        pp(report_data)

    >{'lot_info': {'po_number_nbe': '123456',
              'po_date_nbe': '31.01.2023',  # d.m.y
              'order_number_nbe': '321123 / 000250',
              'order_date_nbe': '31.01.2023',  # d.m.y
              'product_number_nbe': 'CGP123_321_123,123456_ABC',
              'tabcode_lw': 'T8675309',
              'product_name': 'Tape-y-tape 9001',
              'customer_number_nbe': '1234',
              'delivery_number_nbe': '87654321 / 000010',
              'delivery_date_nbe': '31.05.2023',  # d.m.y
              'judgement_nbe': 'Passed'},
 'test_results': {'20230101': {'Total thickness initial ( 3 points )': {'Unit': 'µm',
                                                                        'Value': '50',
                                                                        'Lower Limit': '0',
                                                                        'Upper Limit': '100'},
                               # expect more results here
                               }}}

    :param reader: An instance of PyPDF.Reader representing the PDF reader object.
    :return: dict, A dictionary containing the extracted data.
    """
    pdf_data_dict: dict = {'lot_info': {}, 'test_results': {}}
    for pg_num, page in enumerate(reader.pages):
        if pg_num == 0:  # lot info page
            lot_info: dict = get_lot_info_dict(page)
            pdf_data_dict['lot_info'] = lot_info
        else:  # test results pages
            test_results = get_test_results(page)
            pdf_data_dict['test_results'].update(test_results)
    return pdf_data_dict


if __name__ == '__main__':
    from pprint import pprint, pp
    from os import scandir, path

    # pandas display settings for development
    pd.set_option('display.max_rows', 100)
    pd.set_option('display.max_columns', 100)
    pd.set_option('display.width', 1000)

    for cfd in scandir('../../untracked_sample_files/NBE_test_reports'):
        print(f'Processing: {cfd.name}')
        reader = pypdf.PdfReader(path.abspath(cfd.path))
        report_data = extract_nbe_report_data(reader)
        pp(report_data)
        print()
        break  # only one for a quick test

    pprint(program_performance_results_dict)

pass  # for debug breakpoint
