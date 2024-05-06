from turtle import delay
from typing import Dict, List, Union, Final

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
    lh_row: pd.DataFrame = rows_df.loc[lteq & gteq, :]
    if isinstance(lh_row, tuple):  # todo: this check may no longer be needed
        lh_row = pd.DataFrame(lh_row[1])
    return lh_row


def get_below_row_column(df: pd.DataFrame, search_col_header: str, row_contains: str, value_col_header: str,
                         new_col_header: str, in_place: bool = False):
    df = df.copy().reset_index(drop=True)
    # df = df.sort_values(value_col_header)
    df.loc[:, new_col_header] = np.nan  # initialize 'empty'
    contains_mask = df[search_col_header].str.contains(row_contains, regex=False)
    # only the rows with the row_contains text in the search_col_header column string
    contains_df = df.loc[contains_mask, :]

    for row_y in contains_df.loc[:, value_col_header]:  # loop through the y coordinate for the 'contain' rows
        chr_row_rows = get_tolerance_rows(df, value_col_header,
                                          row_y)  # get the rows with y coordinate values within 1 of the contain rows
        # set the new column's value for contain rows and those within 1 y value to the contain row's y coordinate
        df.loc[chr_row_rows.index, new_col_header] = row_y

    df.loc[:, new_col_header] = df.loc[:, new_col_header].fillna(method='ffill')  # fill the rest from above
    df.loc[:, new_col_header] = df.loc[:, new_col_header].fillna(-1)  # except for the header stuff
    return df.loc[:, new_col_header]


def get_test_results_dict_from_page(coords_df: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, str]]]:
    """Extracts test results information from a DataFrame and returns a dictionary containing the results.

    :param coords_df: The DataFrame containing the coordinates data.
    :return: dict, A dictionary containing the extracted test results information.
    """
    # todo: split this up
    test_results_dict: Dict[str, Dict[str, Dict[str, str]]] = {}
    coords_df = coords_df.sort_values('tm_y', ascending=False).copy().reset_index(
        drop=True)  # sort by vertical position

    # get the rows for date of manufacture, the results below their y value above any other belong to that lot
    dom_left_header = 'Date of Manufacturing (DOM):'
    # mfr_dates = coords_df[coords_df['text'].str.contains(dom_left_header, regex=False)]

    # todo: all off these different dataframes are only for development visibility
    dom_df = add_below_row_column(coords_df, 'text', dom_left_header, 'tm_y', 'dom_y')
    chr_df = add_below_row_column(dom_df, 'text', 'Characteristic', 'tm_y', 'chr_y')

    # add columns for PDF row groups and test result header/value groups
    diff_df: pd.DataFrame = chr_df.copy().reset_index(drop=True)
    diff_df['row_group'] = (abs(diff_df['tm_y'].diff()).fillna(method='bfill') > 1).cumsum()
    diff_df['test_group'] = (abs(diff_df['chr_y'].diff()).fillna(method='bfill') > 1).cumsum()
    diff_df = diff_df.drop(['font_dict', 'encoding', 'subtype', 'type'], axis=1)  # these are just noise atm
    diff_df = diff_df.sort_values('tm_y', ascending=False)

    # add a DoM column
    diff_df.loc[:, 'date_of_manufacture'] = ''
    for d_gn, dom_grp in diff_df.groupby('dom_y'):
        dom_mask = diff_df['dom_y'] == d_gn
        mfr_date = diff_df[diff_df['text'].str.contains(dom_left_header, regex=False) & dom_mask]
        if not mfr_date.empty:
            mfr_date = mfr_date.iloc[0, 0].replace(dom_left_header, '').strip()
            diff_df.loc[dom_mask, 'date_of_manufacture'] = mfr_date

    # add results
    results_column_top_headers: Final[List[str]] = ['Characteristic', 'Unit', 'Value', 'Lower Limit', 'Upper Limit']
    new_col_headers: List[str] = []
    for col_header in results_column_top_headers:
        new_hdr = f"{col_header.lower().replace(' ', '_')}_col"
        new_col_headers.append(new_hdr)
        if new_hdr not in diff_df.index:
            diff_df[new_hdr] = np.nan
        # get the unit col x
        unit_x = diff_df.loc[diff_df['text'].str.contains(col_header, regex=False), 'tm_x'].iloc[0]
        diff_df.loc[get_tolerance_rows(diff_df, 'tm_x', unit_x).index, new_hdr] = True
        diff_df[new_hdr].fillna(False, inplace=True)

    # create a dataframe of results
    results_df = pd.DataFrame(columns=new_col_headers, dtype='object')
    mfr_mask, text_mask = [diff_df[mask_col_hdr] != '' for mask_col_hdr in ['date_of_manufacture', 'text']]
    results_mask = diff_df['chr_y'] > 0
    group_columns = ['date_of_manufacture', 'test_group']
    for (mfr_date, g_index), group in diff_df[mfr_mask & text_mask & results_mask].groupby(group_columns):
        g_dict = {}  # {hdr: None for hdr in new_col_headers}
        for header, new_header in zip(results_column_top_headers, new_col_headers):
            g_dict['date_of_manufacture'] = [mfr_date]
            not_hdr = ~(group['text'].str.contains(header))
            hdr_col_true = group[new_header]
            try:
                result_rows = group.loc[hdr_col_true & not_hdr]
                g_dict[new_header] = [result_rows['text'].iloc[0]]
            except IndexError as ierr:
                print(f'No results for {mfr_date=}:{g_index=}:{header=} - {ierr}')
                g_dict[new_header] = [None]
        results_df = pd.concat([results_df, pd.DataFrame.from_dict(g_dict)])
    results_df = results_df.fillna('')
    return results_df


def add_below_row_column(df: pd.DataFrame, search_col_header: str, row_contains: str, value_col_header: str,
                         new_col_header: str, in_place: bool = False):
    """Add a column to a DataFrame that has values equal to the value column for its row.

    Searches for df-rows that contain the search string in the search column, then any df-rows with value-column values
    within tolerance for each have the new-column value set to that value-column value. Any not set are then filled from
    the last set-value in the new column.

    This allows grouping df-rows by clusters based on a pdf-row header.

    # Example data from PDF:

    Date of manufacture:    20230101
    Characteristic:                                 Value   Lower Limit
    Test name 1                                     5.00    1.00
    Characteristic:                                 Value   Lower Limit
    Test name 2                                     6.00    2.00

    # in this coordinate DataFrame:
    example_df.loc[test1_rows, :]
                        text  tm_0  tm_1  tm_2  tm_3     tm_x     tm_y
              Characteristic   8.0   0.0   0.0   8.0   45.355  595.252
                       Value   8.0   0.0   0.0   8.0  371.339  595.252
                        1.00  10.0   0.0   0.0  10.0  428.032  584.930
                 Test name 1  10.0   0.0   0.0  10.0   45.355  584.374
                        5.00  10.0   0.0   0.0  10.0  371.339  584.374
                 Lower Limit   8.0   0.0   0.0   8.0  428.032  569.141

    # searching for 'Characteristic', in the 'text' column, adding a 'chr_y' column, and setting the values from 'tm_y'
    add_below_row_column(dom_df, 'text', 'Characteristic', 'tm_y', 'chr_y')

        example_df.loc[example_rows, :]
                        text  tm_0  tm_1  tm_2  tm_3     tm_x     tm_y    chr_y
              Characteristic   8.0   0.0   0.0   8.0   45.355  595.252  595.252
                       Value   8.0   0.0   0.0   8.0  371.339  595.252  595.252
                        1.00  10.0   0.0   0.0  10.0  428.032  569.141  595.252
                 Test name 1  10.0   0.0   0.0  10.0   45.355  584.374  595.252
                        5.00  10.0   0.0   0.0  10.0  371.339  569.141  595.252
                 Lower Limit   8.0   0.0   0.0   8.0  428.032  584.374  595.252
                 # above this would represent one 'test result group' and below another
              Characteristic   8.0   0.0   0.0   8.0   45.355  554.930  554.930
                       Value   8.0   0.0   0.0   8.0  371.339  554.930  554.930
                        1.00  10.0   0.0   0.0  10.0  428.032  539.144  554.930
                 Test name 1  10.0   0.0   0.0  10.0   45.355  554.930  554.930
                        5.00  10.0   0.0   0.0  10.0  371.339  539.144  554.930
                 Lower Limit   8.0   0.0   0.0   8.0  428.032  554.930  554.930
    """
    if not in_place:
        df = df.copy().reset_index(drop=True)
    # df = df.sort_values(value_col_header)
    df.loc[:, new_col_header] = np.nan  # initialize 'empty'
    contains_mask = df[search_col_header].str.contains(row_contains, regex=False)
    # only the rows with the row_contains text in the search_col_header column string
    contains_df = df.loc[contains_mask, :]

    for row_y in contains_df.loc[:, value_col_header]:  # loop through the y coordinate for the 'contain' rows
        chr_row_rows = get_tolerance_rows(df, value_col_header,
                                          row_y)  # get the rows with y coordinate values within 1 of the contain rows
        # set the new column's value for contain rows and those within 1 y value to the contain row's y coordinate
        df.loc[chr_row_rows.index, new_col_header] = row_y

    df.loc[:, new_col_header] = df.loc[:, new_col_header].fillna(method='ffill')  # fill the rest from above
    df.loc[:, new_col_header] = df.loc[:, new_col_header].fillna(-1)  # except for the header stuff
    return df


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

    # fix the format of the DN that comes out of this part of the report
    delivery_number = lot_info_dict.get('delivery_number_nbe')
    if delivery_number is not None:
        lot_info_dict['delivery_number_nbe'] = delivery_number.replace(' / ', '').zfill(16)
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
            results_df = pdf_data_dict['test_results'].get('results_df')
            if results_df is None:  # the first page
                pdf_data_dict['test_results']['results_df'] = test_results
            else:  # subsequent pages
                pdf_data_dict['test_results']['results_df'] = pd.concat([results_df, test_results])
            # pdf_data_dict['test_results'].update(test_results)
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
