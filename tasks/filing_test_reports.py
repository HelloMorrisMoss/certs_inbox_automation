import datetime
from pprint import pprint

import pandas as pd
import pypdf

# pandas display settings for development
pd.set_option('display.max_rows', 100)
pd.set_option('display.max_columns', 100)
pd.set_option('display.width', 1000)

results_dict = {'unparsed_count': 0}

reader = pypdf.PdfReader("../untracked_sample_files/NBE_test_reports/Certificate for Delivery0080615826000010.PDF")

# header_df = pd.DataFrame.from_dict({'Unit': [], 'Value': [], 'Lower Limit': [], 'Upper Limit': []})
lot_info_page = reader.pages[0]
test_results_page = reader.pages[1]


def page_text_to_coordinate_dataframe(rdr_page):
    vd = {}
    vd['vdf'] = pd.DataFrame.from_dict({prm: [] for prm in ['text', 'cm', 'tm', 'font_dict', 'font_size']})

    def visitor_body(text, cm, tm, font_dict, font_size):
        # global visitor_df
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
    vd['vdf'][['base_font', 'encoding', 'subtype', 'type']] = vd['vdf']['font_dict'].apply(
        series_default_obj)
    vd['vdf'] = vd['vdf'].drop(['tm', 'cm'], axis=1)  # drop tm and cm columns, no longer needed
    vd['vdf'] = vd['vdf'].loc[~(vd['vdf']['text'] == '\n')]  # drop the rows that only contain new lines
    return vd


def series_default_obj(input_dict):
    return pd.Series(input_dict, dtype='object')


def get_row_by_left_header(coords_df: pd.DataFrame, left_header: str, tolerance: int = 1):
    y_coordinate = coords_df[coords_df['text'] == left_header].iloc[0, :]['tm_y']
    lh_row = get_tolerance_rows(coords_df, 'tm_y', y_coordinate, tolerance)
    return lh_row


def get_tolerance_rows(rows_df: pd.DataFrame, col_header: str, target_value: float, tolerance=1):
    """Get rows where the numeric column is within the tolerance of the target value.

    :param rows_df:
    :param col_header:
    :param target_value:
    :param tolerance:
    :return:
    """
    lteq = (rows_df[col_header] <= (target_value + tolerance))
    gteq = (rows_df[col_header] >= (target_value - tolerance))
    lh_row = rows_df[(lteq & gteq)]
    if isinstance(lh_row, tuple):
        lh_row = pd.DataFrame[lh_row[1]]
    return lh_row


def get_test_results_dict_from_page(coords_df:pd.DataFrame):
    test_results_dict = {}
    results_headers_left_header = 'Characteristic'
    result_header_rows = get_row_by_left_header(coords_df, results_headers_left_header)
    result_header_coords_df = result_header_rows[['text', 'tm_y', 'tm_x']]
    results_column_top_headers = ['Unit', 'Value', 'Lower Limit', 'Upper Limit']

    # loop through the rows that have test names
    test_name_rows_df: pd.DataFrame = coords_df[coords_df['text'].str.contains('initial')]
    for df_index, result_row_left_header_row in test_name_rows_df.iterrows():
        test_header = result_row_left_header_row['text']  # test name
        result_column_header_rows = get_row_by_left_header(coords_df, test_header)  # row results column headers
        result_header_dict = {k: None for k in results_column_top_headers}

        # loop through the column headers
        for col_header in results_column_top_headers:
            _, header_y, header_x = result_header_coords_df[result_header_coords_df['text'] == col_header].values[0]
            # get values that match the x coordinate for this header
            result_value_rows = get_tolerance_rows(result_column_header_rows, 'tm_x', header_x, 1)
            value_row_count = len(result_value_rows)
            if value_row_count == 1:
                result_value = result_value_rows.loc[0, 'text']
            elif value_row_count > 0:
                result_value = [value for value in result_value_rows['text']]
            else:
                result_value = 'None'
            result_header_dict[col_header] = result_value
            test_results_dict[test_header] = result_header_dict
    return test_results_dict



visitor_dict: dict = page_text_to_coordinate_dataframe(test_results_page)
vdf = visitor_dict['vdf']
test_results: dict = get_test_results_dict_from_page(vdf)


# # better to get this part from page.extract_text
# report_date_left_header = 'Date:'
#
#
# def get_text_with_left_header(df_lh:pd.DataFrame, left_header:str):
#     # toxdo: filter out /n and the header
#     header_row_df = df_lh[df_lh['tm_y'] == df_lh[df_lh['text'].str.contains(left_header)].loc[0, 'tm_y']]
#     return header_row_df[~header_row_df['text'].str.contains(f'\n|{left_header}')].loc[0]['text']
# # test_report_date = datetime.datetime.strptime(vdf[get_text_with_left_header()]['text'].str.extract(r'(\d{2}\.\d{
# 2}\.\d{4})').dropna().iloc[0, 0], '%d.%m.%Y')
# # TOxDO: this was getting the test results; not clear if these are needed at this time
# #  get the lot and customer etc. for filing first
# unit_headers = vdf[vdf['text'].str.contains('Unit|Value|Lower Limit|Upper Limit')]

def get_lot_info_dict(lot_info_page: pypdf._page.PageObject):
    """Get a dictionary of lot information from the lot page of an NBE test report

    :param lot_info_page:
    :return:
    """
    lot_keys = ['Purchase Order / date:', 'Delivery / date:', 'Order / date: ', 'Customer number:',
                'Material our / your reference:', 'Commercial Name:', 'Judgement :']
    lot_text_dict = get_left_header_dict_from_page(lot_info_page, lot_keys)

    # use this to split combo lines into multiple dictionary entries; avoids a bunch of case-by-case if/then
    # tuples have a new keys tuple and then the old key: (('new', 'keys'), 'old key')
    split_guide = (('po_number_nbe', 'po_date_nbe'), 'Purchase Order / date:'), \
        (('order_number_nbe', 'order_date_nbe'), 'Order / date: '), \
        (('product_number_nbe', 'tabcode_lw'), 'Material our / your reference:'), \
        (('product_name',), 'Commercial Name:'), \
        (('customer_number_nbe',), 'Customer number:'), \
        (('delivery_number_nbe', 'delivery_date_nbe'), 'Delivery / date:'), \
        (('judgement_nbe',), 'Judgement :')

    # split into a new dictionary with the combo values on their own and (maybe) better names
    lot_info_dict = {}
    for (new_keys, txt_key) in split_guide:
        old_value = lot_text_dict[txt_key]
        new_values = old_value.rsplit('/', 1) if old_value is not None else ['not parsed'] * len(new_keys)
        if old_value is None:
            results_dict['unparsed_count'] += 1
        for (nk, nv) in zip(new_keys, new_values):
            lot_info_dict[nk] = nv.strip()

    return lot_info_dict


def get_left_header_dict_from_page(lot_info_page, lot_keys):
    lot_info_text: str = lot_info_page.extract_text()
    lot_text_dict = {k: None for k in lot_keys}
    split_text = lot_info_text.split('\n')
    while split_text:
        this_line = split_text.pop()
        text_line: str
        for this_key in lot_keys:
            if this_key in this_line:
                lot_text_dict[this_key] = this_line.replace(this_key, '').replace('\n', '').strip()
                break  # stop looking for this key
    return lot_text_dict


from pprint import pprint, pp
from os import scandir, path

for cfd in scandir('../untracked_sample_files/NBE_test_reports'):
    print(f'Processing: {cfd.name}')
    reader = pypdf.PdfReader(path.abspath(cfd.path))
    lot_info_page = reader.pages[0]
    test_results_page = reader.pages[1]
    lot_info: dict = get_lot_info_dict(lot_info_page)
    pp(lot_info)
    print()

pprint(results_dict)
# "Date:" in text, match/group on tm_y, one the of text can be converted to a date dd.mm.yyyy
# same for:
# Purchase Order / date:
# Delivery / date:
# Order / date:
# Customer number:
# Material our / your reference:    THIS HAS THE TABCODE
# Commercial Name:                  THIS HAS THE PRODUCT NAME

# ON THE SECOND PAGE (there can be several of these, each with test results):
# Date of Manufacturing (DOM):
# Prod. Order/Vend.batch:

pass
