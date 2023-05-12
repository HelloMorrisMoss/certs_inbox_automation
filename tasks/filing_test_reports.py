import pandas as pd
import pypdf

# pandas display settings for development
pd.set_option('display.max_rows', 100)
pd.set_option('display.max_columns', 100)
pd.set_option('display.width', 1000)

reader = pypdf.PdfReader("../untracked_sample_files/NBE_test_reports/Certificate for Delivery0080615826000010.PDF")

# header_df = pd.DataFrame.from_dict({'Unit': [], 'Value': [], 'Lower Limit': [], 'Upper Limit': []})


def page_text_to_dataframe(rdr_page):
    vd = {}
    vd['visitor_df'] = pd.DataFrame.from_dict({prm: [] for prm in ['text', 'cm', 'tm', 'font_dict', 'font_size']})
    def visitor_body(text, cm, tm, font_dict, font_size):
        # global visitor_df
        vd['visitor_df'] = pd.concat([vd['visitor_df'],
                                pd.DataFrame.from_dict({'text': [text],
                                                        'cm': [cm],
                                                        'tm': [tm],
                                                        'font_dict': [font_dict],
                                                        'font_size': font_size,
                                                        })])

    rdr_page.extract_text(visitor_text=visitor_body)
    vd['visitor_df'][['tm_0', 'tm_1', 'tm_2', 'tm_3', 'tm_x', 'tm_y']] = vd['visitor_df']['tm'].apply(pd.Series)
    vd['visitor_df'][['cm_0', 'cm_1', 'cm_2', 'cm_3', 'cm_x', 'cm_y']] = vd['visitor_df']['cm'].apply(pd.Series)
    vd['visitor_df'][['base_font', 'encoding', 'subtype', 'type']] = vd['visitor_df']['font_dict'].apply(series_default_obj)
    return vd['visitor_df']


def series_default_obj(input_dict):
    return pd.Series(input_dict, dtype='object')


lot_info_page = reader.pages[0]
test_results_page = reader.pages[1]

vdf = page_text_to_dataframe(lot_info_page)

# TODO: this was getting the test results; not clear if these are needed at this time
#  get the lot and customer etc. for filing first
unit_headers = vdf[vdf['text'].str.contains('Unit|Value|Lower Limit|Upper Limit')]

# "Date:" in text, match/group on tm_y, one the of text can be converted to a date mm.dd.yyyy
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