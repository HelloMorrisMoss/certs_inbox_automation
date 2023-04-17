import pandas as pd


def df_json_handler(df):
    try:
        return df.to_dict()
    except AttributeError as aer:
        if isinstance(df, pd.Timestamp) or isinstance(df, pd.Timedelta):
            return df.isoformat()
        else:
            return repr(df)
