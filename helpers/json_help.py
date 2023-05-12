"""A module for a custom JSON handler for pandas DataFrame serialization.

This module provides a custom JSON handler function, `df_json_handler`, specifically designed for serializing pandas
DataFrame objects. The `df_json_handler` function can be used as a default function for JSON serialization of pandas
objects. It handles the serialization of pandas DataFrame, Series, Timestamp, and Timedelta objects. If the object is a
DataFrame, it is converted to a dictionary representation using the `to_dict()` method. If the object is a Timestamp or
Timedelta, it is converted to its ISO format. For other objects, their representation is returned.

Functions:
    df_json_handler: Custom JSON handler for DataFrame serialization.

"""

from typing import Any, Union

import pandas as pd


def df_json_handler(df: Union[pd.DataFrame, pd.Series, pd.Timestamp, pd.Timedelta, Any]) -> Any:
    """Custom JSON handler for DataFrame serialization.

    This function is used as a default function for JSON serialization of pandas objects. It handles the serialization of
    pandas DataFrame, Series, Timestamp, and Timedelta objects. If the object is a DataFrame, it is converted to a
    dictionary representation using the `to_dict()` method. If the object is a Timestamp or Timedelta, it is converted
    to its ISO format. For other objects, their representation is returned.

    Args:
        df (Union[pd.DataFrame, pd.Series, pd.Timestamp, pd.Timedelta, Any]): The pandas object to be serialized.

    Returns:
        Any: The serialized representation of the pandas object.
    """
    try:
        return df.to_dict()
    except AttributeError as aer:
        if isinstance(df, pd.Timestamp) or isinstance(df, pd.Timedelta):
            return df.isoformat()
        else:
            return repr(df)
