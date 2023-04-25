import pandas as pd

from helpers.outlook_helpers import colorize_outlook_email_list, clear_of_all_category_colors_from_list


def set_priority_customer_category(df: pd.DataFrame, priority_flag_dict: dict, color_category: str = 'red') -> None:
    """Sets the color category of mail items from priority customers in the given DataFrame to the specified color.

    :param df: The DataFrame containing the mail items to filter and colorize.
    :param priority_flag_dict: A dictionary containing the customer names to flag as highest priority.
    :param color_category: The name of the color category to apply to the mail items (default is 'red').
    """
    # filter on priority customers
    flag_df = df.loc[df.customer.str.match('|'.join(priority_flag_dict['highest'].keys()))]
    # set priority customer e-mails to color category
    colorize_outlook_email_list(flag_df['o_item'], color_category)
    pass  # for development breakpoint
    clear_of_all_category_colors_from_list(flag_df['o_item'])
