import pandas as pd

from helpers.outlook_helpers import colorize_outlook_email_list, set_follow_up_on_list


def set_priority_customer_category(df: pd.DataFrame, priority_flag_dict: dict, follow_up=True, color_category: str = '') -> None:
    """Sets the color category of mail items from priority customers in the given DataFrame to the specified color.

    :param df: The DataFrame containing the mail items to filter and colorize.
    :param priority_flag_dict: A dictionary containing the customer names to flag as highest priority.
    :param follow_up: bool, whether to mark an e-mail with a follow-up flag
    :param color_category: The name of the color category to apply to the mail items (default is 'red').
    """
    # filter on priority customers
    flag_df = df.loc[df.customer.str.match('|'.join(priority_flag_dict['highest'].keys()))]
    if follow_up:
        set_follow_up_on_list(flag_df['o_item'])

    if color_category:
        # set priority customer e-mails to color category
        colorize_outlook_email_list(flag_df['o_item'], color_category)


