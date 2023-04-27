from typing import Dict, List, Optional, Union, Tuple, Final

import pandas as pd
from win32com import client as wclient

color_map: Final[dict] = {'red': 'Red Category',
             'orange': 'Orange Category',
             'yellow': 'Yellow Category',
             'olive': 'olive',
             'green': 'Green Category',
             'blue': 'Blue Category',
             'purple': 'Purple Category',
             'pink': 'Pink Category',
             'grey': 'Grey Category',
             }
valid_colors = color_map.keys()
valid_categories = color_map.values()

default_follow_up_text = 'Follow up'


def add_categories_to_mail(mail: wclient.CDispatch, categories: Union[str, List[str]]) -> None:
    """Add categories to an Outlook mail item.

    :param mail: Outlook mail item to add categories to.
    :param categories: A string or list of strings specifying the categories to add.
    :raises ValueError: If the input is not a string or a list of strings.
    """

    action_text = 'added'

    existing_categories, normalized_categories = get_actionable_categories(action_text, categories, mail)

    # Add the new categories that are not already in the existing categories
    for category in normalized_categories:
        if category not in existing_categories:
            existing_categories.append(category)

    # Set the new categories to the mail item
    mail.Categories = ", ".join(existing_categories)
    mail.Save()


def remove_categories_from_mail(mail: wclient.CDispatch, categories: Union[str, List[str]]) -> None:
    """Remove categories from an Outlook mail item.

    :param mail: Outlook mail item to remove categories from.
    :param categories: A string or list of strings specifying the categories to remove.
    :raises ValueError: If the input is not a string or a list of strings.
    """

    action_text = 'removed'

    existing_categories, normalized_categories = get_actionable_categories(action_text, categories, mail)

    # Remove the categories that match the categories to remove
    for category in normalized_categories:
        if category in existing_categories:
            existing_categories.remove(category)

    # Set the new categories to the mail item
    mail.Categories = ", ".join(existing_categories)
    mail.Save()


def get_actionable_categories(action_text: str, categories: Union[str, List[str]], mail: wclient.CDispatch) -> Tuple[List[str], List[str]]:
    """Normalize and validate color categories for use in an Outlook mail item.

    :param action_text: A string indicating the action being taken (e.g., "added" or "removed").
    :param categories: A string or list of strings specifying the categories to normalize and validate.
    :param mail: The Outlook mail item to validate the categories for.
    :raises ValueError: If the input is not a string or a list of strings.
    :raises ValueError: If a category is not a valid color category.
    :returns: A tuple containing two lists: the existing categories of the mail item, and the normalized categories to use.
    """

    if isinstance(categories, str):
        categories = [categories]

    # Raise an error if the input is not a list of strings
    if not isinstance(categories, list) or not all(isinstance(x, str) for x in categories):
        raise ValueError("Categories must be a string or a list of strings")

    # normalize the categories and ensure they are valid color categories
    normalized_categories = normalize_color_categories_list(categories)

    # Get the existing categories of the mail item
    existing_categories = mail.Categories.split(", ") if mail.Categories else []

    return existing_categories, normalized_categories


def normalize_color_categories_list(categories: list):
    """Normalize the categories and ensure they are valid color categories

    example:
    nccl = normalize_color_categories_list(['grey', 'Blue', 'Red Category'])
    print(nccl)
    >[['Grey Category', 'Blue Category', 'Red Category']]

    :param categories: list, containing strings of either categories or the color name.
    :returns: list, the list of strings with valid color categories.
    """
    normalized_categories = []
    for color in categories:
        color = color.lower()
        if color in valid_categories:
            normalized_categories.append(color)
            continue
        elif color not in valid_colors:
            raise ValueError(f"{color} is not a valid color.")
        else:
            color_cat = color_map[color]
            normalized_categories.append(color_cat)
    return normalized_categories


def get_store_by_name(store_name_filter: str, outlook_obj: wclient.CDispatch) -> Optional[object]:
    """Searches for an Outlook store with a display name that contains the given filter string and returns the first
    store that matches. If no matching store is found, returns None.

    example outlook object:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    :param outlook_obj: win32com.CDispatch, Outlook object.
    :param store_name_filter: A string to search for in the display names of the Outlook stores.
    :return: The first Outlook store that matches the search filter, or None if no match is found.
    """
    for olStore in outlook_obj.Stores:
        if store_name_filter not in olStore.DisplayName:
            continue
        target_store = olStore
        return target_store
    return None


def map_folder_structure_to_flat_dict(folders_dict: Dict[str, any], parent_folder: wclient.CDispatch,
                                      must_find_list: List[str]) -> None:
    """Iteratively searches for all folders within the specified parent_folder object and updates the
    folders_dict dictionary with the folder paths and olFolder objects. Stops searching as soon as all
    folders in must_find_list have been found.

    :param folders_dict: A dictionary to store the folder paths and olFolder objects.
    :type folders_dict: Dict[str, any]
    :param parent_folder: The parent folder object to search within.
    :type parent_folder: any
    :param must_find_list: A list of folder paths that must be found. If not all are found, continue searching.
    :type must_find_list: List[str]
    :return: None
    :rtype: None
    """
    folders_stack = [parent_folder]
    while folders_stack:
        current_folder = folders_stack.pop()
        for olFolder in current_folder.Folders:
            folder_path = olFolder.FolderPath
            folders_dict[folder_path] = olFolder
            if must_find_list:
                if all(mfitem in folders_dict.keys() for mfitem in must_find_list):
                    return
            folders_stack.append(olFolder)


def find_folders_in_outlook(outlook_obj: wclient.CDispatch, store_name_filter: str, must_find_list: List[str] = '',
                            map_all=False) -> Dict[str, any]:
    """Get outlook folders in a dictionary from an Outlook object for a specified account.

    Searches for all folders within Outlook stores whose display names contain the specified
    store_name_filter string, and returns a dictionary where the keys are the folder paths and the values
    are the corresponding olFolder objects. Raises a custom exception if any of the folders in must_find_list
    are not found.
    """
    must_find_list = must_find_list if (must_find_list and not map_all) else ''
    folders_dict = {}
    target_store = get_store_by_name(store_name_filter, outlook_obj)
    parent_folder = target_store.GetRootFolder()
    map_folder_structure_to_flat_dict(folders_dict, parent_folder, must_find_list)

    for folder in must_find_list:
        if folder not in folders_dict.keys():
            raise Exception(f"Required folder '{folder}' not found!")
    return folders_dict


def colorize_outlook_email_list(mail_items: list, color: str):
    """Add the color category to all the mail items in the list.

    :param mail_items: list, list of w32com.CDispatch.client Outlook mail items.
    :param color: str, the color categories to set on the mail items.
    """
    for mail_item in mail_items:
        add_categories_to_mail(mail_item, color)


def clear_all_category_colors(o_item: wclient.CDispatch) -> None:
    """Removes all categories from the mail items in the given DataFrameGroupBy object.

    :param o_item: The mail item to remove the categories from.
    """
    o_item.Categories = ''
    o_item.Save()


def clear_all_category_colors_foam(dfg: List[Tuple[str, pd.DataFrame]]) -> None:
    """Removes all categories from the mail items in the given DataFrameGroupBy object.

    :param dfg: The DataFrameGroupBy object containing the mail items to remove the categories from.
    """
    for group_name, group_df in dfg:
        for _, row in group_df.iterrows():
            o_item = row['o_item']
            clear_all_category_colors(o_item)


def clear_of_all_category_colors_from_list(o_items: List[wclient.CDispatch]) -> None:
    """Removes all categories from the given list of mail items.

    :param o_items: A list of mail items to remove the categories from.
    """
    for item in o_items:
        clear_all_category_colors(item)


def move_mail_items_to_folder(mail_items_list: List[wclient.CDispatch], destination_folder: wclient.Dispatch):
    """Moves the given list of mail items to the specified destination folder.

    :param mail_items_list: The list of mail items to be moved.
    :param destination_folder: The CDispatch object of the destination folder.
    """
    for mail in mail_items_list:
        mail.move(destination_folder)


def set_follow_up_on_list(item_list: List[wclient.CDispatch], follow_up_text: str = default_follow_up_text,
                          overwrite_if_set: bool = False) -> None:
    """Sets a follow-up flag on the given list of mail items. By default, will not overwrite any existing follow-up.

    :param item_list: The list of mail items to set the follow-up flag on.
    :param follow_up_text: The text for the follow-up flag. If not provided, it will use the default text.
    :param overwrite_if_set: bool, whether to overwrite existing follow-up status that may be set. Default False.
    """
    change_setting = False
    for item in item_list:
        if overwrite_if_set:  # always overwrite don't check
            change_setting = True
        elif not is_follow_up_set(item):  # check for existing follow-up
            change_setting = True

        if change_setting:
            set_follow_up(item, follow_up_text)


def set_follow_up(mail_item: wclient.CDispatch, follow_up_text: str = default_follow_up_text):
    """Sets a follow-up flag on the given mail item.

    :param mail_item: The mail item to set the follow-up flag on.
    :param follow_up_text: The text for the follow-up flag. If not provided, it will use the default text.
    """
    mail_item.FlagRequest = follow_up_text
    mail_item.save()


def reset_testing_mods(mail_list: List[wclient.CDispatch]):
    """Resets any testing modifications made to the given list of mail items.

    This function clears all color categories and removes any follow-up flags.

    :param mail_list: The list of mail items to reset the modifications on.
    """
    clear_of_all_category_colors_from_list(mail_list)
    set_follow_up_on_list(mail_list, '')


def is_follow_up_set(outlook_mail_item: wclient.CDispatch) -> bool:
    """Returns True if the given Outlook mail item has a follow-up flag set; otherwise, returns False.

    :param outlook_mail_item: The Outlook mail item to check for a follow-up flag.
    :type outlook_mail_item: Dispatch
    :return: True if the mail item has a follow-up flag set, False otherwise.
    :rtype: bool
    """
    if any([outlook_mail_item.FlagRequest, outlook_mail_item.FlagStatus]):
        return True
    else:
        return False
