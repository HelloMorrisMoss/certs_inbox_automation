from typing import Dict, List, Optional, Union, Tuple, Final

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
            print(f'{color} being interpreted as {color_cat}')
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