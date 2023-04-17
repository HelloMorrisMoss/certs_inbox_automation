from typing import Dict, List, Optional

from win32com import client as wclient


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


def find_folders(outlook_obj: wclient.CDispatch, store_name_filter: str, must_find_list: List[str] = '', map_all=False) -> Dict[str, any]:
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
