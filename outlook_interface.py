"""A win32com interface for dealing with Outlook.

This module provides a win32com interface for interacting with Microsoft Outlook. It includes functionality for handling
the Outlook application, accessing the MAPI namespace, resetting the Outlook instance, terminating the Outlook process,
and starting Outlook.

Classes:
    OutlookSingleton: Provides a single instance of the Outlook application and handles Outlook being unavailable for
        several issues.

Functions:
    get_outlook_installation_path: Retrieve the installation path of Microsoft Outlook from the registry.
    start_outlook: Start Microsoft Outlook using the specified application path.

Variables:
    wc_outlook: An instance of the `OutlookSingleton` class representing the single instance of the Outlook application.

"""

import subprocess
import time
import winreg
from typing import Union

import pythoncom
import win32com.client

from log_setup import lg


class OutlookSingleton:
    """Provides a single instance of the Outlook application and handles Outlook being unavailable for several issues.
    """
    _instance = None

    def __new__(cls) -> 'OutlookSingleton':
        if cls._instance is None:
            lg.debug('Initializing Outlook instance.')
            cls._instance = super().__new__(cls)
            cls._instance._outlook = None
        return cls._instance

    def _get_outlook(self) -> win32com.client.Dispatch:
        """Gets a working instance of Outlook.

        Returns the instance of the Outlook application and checks if the Outlook session is still valid.
        If the session has expired, the method reopens the application using win32com.client.Dispatch. CoInitilize will
            be reset.

        :return: The win32com Dispatch object representing the Outlook application instance.
        """
        if self._outlook is None:
            pythoncom.CoInitialize()
            self._outlook = win32com.client.Dispatch("Outlook.Application")
        try:
            self._outlook.Session
        except win32com.client.pywintypes.com_error as py_win_err:
            if py_win_err.hresult == -2147023174:
                # Outlook session has expired, reopen Outlook
                lg.warning("Outlook session expired, reopening")
                self._reset_coinitialize()
                self._outlook = win32com.client.Dispatch("Outlook.Application")
            elif py_win_err.hresult == -2147220995:
                lg.debug('Not connected to the server, may still be loading.')
                time.sleep(15)
                try:
                    self._outlook.Session  # try one more time
                except Exception:
                    lg.info('Could not connect to Outlook server, restarting Outlook application.')
                    self.reset_outlook()
            else:
                # Other errors, log and raise the exception
                lg.exception("Error accessing Outlook", exc_info=py_win_err)
                raise
        return self._outlook

    def get_outlook(self) -> win32com.client.Dispatch:
        """Returns the instance of the Outlook application.

         Checks if the Outlook session is still valid. If not, the method reopens the application using
          win32com.client.Dispatch. This will reset CoInitialized resources.

        :return: The win32com Dispatch object representing the Outlook application instance.
        """
        return self._get_outlook()

    def get_outlook_folders(self) -> win32com.client.Dispatch:
        """Get the MAPI namespace of the Outlook application.

        The MAPI namespace is a hierarchy of folders that represent different Outlook data stores, such as email
        accounts, calendars, contacts, and tasks. The MAPI namespace is represented by a win32com Dispatch object that
        can be used to access and manipulate the Outlook data stores.

        :return: The win32com Dispatch object representing the MAPI namespace of the Outlook application.
        """
        outlook = self.get_outlook()
        mapi_namespace = outlook.GetNamespace("MAPI")
        return mapi_namespace

    @staticmethod
    def _reset_coinitialize() -> None:
        """Reset the COM library resources.

        Resets the COM library by uninitializing and then reinitializing it using the `CoUninitialize()` and
        `CoInitialize()` functions from the `pythoncom` module.

        This method is useful for handling errors related to the win32com Outlook connection, such as session expiration
        or Outlook not running, where the COM library needs to be reset before creating a new instance of the Outlook
        application using `win32com.client.Dispatch()`.

        :return: None
        """
        pythoncom.CoUninitialize()
        pythoncom.CoInitialize()

    def __del__(self):
        """Uninitializes the COM library when the object is destroyed."""
        pythoncom.CoUninitialize()

    def terminate_outlook(self) -> None:
        """Terminate the existing instance of the Outlook application.

        This method closes the existing Outlook process and waits for 5 seconds before returning. This allows time for
        the process to fully shut down before starting a new process.

        Returns:
            None
        """
        # Close the existing Outlook process
        lg.info('Terminating existing Outlook Application.')
        self._outlook.Quit()
        pythoncom.CoUninitialize()
        self._outlook = None

        # Wait for 5 seconds before starting a new process
        time.sleep(5)

    def reset_outlook(self) -> win32com.client.Dispatch:
        """Reset the instance of the Outlook application.

        This method closes the existing instance of the Outlook application and returns a new instance.

        Returns:
            The Dispatch object representing the new instance of the Outlook application.
        """
        lg.info('Restarting Outlook Application, this will take at least 15 seconds.')
        self.terminate_outlook()
        start_outlook()
        time.sleep(10)

        try:
            return self.get_outlook()
        except Exception as e:
            lg.error(f"Error resetting Outlook application: {e}")
            raise e


def get_outlook_installation_path() -> Union[str, None]:
    """Retrieve the installation path of Microsoft Outlook from the registry.

    Returns:
        Union[str, None]: The installation path of Microsoft Outlook or None if it is not found.
    """
    try:
        reg_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
            value, _ = winreg.QueryValueEx(key, None)
            return value
    except FileNotFoundError:
        return None


def start_outlook(application_path: Union[str, None] = None) -> None:
    """Start Microsoft Outlook using the specified application path.

    Args:
        application_path (Union[str, None], optional): The path of the Outlook application executable. If None, the default installation path will be used. Defaults to None.

    Returns:
        None
    """
    application_path = get_outlook_installation_path() if application_path is None else application_path
    subprocess.Popen(application_path)



wc_outlook = OutlookSingleton()
