"""A win32com interface for dealing with Outlook"""

import pythoncom
import win32com.client

from log_setup import lg


class OutlookSingleton:
    """Provides a single instance of the Outlook application and handles Outlook being unavailable for several issues.
    """
    _instance = None

    def __new__(cls) -> 'OutlookSingleton':
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._outlook = None
        return cls._instance

    def _get_outlook(self) -> win32com.client.CDispatch:
        """Gets a working instance of Outlook.

        Returns the instance of the Outlook application and checks if the Outlook session is still valid.
        If the session has expired, the method reopens the application using win32com.client.Dispatch. CoInitilize will
            be reset.

        :return: The win32com CDispatch object representing the Outlook application instance.
        """
        if self._outlook is None:
            pythoncom.CoInitialize()
            self._outlook = win32com.client.Dispatch("Outlook.Application")
        try:
            self._outlook.Session
        except Exception as e:
            if isinstance(e, win32com.client.pywintypes.com_error) and e.hresult == -2147023174:
                # Outlook session has expired, reopen Outlook
                lg.warning("Outlook session expired, reopening")
                self._reset_coinitialize()
                self._outlook = win32com.client.Dispatch("Outlook.Application")
            elif isinstance(e, win32com.client.pywintypes.com_error) and e.hresult == -2147352567:
                # Outlook is not running, try opening it
                lg.warning("Outlook is not running, opening Outlook")
                self._reset_coinitialize()
                self._outlook = win32com.client.Dispatch("Outlook.Application")
            else:
                # Other errors, log and raise the exception
                lg.exception("Error accessing Outlook")
                raise
        return self._outlook

    def get_outlook(self) -> win32com.client.CDispatch:
        """Returns the instance of the Outlook application.

         Checks if the Outlook session is still valid. If not, the method reopens the application using
          win32com.client.Dispatch. This will reset CoInitialized resources.

        :return: The win32com CDispatch object representing the Outlook application instance.
        """
        return self._get_outlook()

    def get_outlook_folders(self) -> win32com.client.CDispatch:
        """Get the MAPI namespace of the Outlook application.

        The MAPI namespace is a hierarchy of folders that represent different Outlook data stores, such as email
        accounts, calendars, contacts, and tasks. The MAPI namespace is represented by a win32com CDispatch object that
        can be used to access and manipulate the Outlook data stores.

        :return: The win32com CDispatch object representing the MAPI namespace of the Outlook application.
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
        """
        Uninitializes the COM library when the object is destroyed.
        """
        pythoncom.CoUninitialize()


wc_outlook = OutlookSingleton()