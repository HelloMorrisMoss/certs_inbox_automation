"""Contains the logger setup and a simple script to read the log file into a pandas dataframe."""

import logging
import os
import sys
import types
from logging.handlers import RotatingFileHandler
from typing import Optional, Type

from untracked_config.development_node import ON_DEV_NODE, RUNNING_IN_DEBUG, UNIT_TESTING

# text to prepend to the breadcrumb path detailing running modes (it would be nice if DEBUG could be unambiguous wrt to
# the logging level, but DEBUGGING is too long and anything shorter adds new ambiguity
MODE_PREFIX = 'TEST' * UNIT_TESTING + 'DEBUG' * RUNNING_IN_DEBUG + '.' * any([UNIT_TESTING, RUNNING_IN_DEBUG])
bread_crumb_str = MODE_PREFIX + "{}.{}.{}"


class BreadcrumbFilter(logging.Filter):
    """Provides %(breadcrumbs) field for the logger formatter.

    The breadcrumbs field returns module.funcName.lineno as a single string.
     example:
        formatters={
        'console_format': {'format':
                           '%(asctime)-30s %(breadcrumbs)-35s %(levelname)s: %(message)s'}
                   }
       self.logger.debug('handle_accept() -> %s', client_info[1])
        2020-11-08 14:04:40,561        echo_server03.handle_accept.24      DEBUG: handle_accept() -> ('127.0.0.1',
        49515)
    """

    def filter(self, record):
        record.breadcrumbs = bread_crumb_str.format(record.module, record.funcName, record.lineno)
        return True


def setup_logger(log_file_path: str = './logs/program.log') -> logging.Logger:
    """Set up the logger with console and file handlers.

    Args:
        log_file_path (str): Path to the log file. Defaults to './logs/program.log'.

    Returns:
        logging.Logger: The configured logger object.
    """
    logr = logging.getLogger()
    base_log_level = logging.DEBUG if ON_DEV_NODE else logging.INFO
    logr.setLevel(base_log_level)

    # Console logger
    c_handler = logging.StreamHandler()
    c_handler.setLevel(base_log_level)
    c_format = logging.Formatter('%(asctime)-30s %(breadcrumbs)-45s %(levelname)s: %(message)s')
    c_handler.setFormatter(c_format)
    c_handler.addFilter(BreadcrumbFilter())
    logr.addHandler(c_handler)

    # File logger
    log_dir = os.path.dirname(log_file_path)
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    f_handler = RotatingFileHandler(log_file_path, maxBytes=2000000)
    f_handler.setLevel(base_log_level)
    f_string = '"%(asctime)s","%(name)s", "%(breadcrumbs)s","%(funcName)s","%(lineno)d","%(levelname)s","%(message)s"'
    f_format = logging.Formatter(f_string)
    f_handler.addFilter(BreadcrumbFilter())
    f_handler.setFormatter(f_format)
    logr.addHandler(f_handler)

    def handle_exception(
            exc_type: Type[BaseException],
            exc_value: BaseException,
            exc_traceback: Optional[types.TracebackType]
            ) -> None:
        """Log unhandled exceptions.

        This function is called when an unhandled exception occurs. It logs the exception details using the configured
        logger.

        Args:
            exc_type (Type[BaseException]): The type of the exception.
            exc_value (BaseException): The exception instance.
            exc_traceback (Any): The traceback object.

        Returns:
            None
        """
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        logr.critical("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))

    sys.excepthook = handle_exception

    return logr

    # protect against multiple loggers from importing in multiple files


if __name__ != '__main__':
    lg = setup_logger() if not logging.getLogger().hasHandlers() else logging.getLogger()
