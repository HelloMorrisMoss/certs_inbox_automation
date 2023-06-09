"""Contains a boolean 'constant' that is true on the development system and false otherwise."""
import platform
import sys
from typing import Final, List

__dev_node_names: List[str] = [
    'development_host_name',
]
# __dev_node_name = ''  # for testing non-dev system behavior

# this section is to remove the old database table if the DefectModel table needs to be changed:
node = platform.node()
ON_DEV_NODE: Final[bool] = node in __dev_node_names

# whether the program is running as a unit test or in debug mode; Outlook folder paths are switched to testing folders
UNIT_TESTING = any(['unittest' in arg for arg in sys.argv])
RUNNING_IN_DEBUG = (sys.gettrace() is not None)
DEV_TEST_MODE: Final[bool] = UNIT_TESTING or RUNNING_IN_DEBUG
