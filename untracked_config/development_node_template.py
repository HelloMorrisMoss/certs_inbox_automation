"""Contains a boolean 'constant' that is true on the development system and false otherwise."""
import os
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
DEV_TEST_MODE: Final[bool] = os.environ.get('UNITTEST') or (sys.gettrace() is not None)
