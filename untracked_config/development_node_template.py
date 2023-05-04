"""Contains a boolean 'constant' that is true on the development system and false otherwise."""

import platform
from typing import Final, List

__dev_node_names: List[str] = [
    'development_host_name',
]
# __dev_node_name = ''  # for testing non-dev system behavior

# this section is to remove the old database table if the DefectModel table needs to be changed:
node = platform.node()
ON_DEV_NODE: Final[bool] = node in __dev_node_names
