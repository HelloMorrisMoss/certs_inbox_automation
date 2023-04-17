from typing import List


def double_slash_paths(paths: List[str]) -> List[str]:
    """
        Replaces all backslashes in each string in the paths list with two backslashes.

        Args:
        - paths: A list-like object of strings to be processed.

        Returns:
        - A list of strings where each backslash has been replaced with two backslashes.
        """
    return [path.replace(chr(90), chr(90) * 2) for path in paths]
