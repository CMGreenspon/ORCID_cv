import os
import logging
from typing import Any, Dict, List, Union

logger = logging.getLogger("orcid_cv")

# package_directory resolves to the project root (where external_link_img resides)
package_directory = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def initialize_name(input_str: str) -> str:
    """
    Initializes first and middle names, keeping only initials followed by a dot,
    while leaving the last name intact. E.g., 'Charles M. Greenspon' -> 'C.M. Greenspon'.
    """
    if not input_str or not input_str.strip():
        return ""

    str_split = [part for part in input_str.split(" ") if part]
    if len(str_split) <= 1:
        return input_str

    for i in range(len(str_split) - 1):
        part = str_split[i]
        if "." in part:
            continue
        str_split[i] = part[0] + "."

    first_parts = "".join(str_split[:-1])
    return f"{first_parts} {str_split[-1]}"


def initalize_name(input_str: str) -> str:
    """Backward compatibility alias for initialize_name."""
    logger.debug("initalize_name is deprecated, use initialize_name instead.")
    return initialize_name(input_str)


def is_self_author(person: Dict[str, Any], author: str) -> bool:
    """
    Returns True if an author string refers to the CV's owner. Markup-agnostic so
    that every rendering backend can decide how to highlight the name.
    """
    lastname = person.get("lastname", "")
    if not lastname or lastname not in author:
        return False

    fullname = person.get("fullname", "")
    name_short = person.get("name-short", "")
    firstname = person.get("firstname", "")

    if fullname in author:
        return True
    if name_short in author:
        return True
    if f"{firstname} {lastname}" in author:
        return True
    if f"{firstname[0] if firstname else ''}. {lastname}" in author:
        return True

    logger.info(f"Did not embolden: {author}")
    return False


def embolden_authors(person: Dict[str, Any], author_list: List[str]) -> List[str]:
    """
    Emboldens the target person's name in a list of author names by wrapping
    it in HTML <b> tags.
    """
    for i, author in enumerate(author_list):
        if is_self_author(person, author):
            author_list[i] = f"<b>{author}</b>"

    return author_list


def add_equal_author(
    author_list: List[str], num_first: int = 0, num_last: int = 0
) -> None:
    """
    Appends an asterisk (*) to the names of the first `num_first` and last
    `num_last` authors to indicate equal contribution.
    """
    num_authors = len(author_list)
    for i in range(1, num_authors + 1):
        if i <= num_first:
            author_list[i - 1] += "*"
        if i > num_authors - num_last:
            author_list[i - 1] += "*"


def get_recursive_key(input_dict: Dict[str, Any], *keys: str) -> Any:
    """
    Safely retrieves a value nested deep inside a dictionary.
    Returns an empty string if any intermediate key does not exist or has a value of None.
    """
    if not isinstance(input_dict, dict):
        raise TypeError("first argument must be a dict.")

    if len(keys) == 0 or not all(isinstance(k, str) for k in keys):
        raise TypeError("Keys must all be of type 'str'.")

    _dict = input_dict
    for key in keys:
        if isinstance(_dict, dict) and key in _dict and _dict[key] is not None:
            _dict = _dict[key]
        else:
            put_code = input_dict.get("@put-code", "unknown")
            logger.info(f"Could not find: {'-'.join(keys)} in item #{put_code}")
            return ""

    return _dict


def dict_to_list(input_dict: Dict[str, Any]) -> List[Any]:
    """Converts a dictionary's values to a list."""
    return list(input_dict.values())
