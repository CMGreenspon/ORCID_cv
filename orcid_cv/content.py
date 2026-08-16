"""
Backend-neutral content preparation.

Everything in this module turns the raw ORCID dictionary into plain Python data
structures (no markup, no layout) so that the reportlab and typst backends can
render the exact same CV content in their own idioms.
"""

import logging
from typing import Any, Dict, List, Optional, Tuple, Union

from orcid_cv.utils import dict_to_list, initialize_name, is_self_author

logger = logging.getLogger("orcid_cv")

# Author entries are (name, is_the_cv_owner) pairs.
Author = Tuple[str, bool]

# Characters that ORCID hands back which look wrong in a CV.
_CHAR_FIXES = {
    "‐": "-",  # Unicode hyphen -> ascii hyphen
}
_AUTHOR_CHAR_FIXES = {
    "‐": "-",
    "ř": "r",
}


def _apply_fixes(text: str, fixes: Dict[str, str]) -> str:
    """Applies a table of character substitutions to a string."""
    for old, new in fixes.items():
        if old in text:
            text = text.replace(old, new)
    return text


def prepare_person(orcid_dict: Dict[str, Any]) -> Dict[str, Any]:
    """
    Returns the header block content: full name, current role/organization,
    email and the external link dictionary.
    """
    personal = orcid_dict.get("personal", {})
    info = prepare_affiliations(orcid_dict, "employment")

    current = info[0] if info else {}
    return {
        "fullname": personal.get("fullname", ""),
        "role": current.get("role", ""),
        "organization": current.get("organization", ""),
        "email": personal.get("email", ""),
        "links": personal.get("links", {}),
    }


def prepare_affiliations(
    orcid_dict: Dict[str, Any], affiliation_type: str
) -> List[Dict[str, Any]]:
    """Returns employments or educations sorted by start date, most recent first."""
    if affiliation_type not in orcid_dict:
        logger.warning(f"Dict does not contain affiliation type: {affiliation_type}")
        return []

    affiliations = dict_to_list(orcid_dict[affiliation_type])
    try:
        affiliations = sorted(
            affiliations, key=lambda v: int(v.get("start_date", 0)), reverse=True
        )
    except (ValueError, TypeError):
        pass
    return affiliations


def prepare_funding(orcid_dict: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Returns funding entries sorted by start year, most recent first."""
    if "funding" not in orcid_dict:
        logger.warning("Dict does not contain funding")
        return []

    fund = dict_to_list(orcid_dict["funding"])
    try:
        fund = sorted(fund, key=lambda v: int(v.get("start_year", 0)), reverse=True)
    except (ValueError, TypeError):
        pass
    return fund


def prepare_reviews(orcid_dict: Dict[str, Any]) -> List[Tuple[str, int]]:
    """
    Counts peer reviews per journal and returns (journal, count) pairs sorted
    alphabetically by journal name.
    """
    if "reviews" not in orcid_dict:
        return []

    review_dict: Dict[str, int] = {}
    for v in orcid_dict["reviews"].values():
        org_name = v.get("org", "")
        if org_name:
            review_dict[org_name] = review_dict.get(org_name, 0) + 1

    return [(org, review_dict[org]) for org in sorted(review_dict.keys())]


def format_review(org: str, count: int) -> str:
    """Renders a single peer review tally, e.g. 'eLife, 3 reviews'."""
    if not org:
        return ""
    return f"{org}, {count} reviews" if count > 1 else f"{org}, 1 review"


def _prepare_link(doi_str: str) -> Optional[Dict[str, str]]:
    """
    Splits a work's URL into a display prefix, a short label and the target URL.
    Returns None when the work has no link.
    """
    if not doi_str:
        return None

    if "doi.org/" in doi_str:
        short = doi_str[doi_str.find("doi.org/") + 8 :]
        return {
            "prefix": "DOI: ",
            "label": short,
            "url": f"https://www.doi.org/{short}",
        }
    if "github.com/" in doi_str:
        short = doi_str[doi_str.find("github.com/") + 11 :]
        return {
            "prefix": "GitHub: ",
            "label": short,
            "url": f"https://www.github.com/{short}",
        }
    return {"prefix": "", "label": doi_str, "url": doi_str}


def prepare_works(
    orcid_dict: Dict[str, Any],
    config: Dict[str, Any],
    search_str: Union[str, List[str]],
) -> List[Dict[str, Any]]:
    """
    Returns the works matching `search_str`, sorted newest first, with authors,
    journal and link information resolved into markup-free fields.
    """
    if isinstance(search_str, str):
        search_str = [search_str]

    works = dict_to_list(orcid_dict["work"])
    works = [w for w in works if w.get("type") in search_str]
    if not works:
        logger.warning(f"No matching works for: {search_str}")
        return []

    try:
        works = sorted(
            works,
            key=lambda v: int(v.get("year", 0)) * 1000 + int(v.get("month", 0)),
            reverse=True,
        )
    except (ValueError, TypeError):
        pass

    personal = orcid_dict.get("personal", {})
    prepared = []
    for work in works:
        work_date = str(work.get("year", ""))
        work_journal = work.get("journal", "")
        work_title = _apply_fixes(work.get("title", ""), _CHAR_FIXES)
        subtitle = work.get("subtitle", "")

        # Software entries store the repository in the subtitle and the year in journal
        if work.get("type") == "software":
            work_journal = subtitle
            work_date = str(work.get("journal", ""))
            subtitle = ""

        author_list = list(work.get("authors", []))
        if config.get("initalize_authors"):
            author_list = [initialize_name(i) for i in author_list]
        embolden = bool(config.get("embolden_author"))
        authors: List[Author] = [
            (
                _apply_fixes(a, _AUTHOR_CHAR_FIXES),
                embolden and is_self_author(personal, a),
            )
            for a in author_list
        ]

        prepared.append(
            {
                "title": work_title,
                "date": work_date,
                "journal": work_journal,
                "subtitle": subtitle,
                "link": _prepare_link(work.get("doi", "")),
                "authors": authors,
            }
        )

    return prepared


def join_authors(authors: List[Author], bold: Any = None, plain: Any = None) -> str:
    """
    Joins authors into a citation string ('A', 'A and B', 'A, B, and C'). The
    owner's name is passed through `bold` and every other name through `plain`,
    letting each backend apply its own markup and escaping.
    """
    if not authors:
        return ""

    def render(name: str, is_owner: bool) -> str:
        fun = bold if (is_owner and bold) else plain
        return fun(name) if fun else name

    names = [render(name, flag) for name, flag in authors]
    if len(names) == 1:
        return names[0]
    if len(names) == 2:
        return f"{names[0]} and {names[1]}"

    names[-1] = "and " + names[-1]
    return ", ".join(names)
