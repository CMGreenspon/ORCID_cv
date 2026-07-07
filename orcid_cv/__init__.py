# Public API of orcid_cv package

from orcid_cv.utils import (
    package_directory,
    initialize_name,
    initalize_name,
    embolden_authors,
    add_equal_author,
    get_recursive_key,
    dict_to_list,
)

from orcid_cv.config import make_document_config

from orcid_cv.parser import (
    load_xml,
    list_works,
    load_affiliation,
    load_work,
    check_duplicates,
    prune_duplicate_works,
    find_preprint_repository,
    load_funding,
    load_review,
    extract_orcid_info,
    folder_to_dict,
)

from orcid_cv.builder import (
    HyperlinkedImage,
    FooterCanvas,
    get_column_widths,
    make_affiliation_table,
    make_work_table,
    make_funding_table,
    make_review_table,
    process_external_links,
    add_person_section,
    add_affiliation_section,
    add_work_section,
    add_funding_section,
    add_review_section,
    quick_build,
)

# Re-exposing reportlab utilities for backward compatibility
from reportlab.platypus import SimpleDocTemplate
from reportlab.lib.pagesizes import letter

__all__ = [
    "package_directory",
    "initialize_name",
    "initalize_name",
    "embolden_authors",
    "add_equal_author",
    "get_recursive_key",
    "dict_to_list",
    "make_document_config",
    "load_xml",
    "list_works",
    "load_affiliation",
    "load_work",
    "check_duplicates",
    "prune_duplicate_works",
    "find_preprint_repository",
    "load_funding",
    "load_review",
    "extract_orcid_info",
    "folder_to_dict",
    "HyperlinkedImage",
    "FooterCanvas",
    "get_column_widths",
    "make_affiliation_table",
    "make_work_table",
    "make_funding_table",
    "make_review_table",
    "process_external_links",
    "add_person_section",
    "add_affiliation_section",
    "add_work_section",
    "add_funding_section",
    "add_review_section",
    "quick_build",
    "SimpleDocTemplate",
    "letter",
]
