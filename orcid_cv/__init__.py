# Public API of orcid_cv package

from orcid_cv.utils import (
    package_directory,
    initialize_name,
    initalize_name,
    embolden_authors,
    is_self_author,
    add_equal_author,
    get_recursive_key,
    dict_to_list,
)

from orcid_cv.config import make_document_config, BACKENDS

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

from orcid_cv.content import (
    prepare_person,
    prepare_affiliations,
    prepare_works,
    prepare_funding,
    prepare_reviews,
    format_review,
    join_authors,
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
    build_document,
    quick_build,
)

# The typst backend is reached through the functions above by passing
# backend="typst" to make_document_config; the module is exposed for direct use.
from orcid_cv import typst_builder
from orcid_cv.typst_builder import assemble_source

# Re-exposing reportlab utilities for backward compatibility
from reportlab.platypus import SimpleDocTemplate
from reportlab.lib.pagesizes import letter

__all__ = [
    "package_directory",
    "initialize_name",
    "initalize_name",
    "embolden_authors",
    "is_self_author",
    "add_equal_author",
    "get_recursive_key",
    "dict_to_list",
    "make_document_config",
    "BACKENDS",
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
    "prepare_person",
    "prepare_affiliations",
    "prepare_works",
    "prepare_funding",
    "prepare_reviews",
    "format_review",
    "join_authors",
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
    "build_document",
    "quick_build",
    "typst_builder",
    "assemble_source",
    "SimpleDocTemplate",
    "letter",
]
