"""
Typst CV builder.

This is the typst counterpart of `orcid_cv.builder`. It exposes the same
`add_*_section` / `quick_build` surface, but instead of accumulating reportlab
flowables it accumulates strings of Typst markup which are compiled to a PDF by
the `typst` package (which bundles the compiler, so no external install needed).
"""

import logging
import os
import shutil
import tempfile
from datetime import datetime
from typing import Any, Dict, List, Optional, Union

from orcid_cv.content import (
    prepare_affiliations,
    prepare_funding,
    prepare_person,
    prepare_reviews,
    prepare_works,
)
from orcid_cv.utils import package_directory

logger = logging.getLogger("orcid_cv")


def _ensure_renderer(config: Dict[str, Any]) -> None:
    """Ensures the typst style renderer and asset registry are present in the config."""
    if "typst_assets" not in config:
        config["typst_assets"] = {}
    if "renderer" not in config:
        style = config.get("style", "")
        if style == "greenspon-default":
            from orcid_cv.typst_styles import GreensponDefaultTypstRenderer

            config["renderer"] = GreensponDefaultTypstRenderer(config)
        else:
            raise ValueError(f"Invalid style: {style}")


def process_external_links(
    link_dict: Dict[str, str], config: Dict[str, Any]
) -> Dict[str, str]:
    """
    Maps each website with an available icon to its URL and registers the icon
    file so that `build_document` can copy it next to the generated .typ source.
    """
    _ensure_renderer(config)
    icons: Dict[str, str] = {}
    for name, url in link_dict.items():
        im_path = os.path.join(package_directory, "external_link_img", f"{name}.png")
        if os.path.exists(im_path):
            asset_name = f"{name}.png"
            config["typst_assets"][asset_name] = im_path
            icons[asset_name] = url
        else:
            logger.warning(f"External link image missing at {im_path}")
    return icons


def _append_section(
    elements: List[str],
    config: Dict[str, Any],
    heading: str,
    blocks: List[str],
    spacing: Optional[float] = None,
) -> None:
    """
    Appends rendered entries to the document, keeping the section heading with
    its first entry and preventing any single entry from splitting across pages.
    """
    if not blocks:
        return

    renderer = config["renderer"]
    if spacing is None:
        spacing = config["item_spacing"]
    if heading:
        elements.append(f"#v({config['section_spacing']}pt, weak: true)")

    for i, block in enumerate(blocks):
        body = renderer.pad(block)
        if i == 0 and heading:
            body = f"{renderer.section_heading(heading)}\n{body}"
        elements.append(f"#block(breakable: false, width: 100%)[\n{body}\n]")
        elements.append(f"#v({spacing}pt, weak: true)")


def add_person_section(
    elements: List[str], orcid_dict: Dict[str, Any], config: Dict[str, Any]
) -> None:
    """Appends the name, title, current affiliation, email, and social links."""
    _ensure_renderer(config)
    person = prepare_person(orcid_dict)
    icons = process_external_links(person["links"], config)
    elements.append(config["renderer"].make_person_block(person, icons))


def add_affiliation_section(
    elements: List[str],
    orcid_dict: Dict[str, Any],
    config: Dict[str, Any],
    heading: str,
    affiliation_type: str,
) -> None:
    """Appends employments or educations as a stylized section."""
    _ensure_renderer(config)
    renderer = config["renderer"]
    blocks = [
        renderer.make_affiliation_block(af)
        for af in prepare_affiliations(orcid_dict, affiliation_type)
    ]
    _append_section(elements, config, heading, blocks)


def add_work_section(
    elements: List[str],
    orcid_dict: Dict[str, Any],
    config: Dict[str, Any],
    heading: str,
    search_str: Union[str, List[str]],
) -> None:
    """Appends the specified work categories as a stylized section."""
    _ensure_renderer(config)
    renderer = config["renderer"]
    blocks = [
        renderer.make_work_block(w)
        for w in prepare_works(orcid_dict, config, search_str)
    ]
    _append_section(elements, config, heading, blocks)


def add_funding_section(
    elements: List[str],
    orcid_dict: Dict[str, Any],
    config: Dict[str, Any],
    heading: str,
) -> None:
    """Appends funding entries as a stylized section."""
    _ensure_renderer(config)
    renderer = config["renderer"]
    blocks = [renderer.make_funding_block(f) for f in prepare_funding(orcid_dict)]
    _append_section(elements, config, heading, blocks)


def add_review_section(
    elements: List[str],
    orcid_dict: Dict[str, Any],
    config: Dict[str, Any],
    heading: str,
) -> None:
    """Appends peer review tallies grouped by journal in two columns."""
    _ensure_renderer(config)
    renderer = config["renderer"]
    reviews = prepare_reviews(orcid_dict)
    rows = [
        (reviews[i], reviews[i + 1] if i + 1 < len(reviews) else None)
        for i in range(0, len(reviews), 2)
    ]
    blocks = [renderer.make_review_block(row) for row in rows]
    # Review rows are a single line each, so they sit tighter than other entries
    _append_section(
        elements, config, heading, blocks, spacing=config["review_row_spacing"]
    )


def assemble_source(
    elements: List[str], config: Dict[str, Any], title: str = "", author: str = ""
) -> str:
    """Joins the preamble and every rendered element into one Typst document."""
    _ensure_renderer(config)
    if not config.get("footer_date"):
        config["footer_date"] = datetime.today().strftime("%d-%b-%Y")
    preamble = config["renderer"].preamble(title=title, author=author)
    return preamble + "\n".join(elements) + "\n"


def build_document(
    output_fname: str,
    elements: List[str],
    config: Dict[str, Any],
    title: str = "",
    author: str = "",
    save_source: Optional[str] = None,
) -> str:
    """
    Compiles the accumulated Typst markup into a PDF at `output_fname` and
    returns the generated Typst source. Pass `save_source` to also keep the .typ.
    """
    try:
        import typst
    except ImportError as e:  # pragma: no cover - depends on the environment
        raise ImportError(
            "The typst backend requires the 'typst' package: pip install typst"
        ) from e

    source = assemble_source(elements, config, title=title, author=author)

    # Build inside a scratch directory holding copies of every referenced asset,
    # so that typst's project root never needs to reach into the user's folders.
    build_dir = tempfile.mkdtemp(prefix="orcid_cv_typst_")
    try:
        for asset_name, asset_path in config.get("typst_assets", {}).items():
            shutil.copyfile(asset_path, os.path.join(build_dir, asset_name))

        typ_path = os.path.join(build_dir, "cv.typ")
        with open(typ_path, "w", encoding="utf-8") as f:
            f.write(source)

        output_dir = os.path.dirname(os.path.abspath(output_fname))
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        typst.compile(typ_path, output=output_fname, root=build_dir)
    finally:
        shutil.rmtree(build_dir, ignore_errors=True)

    if save_source:
        with open(save_source, "w", encoding="utf-8") as f:
            f.write(source)
        logger.info(f"Wrote typst source to {save_source}")

    return source
