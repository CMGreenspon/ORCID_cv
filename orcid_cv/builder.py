import os
import logging
from datetime import datetime
from typing import List, Dict, Any, Tuple, Union

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Paragraph, Spacer, Table, SimpleDocTemplate, Image

from orcid_cv.utils import package_directory
from orcid_cv.content import (
    join_authors,
    prepare_affiliations,
    prepare_funding,
    prepare_reviews,
    prepare_service,
    prepare_works,
)
from orcid_cv.parser import extract_orcid_info

logger = logging.getLogger("orcid_cv")


def _typst_delegate(config: Dict[str, Any], function_name: str):
    """
    Returns the typst implementation of `function_name` when the config asks for
    the typst backend, otherwise None so the caller renders with reportlab.
    """
    if config.get("backend", "reportlab") == "typst":
        from orcid_cv import typst_builder

        return getattr(typst_builder, function_name)
    return None


class HyperlinkedImage(Image, object):
    """An Image subclass that overlays a clickable hyperlink on the rendered PDF canvas."""

    def __init__(
        self,
        filename: str,
        hyperlink: str = None,
        width: float = None,
        height: float = None,
        kind: str = "direct",
        mask: str = "auto",
        lazy: int = 1,
    ):
        super(HyperlinkedImage, self).__init__(
            filename, width, height, kind, mask, lazy
        )
        self.hyperlink = hyperlink

    def drawOn(self, canvas: canvas.Canvas, x: float, y: float, _sW: float = 0) -> None:
        if self.hyperlink:
            x1 = x
            y1 = y
            x2 = x1 + self._width
            y2 = y1 + self._height
            canvas.linkURL(
                url=self.hyperlink, rect=(x1, y1, x2, y2), thickness=0, relative=1
            )
        super(HyperlinkedImage, self).drawOn(canvas, x, y, _sW)


class FooterCanvas(canvas.Canvas):
    """Custom canvas that captures page attributes to draw a 'Page X of Y' footer on save."""

    left_str: str = ""

    def __init__(self, *args, **kwargs):
        super(FooterCanvas, self).__init__(*args, **kwargs)
        self.pages = []

    def showPage(self) -> None:
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self) -> None:
        page_count = len(self.pages)
        for page in self.pages:
            self.__dict__.update(page)
            self.draw_canvas(page_count)
            super(FooterCanvas, self).showPage()
        super(FooterCanvas, self).save()

    def draw_canvas(self, page_count: int) -> None:
        page_str = f"Page {self._pageNumber} of {page_count}"
        y = 40
        self.saveState()
        self.setStrokeColorRGB(0, 0, 0)
        self.setLineWidth(0.5)
        self.line(40, y + 10, letter[0] - 40, y + 10)
        self.setFont("Helvetica", 9)
        self.drawString(letter[0] - 90, y, page_str)
        self.drawString(40, y, datetime.today().strftime("%d-%b-%Y"))
        self.restoreState()


def _ensure_renderer(config: Dict[str, Any]) -> None:
    """Ensures that the style renderer is present in the configuration dictionary."""
    if "renderer" not in config:
        style = config.get("style", "")
        if style == "greenspon-default":
            from orcid_cv.styles import GreensponDefaultRenderer

            config["renderer"] = GreensponDefaultRenderer(config)
        else:
            raise ValueError(f"Invalid style: {style}")


def get_column_widths(config: Dict[str, Any], section_type: str) -> List[float]:
    """Computes column widths based on the selected design style."""
    _ensure_renderer(config)
    return config["renderer"].get_column_widths(section_type)


def make_affiliation_table(
    config: Dict[str, Any], affiliation: Dict[str, Any], section_heading: str = ""
) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
    """Generates the table data and layout styles for an affiliation entry."""
    _ensure_renderer(config)
    return config["renderer"].make_affiliation_table(affiliation, section_heading)


def make_work_table(
    config: Dict[str, Any],
    work_title: str,
    work_body: Paragraph,
    work_date: str,
    section_heading: str = "",
) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
    """Generates the table data and layout styles for a publication/work entry."""
    _ensure_renderer(config)
    return config["renderer"].make_work_table(
        work_title, work_body, work_date, section_heading
    )


def make_funding_table(
    config: Dict[str, Any], fund: Dict[str, Any], section_heading: str = ""
) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
    """Generates the table data and layout styles for a funding entry."""
    _ensure_renderer(config)
    return config["renderer"].make_funding_table(fund, section_heading)


def make_review_table(
    config: Dict[str, Any], r: Tuple[str, int, str, int], section_heading: str = ""
) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
    """Generates the table data and layout styles for a peer review entry."""
    _ensure_renderer(config)
    return config["renderer"].make_review_table(r, section_heading)


def process_external_links(link_dict: Dict[str, str]) -> List[HyperlinkedImage]:
    """Converts a dictionary of website titles and URLs into hyperlinked image flowables."""
    link_list = []
    for k, v in link_dict.items():
        im_path = os.path.join(package_directory, "external_link_img", f"{k}.png")
        if os.path.exists(im_path):
            link_list.append(
                HyperlinkedImage(im_path, hyperlink=v, height=15, width=15)
            )
        else:
            logger.warning(f"External link image missing at {im_path}")
    return link_list


def add_person_section(
    elements: List[Any], orcid_dict: Dict[str, Any], config: Dict[str, Any]
) -> None:
    """Appends the name, title, current affiliation, email, and social links to the CV flowables."""
    typst = _typst_delegate(config, "add_person_section")
    if typst:
        return typst(elements, orcid_dict, config)

    _ensure_renderer(config)
    config["renderer"].add_person_section(elements, orcid_dict)


def add_affiliation_section(
    elements: List[Any],
    orcid_dict: Dict[str, Any],
    config: Dict[str, Any],
    heading: str,
    affiliation_type: str,
) -> None:
    """Appends employments or educations as a stylized table to the CV flowables."""
    typst = _typst_delegate(config, "add_affiliation_section")
    if typst:
        return typst(elements, orcid_dict, config, heading, affiliation_type)

    column_widths = get_column_widths(config, "affiliation")
    affiliations = prepare_affiliations(orcid_dict, affiliation_type)

    is_heading = True
    for af in affiliations:
        if is_heading:
            table_data, table_style = make_affiliation_table(
                config, af, section_heading=heading
            )
            is_heading = False
        else:
            table_data, table_style = make_affiliation_table(config, af)

        t = Table(table_data, colWidths=column_widths)
        t.setStyle(table_style)
        elements.append(t)
        elements.append(Spacer(0, config["item_spacing"]))


def add_service_section(
    elements: List[Any],
    orcid_dict: Dict[str, Any],
    config: Dict[str, Any],
    heading: str,
    match: Union[str, List[str], None] = None,
    exclude: Union[str, List[str], None] = None,
) -> None:
    """
    Appends mentorship/service entries as a stylized table to the CV flowables.

    Service records share their shape with employments and educations, so they
    are laid out by the same style hooks. `match` and `exclude` filter on the
    role title, letting mentorship and other service become separate sections.
    """
    typst = _typst_delegate(config, "add_service_section")
    if typst:
        return typst(
            elements, orcid_dict, config, heading, match=match, exclude=exclude
        )

    column_widths = get_column_widths(config, "affiliation")
    services = prepare_service(orcid_dict, match=match, exclude=exclude)

    is_heading = True
    for sv in services:
        if is_heading:
            table_data, table_style = make_affiliation_table(
                config, sv, section_heading=heading
            )
            is_heading = False
        else:
            table_data, table_style = make_affiliation_table(config, sv)

        t = Table(table_data, colWidths=column_widths)
        t.setStyle(table_style)
        elements.append(t)
        elements.append(Spacer(0, config["item_spacing"]))


def add_work_section(
    elements: List[Any],
    orcid_dict: Dict[str, Any],
    config: Dict[str, Any],
    heading: str,
    search_str: Union[str, List[str]],
) -> None:
    """Appends specified work categories as a stylized table to the CV flowables."""
    typst = _typst_delegate(config, "add_work_section")
    if typst:
        return typst(elements, orcid_dict, config, heading, search_str)

    column_widths = get_column_widths(config, "work")
    works = prepare_works(orcid_dict, config, search_str)

    is_heading = True
    for work in works:
        work_title = work["title"]
        work_date = work["date"]

        author_cat = join_authors(work["authors"], bold=lambda name: f"<b>{name}</b>")

        # Process DOI/link
        link = work["link"]
        if link is None:
            work_str = work["journal"]
        else:
            anchor = (
                f'<link href="{link["url"]}">{link["prefix"]}'
                f'<u>{link["label"]} </u></link>'
            )
            work_str = f'{work["journal"]}, {anchor}'

        if work["subtitle"]:
            work_str = f'{work_str}, {work["subtitle"]}'

        work_body = Paragraph(
            f"{work_str}<br/>{author_cat}", style=config["item_body_style"]
        )

        if is_heading:
            table_data, table_style = make_work_table(
                config, work_title, work_body, work_date, section_heading=heading
            )
            is_heading = False
        else:
            table_data, table_style = make_work_table(
                config, work_title, work_body, work_date, section_heading=""
            )

        t = Table(table_data, colWidths=column_widths)
        t.setStyle(table_style)
        elements.append(t)
        elements.append(Spacer(0, config["item_spacing"]))


def add_funding_section(
    elements: List[Any],
    orcid_dict: Dict[str, Any],
    config: Dict[str, Any],
    heading: str,
) -> None:
    """Appends funding entries as a stylized table to the CV flowables."""
    typst = _typst_delegate(config, "add_funding_section")
    if typst:
        return typst(elements, orcid_dict, config, heading)

    column_widths = get_column_widths(config, "affiliation")
    fund = prepare_funding(orcid_dict)

    is_heading = True
    for f in fund:
        if is_heading:
            table_data, table_style = make_funding_table(
                config, f, section_heading=heading
            )
            is_heading = False
        else:
            table_data, table_style = make_funding_table(config, f)

        t = Table(table_data, colWidths=column_widths)
        t.setStyle(table_style)
        elements.append(t)
        elements.append(Spacer(0, config["item_spacing"]))


def add_review_section(
    elements: List[Any],
    orcid_dict: Dict[str, Any],
    config: Dict[str, Any],
    heading: str,
) -> None:
    """Appends peer review summaries grouped by journal as a multi-column table."""
    typst = _typst_delegate(config, "add_review_section")
    if typst:
        return typst(elements, orcid_dict, config, heading)

    reviews = prepare_reviews(orcid_dict)
    if not reviews:
        return

    column_widths = get_column_widths(config, "review")
    review_dict = dict(reviews)

    # Sort and make even for columns
    sk = [org for org, _ in reviews]
    if len(sk) % 2 == 1:
        sk.append("")
        review_dict[""] = 0

    is_heading = True
    for i in range(0, len(sk), 2):
        if is_heading:
            table_data, table_style = make_review_table(
                config,
                (sk[i], review_dict[sk[i]], sk[i + 1], review_dict[sk[i + 1]]),
                section_heading=heading,
            )
            is_heading = False
        else:
            table_data, table_style = make_review_table(
                config, (sk[i], review_dict[sk[i]], sk[i + 1], review_dict[sk[i + 1]])
            )

        t = Table(table_data, colWidths=column_widths)
        t.setStyle(table_style)
        elements.append(t)
        elements.append(Spacer(0, config["item_spacing"]))


def build_document(
    output_fname: str,
    elements: List[Any],
    config: Dict[str, Any],
    title: str = "",
    author: str = "",
    **kwargs: Any,
) -> Any:
    """
    Renders the accumulated elements to `output_fname` with whichever backend the
    config selects, so the same script can target reportlab or typst.

    Set config['page_footer'] = True for a 'Page X of Y' footer with today's date.
    """
    typst = _typst_delegate(config, "build_document")
    if typst:
        return typst(
            output_fname, elements, config, title=title, author=author, **kwargs
        )

    bottom_margin = (
        config["margin"] + 10 if config.get("page_footer") else config["margin"]
    )
    doc = SimpleDocTemplate(
        output_fname,
        pagesize=config["pagesize"],
        leftMargin=config["margin"],
        rightMargin=config["margin"],
        topMargin=config["margin"],
        bottomMargin=bottom_margin,
        title=title,
        author=author,
        **kwargs,
    )

    if config.get("page_footer"):
        doc.multiBuild(elements, canvasmaker=FooterCanvas)
    else:
        doc.build(elements)
    return None


def quick_build(
    orcid_dir: str,
    output_fname: str,
    style: str = "greenspon-default",
    backend: str = "reportlab",
) -> None:
    """
    Convenience method to construct and save a standard CV using default layout
    choices. `backend` selects the PDF engine: 'reportlab' or 'typst'.
    """
    from orcid_cv.config import make_document_config

    orcid_dict = extract_orcid_info(orcid_dir)
    config = make_document_config(style, backend=backend)

    fullname = orcid_dict["personal"]["fullname"]
    elements: List[Any] = []
    add_person_section(elements, orcid_dict, config)
    add_affiliation_section(elements, orcid_dict, config, "Employment", "employment")
    add_affiliation_section(elements, orcid_dict, config, "Education", "education")
    add_work_section(
        elements, orcid_dict, config, "Research Publications", "journal-article"
    )
    add_work_section(elements, orcid_dict, config, "Talks", "public-speech")
    add_work_section(elements, orcid_dict, config, "Preprints", "preprint")
    add_service_section(elements, orcid_dict, config, "Mentorship & Service")

    build_document(
        output_fname, elements, config, title=f"{fullname} - CV", author=fullname
    )
    print("Success!")
