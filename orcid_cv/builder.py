import os
import logging
from datetime import datetime
from typing import List, Dict, Any, Tuple, Union

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Paragraph, Spacer, Table, SimpleDocTemplate, Image
from reportlab.lib import colors

from orcid_cv.utils import (
    package_directory,
    dict_to_list,
    initialize_name,
    embolden_authors,
)
from orcid_cv.parser import extract_orcid_info

logger = logging.getLogger("orcid_cv")


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
    config: Dict[str, Any], work_title: str, work_body: Paragraph, work_date: str, section_heading: str = ""
) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
    """Generates the table data and layout styles for a publication/work entry."""
    _ensure_renderer(config)
    return config["renderer"].make_work_table(work_title, work_body, work_date, section_heading)


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
            link_list.append(HyperlinkedImage(im_path, hyperlink=v, height=15, width=15))
        else:
            logger.warning(f"External link image missing at {im_path}")
    return link_list


def add_person_section(elements: List[Any], orcid_dict: Dict[str, Any], config: Dict[str, Any]) -> None:
    """Appends the name, title, current affiliation, email, and social links to the CV flowables."""
    _ensure_renderer(config)
    config["renderer"].add_person_section(elements, orcid_dict)



def add_affiliation_section(
    elements: List[Any], orcid_dict: Dict[str, Any], config: Dict[str, Any], heading: str, affiliation_type: str
) -> None:
    """Appends employments or educations as a stylized table to the CV flowables."""
    column_widths = get_column_widths(config, "affiliation")

    if affiliation_type not in orcid_dict:
        logger.warning(f"Dict does not contain affiliation type: {affiliation_type}")
        return

    affiliations = dict_to_list(orcid_dict[affiliation_type])
    try:
        affiliations = sorted(
            affiliations, key=lambda v: int(v.get("start_date", 0)), reverse=True
        )
    except (ValueError, TypeError):
        pass

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


def add_work_section(
    elements: List[Any], orcid_dict: Dict[str, Any], config: Dict[str, Any], heading: str, search_str: Union[str, List[str]]
) -> None:
    """Appends specified work categories as a stylized table to the CV flowables."""
    column_widths = get_column_widths(config, "work")
    if isinstance(search_str, str):
        search_str = [search_str]

    # Get subset of publications
    works = dict_to_list(orcid_dict["work"])
    works = [w for w in works if w.get("type") in search_str]
    if not works:
        logger.warning(f"No matching works for: {search_str}")
        return

    try:
        works = sorted(
            works,
            key=lambda v: int(v.get("year", 0)) * 1000 + int(v.get("month", 0)),
            reverse=True,
        )
    except (ValueError, TypeError):
        pass

    is_heading = True
    for work in works:
        work_date = str(work.get("year", ""))
        work_journal = work.get("journal", "")
        work_title = work.get("title", "")
        
        if "‐" in work_title:
            work_title = work_title.replace("‐", "-")
            
        if work.get("type") == "software":
            work_journal = work.get("subtitle", "")
            work_date = str(work.get("journal", ""))
            work["subtitle"] = ""

        # Process authors
        author_cat = ""
        author_list = list(work.get("authors", []))
        if author_list:
            if config.get("initalize_authors"):
                author_list = [initialize_name(i) for i in author_list]
            if config.get("embolden_author"):
                author_list = embolden_authors(orcid_dict["personal"], author_list)
            
            if len(author_list) == 1:
                author_cat = author_list[0]
            elif len(author_list) == 2:
                author_cat = f"{author_list[0]} and {author_list[1]}"
            else:
                author_list[-1] = "and " + author_list[-1]
                author_cat = ", ".join(author_list)
                
        if "‐" in author_cat:
            author_cat = author_cat.replace("‐", "-")
        if "\u0159" in author_cat:
            author_cat = author_cat.replace("\u0159", "r")

        # Process DOI/link
        doi_str = work.get("doi", "")
        if doi_str == "":
            work_str = work_journal
        else:
            if "doi.org/" in doi_str:
                idx = doi_str.find("doi.org/")
                short_doi = doi_str[idx + 8 :]
                doi_str = f'<link href="https://www.doi.org/{short_doi}">DOI: <u>{short_doi} </u></link>'
            elif "github.com/" in doi_str:
                idx = doi_str.find("github.com/")
                short_doi = doi_str[idx + 11 :]
                doi_str = f'<link href="https://www.github.com/{short_doi}">GitHub: <u>{short_doi} </u></link>'
            else:
                doi_str = f'<link href="{doi_str}"><u>{doi_str} </u></link>'
            
            work_str = f"{work_journal}, {doi_str}"

        if work.get("subtitle") and work.get("subtitle") != "":
            work_str = f"{work_str}, {work['subtitle']}"

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
    elements: List[Any], orcid_dict: Dict[str, Any], config: Dict[str, Any], heading: str
) -> None:
    """Appends funding entries as a stylized table to the CV flowables."""
    column_widths = get_column_widths(config, "affiliation")

    fund = dict_to_list(orcid_dict["funding"])
    try:
        fund = sorted(fund, key=lambda v: int(v.get("start_year", 0)), reverse=True)
    except (ValueError, TypeError):
        pass

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
    elements: List[Any], orcid_dict: Dict[str, Any], config: Dict[str, Any], heading: str
) -> None:
    """Appends peer review summaries grouped by journal as a multi-column table."""
    if "reviews" not in orcid_dict:
        return
        
    column_widths = get_column_widths(config, "review")

    # Count reviews per unique Org
    review_dict = {}
    for v in orcid_dict["reviews"].values():
        org_name = v.get("org", "")
        if org_name:
            review_dict[org_name] = review_dict.get(org_name, 0) + 1

    if not review_dict:
        return

    # Sort and make even for columns
    sk = sorted(review_dict.keys())
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


def quick_build(orcid_dir: str, output_fname: str, style: str = "greenspon-default") -> None:
    """Convenience method to construct and save a standard CV using default layout choices."""
    from orcid_cv.config import make_document_config
    
    orcid_dict = extract_orcid_info(orcid_dir)
    config = make_document_config(style)
    
    doc_title = f"{orcid_dict['personal']['fullname']} - CV"
    doc = SimpleDocTemplate(
        output_fname,
        pagesize=letter,
        leftMargin=config["margin"],
        rightMargin=config["margin"],
        topMargin=config["margin"],
        bottomMargin=config["margin"],
        title=doc_title,
    )
    
    elements = []
    add_person_section(elements, orcid_dict, config)
    add_affiliation_section(elements, orcid_dict, config, "Employment", "employment")
    add_affiliation_section(elements, orcid_dict, config, "Education", "education")
    add_work_section(
        elements, orcid_dict, config, "Research Publications", "journal-article"
    )
    add_work_section(elements, orcid_dict, config, "Talks", "public-speech")
    add_work_section(elements, orcid_dict, config, "Preprints", "preprint")
    
    doc.build(elements)
    print("Success!")
