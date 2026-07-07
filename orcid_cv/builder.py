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


def get_column_widths(config: Dict[str, Any], section_type: str) -> List[float]:
    """Computes column widths based on the selected design style."""
    ratio = 1.0
    table_width = int(config["pagesize"][0] - (config["margin"] * 2))
    
    if config["style"] == "greenspon-default":
        if section_type == "work":
            ratio = 7.0
        elif section_type == "affiliation":
            ratio = 6.0
        elif section_type == "person":
            ratio = 3.5
        elif section_type == "review":
            ratio = 2.0
        else:
            raise ValueError(f"Unknown section type: {section_type}")
    else:
        raise ValueError(f"Invalid style: {config['style']}")

    right_col_width = round(table_width / ratio)
    return [table_width - right_col_width, right_col_width]


def make_affiliation_table(
    config: Dict[str, Any], affiliation: Dict[str, Any], section_heading: str = ""
) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
    """Generates the table data and layout styles for an affiliation entry."""
    table_data = []
    table_style = []
    
    if config["style"] == "greenspon-default":
        if section_heading == "":
            table_data = [
                [
                    Paragraph(affiliation["role"], style=config["item_title_style"]),
                    Paragraph(affiliation["date_range"], style=config["item_date_style"]),
                ],
                [
                    Paragraph(
                        affiliation["organization"] + ", " + affiliation["department"],
                        style=config["item_body_style"],
                    ),
                    "",
                ],
            ]
            table_style = [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
            ]
        else:
            table_data = [
                [Paragraph(section_heading, style=config["section_style"]), ""],
                ["", ""],
                [
                    Paragraph(affiliation["role"], style=config["item_title_style"]),
                    Paragraph(affiliation["date_range"], style=config["item_date_style"]),
                ],
                [
                    Paragraph(
                        affiliation["organization"] + ", " + affiliation["department"],
                        style=config["item_body_style"],
                    ),
                    "",
                ],
            ]
            table_style = [
                ("SPAN", (0, 0), (-1, 0)),
                ("LINEBELOW", (0, 1), (-1, 1), 2, colors.gray),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
            ]
    else:
        raise ValueError(f"Invalid style: {config['style']}")

    return table_data, table_style


def make_work_table(
    config: Dict[str, Any], work_title: str, work_body: Paragraph, work_date: str, section_heading: str = ""
) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
    """Generates the table data and layout styles for a publication/work entry."""
    table_data = []
    table_style = []
    
    if work_date == 0 or work_date == "0":
        work_date = ""

    if config["style"] == "greenspon-default":
        if section_heading == "":
            table_data = [
                [
                    Paragraph(work_title, style=config["item_title_style"]),
                    Paragraph(str(work_date), style=config["item_date_style"]),
                ],
                [work_body, ""],
            ]
            table_style = [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
            ]
        else:
            table_data = [
                [Paragraph(section_heading, style=config["section_style"]), ""],
                ["", ""],
                [
                    Paragraph(work_title, style=config["item_title_style"]),
                    Paragraph(str(work_date), style=config["item_date_style"]),
                ],
                [work_body, ""],
            ]
            table_style = [
                ("SPAN", (0, 0), (-1, 0)),
                ("LINEBELOW", (0, 1), (-1, 1), 2, colors.gray),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
            ]
    else:
        raise ValueError(f"Invalid style: {config['style']}")

    return table_data, table_style


def make_funding_table(
    config: Dict[str, Any], fund: Dict[str, Any], section_heading: str = ""
) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
    """Generates the table data and layout styles for a funding entry."""
    table_data = []
    table_style = []
    
    if config["style"] == "greenspon-default":
        fund_head = [
            Paragraph(fund["title"], style=config["item_title_style"]),
            Paragraph(fund["start_year"], style=config["item_date_style"]),
        ]
        fund_body = [
            f"{fund['org']}, {fund['id']}",
            Paragraph(fund["role"], style=config["item_misc_style"]),
        ]
        if section_heading == "":
            table_data = [fund_head, fund_body]
            table_style = [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
            ]
        else:
            table_data = [
                [Paragraph(section_heading, style=config["section_style"]), ""],
                ["", ""],
                fund_head,
                fund_body,
            ]
            table_style = [
                ("SPAN", (0, 0), (-1, 0)),
                ("LINEBELOW", (0, 1), (-1, 1), 2, colors.gray),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
            ]
    else:
        raise ValueError(f"Invalid style: {config['style']}")

    return table_data, table_style


def make_review_table(
    config: Dict[str, Any], r: Tuple[str, int, str, int], section_heading: str = ""
) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
    """Generates the table data and layout styles for a peer review entry."""
    table_data = []
    table_style = []
    
    if config["style"] == "greenspon-default":
        # Format columns
        rev_body1 = f"{r[0]}, {r[1]} reviews" if r[1] > 1 else f"{r[0]}, 1 review"
        
        if r[3] == 0:
            rev_body2 = ""
        else:
            rev_body2 = f"{r[2]}, {r[3]} reviews" if r[3] > 1 else f"{r[2]}, 1 review"

        if section_heading == "":
            table_data = [[rev_body1, rev_body2]]
            table_style = [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
                ("HALIGN", (0, 0), (-1, -1), "LEFT"),
            ]
        else:
            table_data = [
                [Paragraph(section_heading, style=config["section_style"]), ""],
                ["", ""],
                [rev_body1, rev_body2],
            ]
            table_style = [
                ("SPAN", (0, 0), (-1, 0)),
                ("LINEBELOW", (0, 1), (-1, 1), 2, colors.gray),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
                ("HALIGN", (0, 0), (-1, -1), "LEFT"),
            ]
    else:
        raise ValueError(f"Invalid style: {config['style']}")

    return table_data, table_style


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
    if config["style"] == "greenspon-default":
        column_widths = get_column_widths(config, "person")
        info = dict_to_list(orcid_dict["employment"])
        
        # Sort by start_date to find current role
        try:
            info = sorted(info, key=lambda v: int(v.get("start_date", 0)), reverse=True)
        except (ValueError, TypeError):
            pass
            
        fullname = orcid_dict["personal"]["fullname"]
        
        if info:
            current_role = info[0].get("role", "")
            current_org = info[0].get("organization", "")
            person_summary = f"<br/>{current_role}<br/>{current_org}<br/>{orcid_dict['personal']['email']}"
        else:
            person_summary = f"<br/>{orcid_dict['personal']['email']}"
            
        table_data = [
            [
                Paragraph(fullname, style=config["person_title_style"]),
                Paragraph(person_summary, style=config["person_summary_style"]),
            ]
        ]
        table_style = [
            ("NOSPLIT", (0, 0), (-1, -1)),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]
        
        t = Table(table_data, colWidths=column_widths)
        t.setStyle(table_style)
        elements.append(t)
        elements.append(Spacer(0, -15))

        # Add links
        link_list = process_external_links(orcid_dict["personal"]["links"])
        if link_list:
            t = Table([link_list], colWidths=[20] * len(link_list), hAlign="LEFT")
            t.setStyle(table_style)
            elements.append(t)
            
        elements.append(Spacer(0, config["item_spacing"]))
    else:
        raise ValueError(f"Invalid style: {config['style']}")


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
