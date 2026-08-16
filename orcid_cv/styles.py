import os
import logging
from typing import Dict, Any, List, Tuple
from reportlab.platypus import Paragraph, Table, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter

logger = logging.getLogger("orcid_cv")


class BaseStyleRenderer:
    """Abstract base class for CV style renderers."""
    def __init__(self, config: Dict[str, Any]):
        self.config = config

    def get_column_widths(self, section_type: str) -> List[float]:
        raise NotImplementedError()

    def _heading_table_style(self) -> List[Tuple[Any, ...]]:
        """
        Shared styles for the two rows that carry a section heading and its rule.
        The rule is drawn at the bottom of the empty row below the heading, so
        trimming that row's bottom padding lifts the rule closer to the text.
        """
        return [
            ("SPAN", (0, 0), (-1, 0)),
            (
                "LINEBELOW",
                (0, 1),
                (-1, 1),
                self.config["heading_rule_width"],
                colors.gray,
            ),
            ("TOPPADDING", (0, 1), (-1, 1), 0),
            ("BOTTOMPADDING", (0, 1), (-1, 1), self.config["heading_rule_padding"]),
        ]

    def make_affiliation_table(
        self, affiliation: Dict[str, Any], section_heading: str = ""
    ) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
        raise NotImplementedError()

    def make_work_table(
        self, work_title: str, work_body: Paragraph, work_date: str, section_heading: str = ""
    ) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
        raise NotImplementedError()

    def make_funding_table(
        self, fund: Dict[str, Any], section_heading: str = ""
    ) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
        raise NotImplementedError()

    def make_review_table(
        self, r: Tuple[str, int, str, int], section_heading: str = ""
    ) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
        raise NotImplementedError()

    def add_person_section(self, elements: List[Any], orcid_dict: Dict[str, Any]) -> None:
        raise NotImplementedError()


class GreensponDefaultRenderer(BaseStyleRenderer):
    """Renderer implementation for the 'greenspon-default' CV style."""
    
    def get_column_widths(self, section_type: str) -> List[float]:
        ratio = 1.0
        table_width = int(self.config["pagesize"][0] - (self.config["margin"] * 2))
        
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

        right_col_width = round(table_width / ratio)
        return [table_width - right_col_width, right_col_width]

    def make_affiliation_table(
        self, affiliation: Dict[str, Any], section_heading: str = ""
    ) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
        # Service entries do not always carry a department, so only join the
        # parts that are actually present.
        body = ", ".join(
            p
            for p in (
                affiliation.get("organization", ""),
                affiliation.get("department", ""),
            )
            if p
        )
        affiliation_head = [
            Paragraph(affiliation["role"], style=self.config["item_title_style"]),
            Paragraph(affiliation["date_range"], style=self.config["item_date_style"]),
        ]
        affiliation_body = [
            Paragraph(body, style=self.config["item_body_style"]),
            "",
        ]
        if section_heading == "":
            table_data = [affiliation_head, affiliation_body]
            table_style = [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
            ]
        else:
            table_data = [
                [Paragraph(section_heading, style=self.config["section_style"]), ""],
                ["", ""],
                affiliation_head,
                affiliation_body,
            ]
            table_style = self._heading_table_style() + [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
            ]
        return table_data, table_style

    def make_work_table(
        self, work_title: str, work_body: Paragraph, work_date: str, section_heading: str = ""
    ) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
        if section_heading == "":
            table_data = [
                [
                    Paragraph(work_title, style=self.config["item_title_style"]),
                    Paragraph(str(work_date), style=self.config["item_date_style"]),
                ],
                [work_body, ""],
            ]
            table_style = [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
            ]
        else:
            table_data = [
                [Paragraph(section_heading, style=self.config["section_style"]), ""],
                ["", ""],
                [
                    Paragraph(work_title, style=self.config["item_title_style"]),
                    Paragraph(str(work_date), style=self.config["item_date_style"]),
                ],
                [work_body, ""],
            ]
            table_style = self._heading_table_style() + [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
            ]
        return table_data, table_style

    def make_funding_table(
        self, fund: Dict[str, Any], section_heading: str = ""
    ) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
        fund_head = [
            Paragraph(fund["title"], style=self.config["item_title_style"]),
            Paragraph(fund["start_year"], style=self.config["item_date_style"]),
        ]
        fund_body = [
            f"{fund['org']}, {fund['id']}",
            Paragraph(fund["role"], style=self.config["item_misc_style"]),
        ]
        if section_heading == "":
            table_data = [fund_head, fund_body]
            table_style = [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
            ]
        else:
            table_data = [
                [Paragraph(section_heading, style=self.config["section_style"]), ""],
                ["", ""],
                fund_head,
                fund_body,
            ]
            table_style = self._heading_table_style() + [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
            ]
        return table_data, table_style

    def make_review_table(
        self, r: Tuple[str, int, str, int], section_heading: str = ""
    ) -> Tuple[List[List[Any]], List[Tuple[Any, ...]]]:
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
                [Paragraph(section_heading, style=self.config["section_style"]), ""],
                ["", ""],
                [rev_body1, rev_body2],
            ]
            table_style = self._heading_table_style() + [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("NOSPLIT", (0, 0), (-1, -1)),
                ("HALIGN", (0, 0), (-1, -1), "LEFT"),
            ]
        return table_data, table_style

    def add_person_section(self, elements: List[Any], orcid_dict: Dict[str, Any]) -> None:
        from orcid_cv.utils import dict_to_list
        from orcid_cv.builder import process_external_links
        
        column_widths = self.get_column_widths("person")
        info = dict_to_list(orcid_dict["employment"])
        
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
                Paragraph(fullname, style=self.config["person_title_style"]),
                Paragraph(person_summary, style=self.config["person_summary_style"]),
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
            
        elements.append(Spacer(0, self.config["item_spacing"]))
