"""
Typst style renderers.

These mirror the reportlab renderers in `orcid_cv.styles`, but instead of
returning table data and table styles they return snippets of Typst markup.
"""

import logging
from typing import Any, Dict, List, Optional, Tuple

from orcid_cv.content import Author, format_review, join_authors

logger = logging.getLogger("orcid_cv")


def escape(text: Any) -> str:
    """
    Wraps arbitrary text as a Typst string literal expression so that no
    character in the ORCID record can be mistaken for Typst markup.
    """
    s = "" if text is None else str(text)
    s = s.replace("\\", "\\\\").replace('"', '\\"')
    s = s.replace("\n", "\\n").replace("\r", "").replace("\t", "\\t")
    return f'#("{s}")'


def raw_str(text: Any) -> str:
    """Escapes text for use as a bare Typst string literal (e.g. a function argument)."""
    s = "" if text is None else str(text)
    s = s.replace("\\", "\\\\").replace('"', '\\"')
    s = s.replace("\n", "\\n").replace("\r", "").replace("\t", "\\t")
    return f'"{s}"'


class BaseTypstStyleRenderer:
    """Abstract base class for Typst CV style renderers."""

    def __init__(self, config: Dict[str, Any]):
        self.config = config

    def preamble(self, title: str = "", author: str = "") -> str:
        raise NotImplementedError()

    def section_heading(self, heading: str) -> str:
        raise NotImplementedError()

    def make_person_block(self, person: Dict[str, Any], icons: Dict[str, str]) -> str:
        raise NotImplementedError()

    def make_affiliation_block(self, affiliation: Dict[str, Any]) -> str:
        raise NotImplementedError()

    def make_work_block(self, work: Dict[str, Any]) -> str:
        raise NotImplementedError()

    def make_funding_block(self, fund: Dict[str, Any]) -> str:
        raise NotImplementedError()

    def make_review_block(
        self, row: Tuple[Tuple[str, int], Optional[Tuple[str, int]]]
    ) -> str:
        raise NotImplementedError()


class GreensponDefaultTypstRenderer(BaseTypstStyleRenderer):
    """Typst counterpart of `styles.GreensponDefaultRenderer`."""

    def _font(self) -> str:
        fonts = self.config["font_family"]
        return "(" + ", ".join(raw_str(f) for f in fonts) + ")"

    def _column_ratio(self, section_type: str) -> float:
        """Right-hand column width as a fraction of the text width."""
        if section_type == "work":
            return 7.0
        if section_type == "affiliation":
            return 6.0
        if section_type == "person":
            return 3.5
        if section_type == "review":
            return 2.0
        raise ValueError(f"Unknown section type: {section_type}")

    def _columns(self, section_type: str) -> str:
        right = round(100.0 / self._column_ratio(section_type), 4)
        return f"({100 - right}%, {right}%)"

    def pad(self, body: str) -> str:
        """Insets content by the cell padding so that only the rules run full width."""
        padding = self.config["cell_padding"]
        if not padding:
            return body
        return f"#pad(x: {padding}pt)[{body}]"

    def preamble(self, title: str = "", author: str = "") -> str:
        cfg = self.config
        footer = "none"
        if cfg.get("page_footer"):
            footer = (
                "context {\n"
                "    let total = counter(page).final().first()\n"
                "    block(width: 100%)[\n"
                f"      #line(length: 100%, stroke: {cfg['footer_rule_width']}pt + black)\n"
                f"      #v({cfg['footer_gap']}pt, weak: true)\n"
                "      #grid(columns: (50%, 50%), align: (left, right),\n"
                f"        text(size: {cfg['footer_font_size']}pt)[{escape(cfg['footer_date'])}],\n"
                f"        text(size: {cfg['footer_font_size']}pt)[Page #counter(page).display() of #total],\n"
                "      )\n"
                "    ]\n"
                "  }"
            )

        return "\n".join(
            [
                f"#set document(title: {raw_str(title)}, author: {raw_str(author)})",
                "#set page(",
                f"  paper: {raw_str(cfg['paper'])},",
                f"  margin: (x: {cfg['margin']}pt, top: {cfg['margin']}pt, "
                f"bottom: {cfg['bottom_margin']}pt),",
                f"  footer: {footer},",
                f"  footer-descent: {cfg['footer_descent']}pt,",
                ")",
                f"#set text(font: {self._font()}, size: {cfg['body_font_size']}pt, "
                "hyphenate: false)",
                # Absolute leading (like reportlab's) and no automatic block spacing,
                # so every gap in the CV comes from an explicit #v() below.
                f"#set par(leading: {cfg['line_leading']}pt, "
                f"spacing: {cfg['par_spacing']}pt, justify: false)",
                "",
                "// Layout helpers shared by every entry in this style",
                "#let cv-entry(left-body, right-body, columns) = grid(",
                "  columns: columns, align: (left + top, right + top),",
                "  left-body, right-body,",
                ")",
                "",
            ]
        )

    def section_heading(self, heading: str) -> str:
        cfg = self.config
        return "\n".join(
            [
                self.pad(
                    f"#text(size: {cfg['section_font_size']}pt, weight: \"bold\")"
                    f"[{escape(heading)}]"
                ),
                f"#v({cfg['heading_rule_gap']}pt)",
                f"#line(length: 100%, stroke: {cfg['heading_rule_width']}pt + gray)",
                f"#v({cfg['heading_body_gap']}pt)",
            ]
        )

    def _title_date_grid(self, title_markup: str, date: str, section_type: str) -> str:
        cfg = self.config
        left = (
            f"text(size: {cfg['item_title_font_size']}pt, weight: \"bold\")"
            f"[{title_markup}]"
        )
        right = (
            f"text(size: {cfg['item_date_font_size']}pt, weight: \"bold\")"
            f"[{escape(date)}]"
        )
        return f"#cv-entry({left}, {right}, {self._columns(section_type)})"

    def make_person_block(self, person: Dict[str, Any], icons: Dict[str, str]) -> str:
        cfg = self.config

        summary_lines = [
            person.get("role", ""),
            person.get("organization", ""),
            person.get("email", ""),
        ]
        summary = "#linebreak()".join(escape(line) for line in summary_lines if line)

        icon_markup = [
            f"box(link({raw_str(url)})"
            f"[#image({raw_str(name)}, width: {cfg['icon_size']}pt, "
            f"height: {cfg['icon_size']}pt)])"
            for name, url in icons.items()
        ]

        left_parts = [
            f"#text(size: {cfg['person_title_font_size']}pt, weight: \"bold\")"
            f"[{escape(person.get('fullname', ''))}]"
        ]
        if icon_markup:
            left_parts.append(f"#v({cfg['icon_gap']}pt)")
            left_parts.append(
                f"#stack(dir: ltr, spacing: {cfg['icon_spacing']}pt, "
                + ", ".join(icon_markup)
                + ")"
            )

        left = "[" + "".join(left_parts) + "]"
        right = (
            f"text(size: {cfg['person_summary_font_size']}pt)"
            f"[#v({cfg['person_summary_offset']}pt){summary}]"
        )
        return self.pad(f"#cv-entry({left}, {right}, {self._columns('person')})")

    def make_affiliation_block(self, affiliation: Dict[str, Any]) -> str:
        cfg = self.config
        organization = affiliation.get("organization", "")
        department = affiliation.get("department", "")
        body = ", ".join([p for p in (organization, department) if p])

        return "\n".join(
            [
                self._title_date_grid(
                    escape(affiliation.get("role", "")),
                    affiliation.get("date_range", ""),
                    "affiliation",
                ),
                f"#v({cfg['item_line_gap']}pt)",
                f"#text(size: {cfg['item_body_font_size']}pt)[{escape(body)}]",
            ]
        )

    def _work_body(self, work: Dict[str, Any]) -> str:
        """Renders the journal / link / author line(s) of a publication."""
        parts = [escape(work.get("journal", ""))]

        link = work.get("link")
        if link:
            label = f"{link['prefix']}#underline[{escape(link['label'])} ]"
            parts.append(f"#link({raw_str(link['url'])})[{label}]")

        if work.get("subtitle"):
            parts.append(escape(work["subtitle"]))

        first_line = ", ".join(p for p in parts if p)

        authors: List[Author] = work.get("authors", [])
        author_str = join_authors(
            authors,
            bold=lambda name: f"#strong[{escape(name)}]",
            plain=escape,
        )
        if author_str:
            return f"{first_line}#linebreak(){author_str}"
        return first_line

    def make_work_block(self, work: Dict[str, Any]) -> str:
        cfg = self.config
        return "\n".join(
            [
                self._title_date_grid(
                    escape(work.get("title", "")), work.get("date", ""), "work"
                ),
                f"#v({cfg['item_line_gap']}pt)",
                f"#text(size: {cfg['item_body_font_size']}pt)[{self._work_body(work)}]",
            ]
        )

    def make_funding_block(self, fund: Dict[str, Any]) -> str:
        cfg = self.config
        body = f"{fund.get('org', '')}, {fund.get('id', '')}"
        left = f"text(size: {cfg['item_body_font_size']}pt)[{escape(body)}]"
        right = f"text(size: {cfg['item_misc_font_size']}pt)[{escape(fund.get('role', ''))}]"

        return "\n".join(
            [
                self._title_date_grid(
                    escape(fund.get("title", "")),
                    fund.get("start_year", ""),
                    "affiliation",
                ),
                f"#v({cfg['item_line_gap']}pt)",
                f"#cv-entry({left}, {right}, {self._columns('affiliation')})",
            ]
        )

    def make_review_block(
        self, row: Tuple[Tuple[str, int], Optional[Tuple[str, int]]]
    ) -> str:
        cfg = self.config
        first, second = row
        left_text = format_review(*first)
        right_text = format_review(*second) if second else ""

        size = cfg["item_body_font_size"]
        left = f"text(size: {size}pt)[{escape(left_text)}]"
        right = f"text(size: {size}pt)[{escape(right_text)}]"
        return (
            f"#grid(columns: {self._columns('review')}, "
            f"align: (left + top, left + top), {left}, {right})"
        )
