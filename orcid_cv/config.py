from typing import Dict, Any

# Backends able to turn a style config into a PDF
BACKENDS = ("reportlab", "typst")


def make_document_config(style: str, backend: str = "reportlab") -> Dict[str, Any]:
    """
    Returns a style configuration dictionary for the specified style name and
    rendering backend ('reportlab' or 'typst').
    Raises ValueError if the style or backend is invalid.
    """
    backend = backend.lower()
    if backend not in BACKENDS:
        raise ValueError(f"Invalid backend: {backend}. Choose one of {BACKENDS}.")

    if backend == "typst":
        return _make_typst_config(style)
    return _make_reportlab_config(style)


def _make_reportlab_config(style: str) -> Dict[str, Any]:
    """Builds the config for the reportlab backend (holds ParagraphStyle objects)."""
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.enums import TA_LEFT, TA_RIGHT
    from reportlab.lib.styles import ParagraphStyle

    config: Dict[str, Any] = {}
    if style.lower() == "greenspon-default":
        config = {
            "style": style.lower(),
            "backend": "reportlab",
            "pagesize": letter,
            "margin": 40,
            "item_spacing": 5,
            "page_footer": False,
            "initalize_authors": True,
            "embolden_author": True,
            "initalize_primary_author": True,
            "person_title_style": ParagraphStyle(
                "PersonTitle", alignment=TA_LEFT, fontSize=22, fontName="Helvetica-Bold"
            ),
            "person_summary_style": ParagraphStyle(
                "PersonSummary", alignment=TA_RIGHT, fontSize=9, fontName="Helvetica"
            ),
            "section_style": ParagraphStyle(
                "SectionTitle",
                alignment=TA_LEFT,
                fontSize=18,
                fontName="Helvetica-Bold",
            ),
            "item_title_style": ParagraphStyle(
                "ItemTitle", alignment=TA_LEFT, fontSize=11, fontName="Helvetica-Bold"
            ),
            "item_date_style": ParagraphStyle(
                "ItemDate", alignment=TA_RIGHT, fontSize=9, fontName="Helvetica-Bold"
            ),
            "item_misc_style": ParagraphStyle(
                "ItemMisc", alignment=TA_RIGHT, fontSize=9, fontName="Helvetica"
            ),
            "item_body_style": ParagraphStyle(
                "ItemBody",
                alignment=TA_LEFT,
                fontSize=9,
                fontName="Helvetica",
                underlineWidth=1,
                underlineOffset="-0.1*F",
            ),
        }
        from orcid_cv.styles import GreensponDefaultRenderer

        config["renderer"] = GreensponDefaultRenderer(config)
    else:
        raise ValueError(f"Invalid style: {style}")

    return config


def _make_typst_config(style: str) -> Dict[str, Any]:
    """
    Builds the config for the typst backend. Unlike the reportlab config this is
    pure data: font sizes in points, font family preferences and spacing, which
    the typst renderers interpolate into markup.
    """
    config: Dict[str, Any] = {}
    if style.lower() == "greenspon-default":
        config = {
            "style": style.lower(),
            "backend": "typst",
            "paper": "us-letter",
            "margin": 40,
            "bottom_margin": 40,
            # Every gap in the document is explicit: par spacing is switched off so
            # these values are the only thing driving the vertical rhythm.
            "par_spacing": 0,
            "line_leading": 5.5,
            "item_spacing": 15,
            "review_row_spacing": 6,
            "section_spacing": 14,
            "cell_padding": 6,
            "page_footer": False,
            "footer_date": "",
            "initalize_authors": True,
            "embolden_author": True,
            "initalize_primary_author": True,
            # Helvetica is not installed on most machines; typst walks this list
            # and takes the first family it can find.
            "font_family": ("Helvetica", "Arial", "Liberation Sans", "Nimbus Sans"),
            "body_font_size": 9,
            "person_title_font_size": 22,
            "person_summary_font_size": 9,
            "person_summary_offset": 12,
            "section_font_size": 18,
            "item_title_font_size": 11,
            "item_date_font_size": 9,
            "item_misc_font_size": 9,
            "item_body_font_size": 9,
            "item_line_gap": 9,
            "heading_rule_gap": 18,
            "heading_rule_width": 2,
            "heading_body_gap": 3,
            "icon_size": 15,
            "icon_gap": 4,
            "icon_spacing": 5,
            "footer_font_size": 9,
            "footer_rule_width": 0.5,
            "footer_rule_gap": 10,
            "footer_gap": 2,
            "footer_descent": 20,
            "typst_assets": {},
        }
        from orcid_cv.typst_styles import GreensponDefaultTypstRenderer

        config["renderer"] = GreensponDefaultTypstRenderer(config)
    else:
        raise ValueError(f"Invalid style: {style}")

    return config
