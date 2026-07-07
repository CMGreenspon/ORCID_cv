from reportlab.lib.pagesizes import letter
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from typing import Dict, Any


def make_document_config(style: str) -> Dict[str, Any]:
    """
    Returns a style configuration dictionary for the specified style name.
    Raises ValueError if the style is invalid.
    """
    config: Dict[str, Any] = {}
    if style.lower() == "greenspon-default":
        config = {
            "style": style.lower(),
            "pagesize": letter,
            "margin": 40,
            "item_spacing": 5,
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
