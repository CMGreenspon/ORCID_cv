# %% Quick build (reportlab, the default)
import orcid_cv as ocv

orcid_dir = r"C:\Users\somlab\Downloads\0000-0002-6806-3302"
output_fname = r"C:\Users\somlab\Downloads\test_cv_quickbuild.pdf"
ocv.quick_build(orcid_dir, output_fname)

# %% Quick build with typst instead of reportlab
import orcid_cv as ocv

orcid_dir = r"C:\Users\somlab\Downloads\0000-0002-6806-3302"
output_fname = r"C:\Users\somlab\Downloads\test_cv_quickbuild_typst.pdf"
ocv.quick_build(orcid_dir, output_fname, backend="typst")

# %% Custom example - swap 'backend' to switch PDF engines, everything else is shared
import orcid_cv as ocv

orcid_dir = r"C:\Users\Somlab\Downloads\0000-0002-6806-3302"
output_fname = (
    r"C:\Users\somlab\OneDrive - The University of Chicago\Miscellaneous\CMG_CV.pdf"
)
orcid_dict = ocv.extract_orcid_info(orcid_dir)
style = "greenspon-default"
backend = "typst"  # or 'reportlab'
config = ocv.make_document_config(style, backend=backend)
config["page_footer"] = True  # 'Page X of Y' + today's date at the bottom
doc_title = orcid_dict["personal"]["fullname"] + " - CV"
elements = []
ocv.add_person_section(elements, orcid_dict, config)
ocv.add_affiliation_section(elements, orcid_dict, config, "Employment", "employment")
ocv.add_affiliation_section(elements, orcid_dict, config, "Education", "education")
ocv.add_work_section(
    elements,
    orcid_dict,
    config,
    "Research Publications",
    ["journal-article", "preprint"],
)
ocv.add_work_section(elements, orcid_dict, config, "Talks", "lecture-speech")
ocv.add_work_section(elements, orcid_dict, config, "Book Chapters", "book-chapter")
ocv.add_work_section(elements, orcid_dict, config, "Software", "software")
# ORCID keeps mentorship and service in one folder; match/exclude on the role
# title splits them, or drop both arguments for a single combined section
ocv.add_service_section(elements, orcid_dict, config, "Mentorship", match="Advisor")
ocv.add_service_section(elements, orcid_dict, config, "Service", exclude="Advisor")
ocv.build_document(
    output_fname,
    elements,
    config,
    title=doc_title,
    author=orcid_dict["personal"]["fullname"],
)

# %% Keep the generated typst markup to tweak or compile by hand
import orcid_cv as ocv

orcid_dir = r"C:\Users\Somlab\Downloads\0000-0002-6806-3302"
orcid_dict = ocv.extract_orcid_info(orcid_dir)
config = ocv.make_document_config("greenspon-default", backend="typst")
elements = []
ocv.add_person_section(elements, orcid_dict, config)
ocv.add_work_section(
    elements, orcid_dict, config, "Research Publications", "journal-article"
)
ocv.build_document(
    r"C:\Users\somlab\Downloads\cv.pdf",
    elements,
    config,
    save_source=r"C:\Users\somlab\Downloads\cv.typ",
)
