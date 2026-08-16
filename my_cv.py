# %% Load
import orcid_cv as ocv

orcid_dir = r"C:\Users\Somlab\Downloads\0000-0002-6806-3302"
orcid_dict = ocv.extract_orcid_info(orcid_dir)

# %% Custom modifications
# Add co-first/last authors
ocv.add_equal_author(orcid_dict["work"]["104077035"]["authors"], 2)
ocv.add_equal_author(orcid_dict["work"]["117624833"]["authors"], 2)
ocv.add_equal_author(orcid_dict["work"]["146346630"]["authors"], 3)
ocv.add_equal_author(orcid_dict["work"]["217401609"]["authors"], 0, 2)
ocv.add_equal_author(orcid_dict["work"]["184953056"]["authors"], 0, 2)
ocv.add_equal_author(orcid_dict["work"]["189765308"]["authors"], 2)

# Add pending applications
orcid_dict["funding"]["123456"] = {
    "title": "Reshaping encoding and decoding algorithms for bidirectional brain-computer interfaces",
    "role": "CoI",
    "org": "National Institutes of Health",
    "id": "R01 (Pending)",
    "start_year": "2025",
    "end_year": "",
    "value": "",
}

# Add review that don't get added to ORCID
orcid_dict["reviews"]["1"] = {"org": "Cerebral Cortex"}

# Replace some names
for k, v in orcid_dict["reviews"].items():
    if "Proceedings Of The National Academy Of Sciences" in v["org"]:
        v["org"] = "PNAS"


# %% Export
style = "greenspon-default"
backend = "reportlab"  # or "typst"
config = ocv.make_document_config(style, backend=backend)
config["page_footer"] = True
output_fname = (
    r"C:\Users\somlab\OneDrive - The University of Chicago\Miscellaneous\CMG_CV.pdf"
)
# output_fname = r"C:\Users\Somlab\Downloads\test.pdf"
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
ocv.add_work_section(
    elements,
    orcid_dict,
    config,
    "Conference Presentations",
    ["conference-presentation"],
)
ocv.add_work_section(elements, orcid_dict, config, "Invited Talks", "public-speech")
ocv.add_funding_section(elements, orcid_dict, config, "Funding")
ocv.add_review_section(elements, orcid_dict, config, "Peer Review")
ocv.add_work_section(elements, orcid_dict, config, "Book Chapters", "book-chapter")
# ocv.add_work_section(elements, orcid_dict, config, 'Software', 'software')
ocv.build_document(
    output_fname,
    elements,
    config,
    title=doc_title,
    author=orcid_dict["personal"]["fullname"],
)

# %% Papers only
# Add stars to some titles
wlist = [
    "104077035",
    "169596753",
    "173243281",
    "183248945",
    "184953056",
    "189765308",
    "50152425",
]
for w in wlist:
    if orcid_dict["work"][w]["title"][0] == "*":
        continue
    orcid_dict["work"][w]["title"] = "*" + orcid_dict["work"][w]["title"]


style = "greenspon-default"
config = ocv.make_document_config(style, backend=backend)
config["page_footer"] = True
output_fname = r"C:\Users\Somlab\Downloads\CMG_Papers.pdf"
doc_title = orcid_dict["personal"]["fullname"] + " - CV"
elements = []
ocv.add_work_section(
    elements,
    orcid_dict,
    config,
    "Research Publications",
    ["journal-article", "preprint"],
)
ocv.build_document(
    output_fname,
    elements,
    config,
    title=doc_title,
    author=orcid_dict["personal"]["fullname"],
)
