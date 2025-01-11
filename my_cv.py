#%% Load
import orcid_cv as ocv
orcid_dir = r"C:\Users\Somlab\Downloads\0000-0002-6806-3302"
orcid_dict = ocv.extract_orcid_info(orcid_dir)


#%% Custom modifications


#%% Export
style = 'greenspon-default'
config = ocv.make_document_config(style)
# output_fname = r"C:\Users\somlab\OneDrive - The University of Chicago\Miscellaneous\CMG_CV.pdf"
output_fname = r"C:\Users\Somlab\Downloads\test.pdf"
doc_title = orcid_dict['personal']['fullname'] + ' - CV'
doc = ocv.SimpleDocTemplate(output_fname,
                            pagesize = ocv.letter,
                            leftMargin = config['margin'],
                            rightMargin = config['margin'],
                            topMargin = config['margin'],
                            bottomMargin = config['margin'],
                            title = doc_title)
elements = []
ocv.add_person_section(elements, orcid_dict, config)
ocv.add_affiliation_section(elements, orcid_dict, config, 'Employment', 'employment')
ocv.add_affiliation_section(elements, orcid_dict, config, 'Education', 'education')
ocv.add_work_section(elements, orcid_dict, config, 'Research Publications', ['journal-article', 'preprint'])
ocv.add_work_section(elements, orcid_dict, config, 'Talks', 'lecture-speech')
# ocv.add_funding_section(elements, orcid_dict, config, 'Funding')
ocv.add_review_section(elements, orcid_dict, config, 'Peer Review')
ocv.add_work_section(elements, orcid_dict, config, 'Book Chapters', 'book-chapter')
ocv.add_work_section(elements, orcid_dict, config, 'Software', 'software')
doc.build(elements)
