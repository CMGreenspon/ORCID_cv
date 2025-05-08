#%% Load
import orcid_cv as ocv
orcid_dir = r"C:\Users\Somlab\Downloads\0000-0002-6806-3302"
orcid_dict = ocv.extract_orcid_info(orcid_dir)


#%% Custom modifications
# Add co-first/last authors
ocv.add_equal_author(orcid_dict['work']['104077035']['authors'], 2)
ocv.add_equal_author(orcid_dict['work']['117624833']['authors'], 2)
ocv.add_equal_author(orcid_dict['work']['146346630']['authors'], 3)
ocv.add_equal_author(orcid_dict['work']['181828994']['authors'], 0, 2)
ocv.add_equal_author(orcid_dict['work']['163633584']['authors'], 0, 2)

# Add R01 application
orcid_dict['funding']['12345'] = {'title': 'Improving artificial tactile feedback using volumetric intracortical microstimulation',
                                              'role': 'PI',
                                              'org': 'National Institutes of Health - Scored (29th percentile)',
                                              'start_year': '2025',
                                              'end_year': '',
                                              'value': ''}

orcid_dict['work']['181828994']['journal'] = 'eLife'
#%% Export
style = 'greenspon-default'
config = ocv.make_document_config(style)
output_fname = r"C:\Users\somlab\OneDrive - The University of Chicago\Miscellaneous\CMG_CV.pdf"
# output_fname = r"C:\Users\Somlab\Downloads\test.pdf"
doc_title = orcid_dict['personal']['fullname'] + ' - CV'
doc = ocv.SimpleDocTemplate(output_fname,
                            pagesize = ocv.letter,
                            leftMargin = config['margin'],
                            rightMargin = config['margin'],
                            topMargin = config['margin'],
                            bottomMargin = config['margin'] + 10,
                            title = doc_title)
elements = []
ocv.add_person_section(elements, orcid_dict, config)
ocv.add_affiliation_section(elements, orcid_dict, config, 'Employment', 'employment')
ocv.add_affiliation_section(elements, orcid_dict, config, 'Education', 'education')
ocv.add_work_section(elements, orcid_dict, config, 'Research Publications', ['journal-article', 'preprint'])
ocv.add_work_section(elements, orcid_dict, config, 'Talks', 'public-speech')
ocv.add_funding_section(elements, orcid_dict, config, 'Funding')
# ocv.add_review_section(elements, orcid_dict, config, 'Peer Review')
ocv.add_work_section(elements, orcid_dict, config, 'Book Chapters', 'book-chapter')
# ocv.add_work_section(elements, orcid_dict, config, 'Software', 'software')
# doc.build(elements)
doc.multiBuild(elements, canvasmaker=ocv.FooterCanvas)
