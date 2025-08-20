#%% Load
import orcid_cv as ocv
orcid_dir = r"C:\Users\Somlab\Downloads\0000-0002-6806-3302"
orcid_dict = ocv.extract_orcid_info(orcid_dir)

#%% Custom modifications
# Add co-first/last authors
ocv.add_equal_author(orcid_dict['work']['104077035']['authors'], 2)
ocv.add_equal_author(orcid_dict['work']['117624833']['authors'], 2)
ocv.add_equal_author(orcid_dict['work']['146346630']['authors'], 3)
ocv.add_equal_author(orcid_dict['work']['184869731']['authors'], 0, 2)
ocv.add_equal_author(orcid_dict['work']['184953056']['authors'], 0, 2)
ocv.add_equal_author(orcid_dict['work']['189765308']['authors'], 2)

# Add pending applications
orcid_dict['funding']['123456'] = {'title': 'Reshaping encoding and decoding algorithms for bidirectional brain-computer interfaces',
                                            'role': 'CoI',
                                            'org': 'National Institutes of Health',
                                            'id': 'R01 (Pending)',
                                            'start_year': '2025',
                                            'end_year': '',
                                            'value': ''}
orcid_dict['funding']['123466'] = {'title': 'Improving artificial tactile feedback using volumetric intracortical microstimulation',
                                            'role': 'PI',
                                            'org': 'National Institutes of Health',
                                            'id': 'DP2 (Pending)',
                                            'start_year': '2025',
                                            'end_year': '',
                                            'value': ''}

# Add review that don't get added to ORCID
orcid_dict['reviews']['1'] = {'org': 'Cerebral Cortex'}
orcid_dict['reviews']['2'] = {'org': 'Journal of Neural Engineering'}
orcid_dict['reviews']['3'] = {'org': 'Journal of Neural Engineering'}
orcid_dict['reviews']['4'] = {'org': 'Journal of Neural Engineering'}

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
ocv.add_work_section(elements, orcid_dict, config, 'Conference Presentations', ['conference-presentation'])
ocv.add_work_section(elements, orcid_dict, config, 'Invited Talks', 'public-speech')
ocv.add_funding_section(elements, orcid_dict, config, 'Funding')
ocv.add_review_section(elements, orcid_dict, config, 'Peer Review')
ocv.add_work_section(elements, orcid_dict, config, 'Book Chapters', 'book-chapter')
# ocv.add_work_section(elements, orcid_dict, config, 'Software', 'software')
# doc.build(elements)
doc.multiBuild(elements, canvasmaker=ocv.FooterCanvas)
