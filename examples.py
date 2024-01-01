import orcid_cv as ocv

#%% Quick build
orcid_dir = r"C:\Users\somlab\Downloads\0000-0002-6806-3302"
output_fname = r"C:\Users\somlab\Downloads\test_cv.pdf"
ocv.quick_build(orcid_dir, output_fname)

#%% Custom example
orcid_dir = r"C:\Users\Somlab\Downloads\0000-0002-6806-3302"
output_dir = r'C:\Users\Somlab\Downloads\test_resume.pdf'
orcid_dict = ocv.extract_orcid_info(orcid_dir)
style = 'greenspon-default'
config = ocv.make_document_config(style)
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
ocv.add_work_section(elements, orcid_dict, config, 'Research Publications', 'journal-article')
ocv.add_work_section(elements, orcid_dict, config, 'Talks', 'lecture-speech')
ocv.add_work_section(elements, orcid_dict, config, 'Preprints', 'preprint')
ocv.add_work_section(elements, orcid_dict, config, 'Book Chapters', 'book-chapter')
ocv.add_work_section(elements, orcid_dict, config, 'Software', 'software')
doc.build(elements)
