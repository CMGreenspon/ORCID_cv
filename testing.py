import os
import xmltodict
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import Paragraph, Spacer, Table, SimpleDocTemplate  # Image
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.lib.styles import ParagraphStyle
from reportlab.graphics.shapes import Drawing, Rect
from reportlab.lib import colors
import requests

# Enable TTF fonts
pdfmetrics.registerFont(TTFont('GillSans', 'GIL_____.ttf'))
pdfmetrics.registerFont(TTFont('GillSansBold', 'GILB____.ttf'))
pdfmetrics.registerFontFamily('GillSans', normal = 'GillSans', bold = 'GillSansBold')

# Document config
margin = 50  # proportion of page size
embolden_author = True
initalize_authors = True
style = 'greenspon-default'

# Declare fonts
sectionStyle = ParagraphStyle('Section', alignment = TA_LEFT, fontSize = 20, fontName = 'GillSansBold')
itemTitleStyle = ParagraphStyle('Section', alignment = TA_LEFT, fontSize = 11, fontName = 'GillSansBold')
itemDateStyle = ParagraphStyle('Section', alignment = TA_RIGHT, fontSize = 9, fontName = 'GillSansBold')
itemBodyStyle = ParagraphStyle('Section', alignment = TA_LEFT, fontSize = 9, fontName = 'GillSans', underlineWidth = 1,
                               underlineOffset = '-0.1*F')


#%% Functions
def load_xml(xml_path):
    with open(xml_path, encoding='utf-8') as xd:
        xml_dict = xmltodict.parse(xd.read())
    return xml_dict


def list_works(orcid_dir):
    work_xml_list = os.listdir(os.path.join(orcid_dir, 'works'))
    for (i, w) in enumerate(work_xml_list):
        xml_dict = load_xml(os.path.join(orcid_dir, 'works', w))
        print(str(i) + ': ' + xml_dict['work:work']['work:title']['common:title'] + ' (' + xml_dict['work:work']['@put-code'] + ')')


def get_recursive_key(input_dict, *keys):
    if not isinstance(input_dict, dict):
        raise TypeError('first argument must be a dict.')

    if len(keys) == 0 or not all([isinstance(k, str) for k in keys]):
        raise TypeError('Keys must all be of type "str".')

    _dict = input_dict
    for key in keys:
        if _dict.__contains__(key) and not _dict[key] is None:
            _dict = _dict[key]
        else:
            print('Could not find: ' + '-'.join(keys) + ' in item #' + input_dict['@put-code'])
            return ''
    return _dict


def load_affiliation(affiliation_path):
    # Load and enter top level xml
    affiliation_xml = load_xml(affiliation_path)
    if affiliation_xml.__contains__('education:education'):
        affiliation_xml = affiliation_xml['education:education']
    elif affiliation_xml.__contains__('employment:employment'):
        affiliation_xml = affiliation_xml['employment:employment']
    else:
        ValueError('Unable to read affiliation .xml')

    # Extract common info
    affiliation_dict = {'organization': get_recursive_key(affiliation_xml, 'common:organization', 'common:name'),
                        'department': get_recursive_key(affiliation_xml, 'common:department-name'),
                        'role': get_recursive_key(affiliation_xml, 'common:role-title')}
    # Make date str
    sd = get_recursive_key(affiliation_xml, 'common:start-date', 'common:year')
    ed = get_recursive_key(affiliation_xml, 'common:end-date', 'common:year')
    if ed == '':
        affiliation_dict['date'] = sd + ' - present'
    else:
        affiliation_dict['date'] = sd + ' - ' + ed

    return affiliation_dict


def load_work(work_path):
    # Convert .xml to dictionary. Generically loads fields, styles determine later display
    in_work_dict = load_xml(work_path)
    in_work_dict = in_work_dict['work:work']  # Everything is in this one field so just skip to it
    out_work_dict = {'type': in_work_dict['work:type'],  # This field is used for sorting entries and defines other behavior
                     'title': get_recursive_key(in_work_dict, 'work:title', 'common:title'),
                     'journal': get_recursive_key(in_work_dict, 'work:journal-title'),
                     'doi': get_recursive_key(in_work_dict, 'common:url'),
                     'year': get_recursive_key(in_work_dict, 'common:publication-date', 'common:year'),
                     'month': get_recursive_key(in_work_dict, 'common:publication-date', 'common:month'),
                     'authors': get_recursive_key(in_work_dict, 'work:contributors', 'work:contributor')}
    # Extract authors from list
    if not out_work_dict['authors'] == '':
        if isinstance(out_work_dict['authors'], dict):
            out_work_dict['authors'] = [out_work_dict['authors']['work:credit-name']]
        elif isinstance(out_work_dict['authors'], list):
            out_work_dict['authors'] = [i['work:credit-name'] for i in out_work_dict['authors']]

    # Journal / repository
    if out_work_dict['type'] == 'preprint':
        if not out_work_dict['doi'] == '':
            try:
                print('Trying to find host repository for article: ' + out_work_dict['title'])
                doi_data = requests.get(out_work_dict['doi'])
                url = doi_data.url
                start_idx = url.find('www.') + 4
                slash_idx = url.find('/', start_idx)
                url = url[start_idx:slash_idx]
                out_work_dict['journal'] = url[:url.rfind('.')]
            except:
                print('Could not lookup preprint: ' + out_work_dict['title'])

    return out_work_dict


def extract_orcid_info(orcid_dir):
    # Personal info
    personal_info = load_xml(os.path.join(orcid_dir, "person.xml"))
    # Extract name
    personal_info = personal_info['person:person']
    personal = {'lastname': get_recursive_key(personal_info, 'person:name', 'personal-details:family-name'),
                'givenname': get_recursive_key(personal_info, 'person:name', 'personal-details:given-names'),
                'links': {'ORCID': 'https://orcid.org/' + get_recursive_key(personal_info, 'person:name', '@path')}}
    # Workout other elements of name
    personal['fullname'] = personal['givenname'] + ' ' + personal['lastname']
    personal['name-short'] = initalize_name(personal['fullname'])
    personal['firstname'] = personal['fullname'][0:personal['fullname'].find(' ')]
    # Extract links
    if 'researcher-url:researcher-urls' in personal_info.keys():
        # Assign so it's not insanely long
        link_list = get_recursive_key(personal_info, 'researcher-url:researcher-urls', 'researcher-url:researcher-url')
        for i in range(len(link_list)):
            personal['links'][link_list[i]['researcher-url:url-name']] = link_list[i]['researcher-url:url']
    # Get primary email
    personal['email'] = [e['email:email'] for e in personal_info['email:emails']['email:email'] if e['@primary'] == 'true'][0]

    # Employment
    employment_list = []
    affiliation_xml_list = os.listdir(os.path.join(orcid_dir, 'affiliations', 'employments'))
    for i in affiliation_xml_list:
        employment_list.append(load_affiliation(os.path.join(orcid_dir, 'affiliations', 'employments', i)))

    # Education
    education_list = []
    education_xml_list = os.listdir(os.path.join(orcid_dir, 'affiliations', 'educations'))
    for i in education_xml_list:
        education_list.append(load_affiliation(os.path.join(orcid_dir, 'affiliations', 'educations', i)))

    # Works
    work_list = []
    work_xml_list = os.listdir(os.path.join(orcid_dir, 'works'))
    for i in work_xml_list:
        out_work_dict = load_work(os.path.join(orcid_dir, 'works', i))
        work_list.append(out_work_dict)

    return {'personal': personal, 'work': work_list, 'employment': employment_list, 'education': education_list}


def initalize_name(input_str):
    str_split = input_str.split(' ')
    for i in range(len(str_split) - 1):
        if '.' in str_split[i] and str_split[i][1] == '.':
            continue
        else:
            str_split[i] = str_split[i][0] + '.'
    output_str = ''.join([*str_split[:-1], ' ', str_split[-1]])
    return output_str


def embolden_authors(person, author_list):
    for (i, a) in enumerate(author_list):
        if person['lastname'] in a:
            embolden = False
            if a == person['fullname']:
                embolden = True
            elif a == person['name-short']:
                embolden = True
            elif a == person['firstname'] + ' ' + person['lastname']:
                embolden = True
            elif a == person['firstname'][0] + '. ' + person['lastname']:
                embolden = True
            else:
                print('Did not embolden: ' + a)

            if embolden:
                author_list[i] = ''.join(['<b>', a, '</b>'])

    return author_list


def make_work_table(style, work_title, work_body, work_date, section_heading = ''):
    if style.lower() == 'greenspon-default':
        if section_heading == '':
            table_data = [[Paragraph(work_title, style = itemTitleStyle), Paragraph(work_date, style = itemDateStyle)],
                          [work_body, '']]
            table_style = [('VALIGN', (0, 0), (-1, -1), 'TOP'),
                           ('NOSPLIT', (0, 0), (-1, -1))]  # ('GRID', (0,0), (-1, -1), 0.5, colors.gray)
        else:
            table_data = [[Paragraph(section_heading, style = sectionStyle), ''],
                          ['', ''],  # Padding for large sectionStyle
                          ['', ''],
                          [Paragraph(work_title, style = itemTitleStyle), Paragraph(work_date, style = itemDateStyle)],
                          [work_body, '']]
            table_style = [('SPAN', (0, 0), (-1, 0)),
                           ('LINEBELOW', (0, 1), (-1, 1), 2, colors.gray),
                           ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                           ('NOSPLIT', (0, 0), (-1, -1))]
    else:
        ValueError('Invalid style')

    return table_data, table_style


def add_work_section(elements, orcid_dict, heading, search_str):
    # Compute column size
    table_width = int(doc.pagesize[0] - (margin * 2))
    right_col_width = round(table_width / 7)
    column_widths = [table_width - right_col_width, right_col_width]

    # Get subset of publications
    works = [i for i in orcid_dict['work'] if i['type'] == search_str]
    works = sorted(works, key = lambda v: int(v['year']) * 1000 + int(v['month']), reverse = True)  # Sort by year then month

    for (i, work) in enumerate(works):
        # work = works[w]
        # Process title
        work_title = work['title']
        if '‐' in work_title:  # Replace bad character
            work_title = work_title.replace('‐', '-')
        # Process author
        if not work['authors'] == []:
            author_list = work['authors']
            if initalize_authors:
                author_list = [initalize_name(i) for i in author_list]
            if embolden_author:
                author_list = embolden_authors(orcid_dict['personal'], author_list)
            if len(author_list) > 1:
                author_list[-1] = 'and ' + author_list[-1]
            author_cat = ', '.join(author_list)
        else:
            author_cat = ''

        # Process DOI
        doi_str = work['doi']
        if 'doi.org/' in doi_str:
            idx = doi_str.find('doi.org/')
            short_doi = doi_str[idx + 8:]
            doi_str = '<link href="' + 'https://www.hdoi.org/' + short_doi + '">' + 'DOI: <u>' + short_doi + ' </u></link>'
            work_body = Paragraph(work['journal'] + ', ' + doi_str + '<br/>' + author_cat, style = itemBodyStyle)
        else:  # Remove "," and go straight to author line
            work_body = Paragraph(work['journal'] + '<br/>' + author_cat, style = itemBodyStyle)

        # prepare table
        if i == 0:
            table_data, table_style = make_work_table(style, work_title, work_body, work['year'], section_heading = heading)
        else:
            table_data, table_style = make_work_table(style, work_title, work_body, work['year'], section_heading = '')

        # Convert to table
        t = Table(table_data, colWidths = column_widths)
        t.setStyle(table_style)
        # Append
        elements.append(t)
        elements.append(Spacer(0, 10))


#%% Build document
orcid_dir = r"C:\Users\Somlab\Downloads\0000-0002-6806-3302"
orcid_info = extract_orcid_info(orcid_dir)
#%%
output_dir = r'C:\Users\Somlab\Downloads\test_resume.pdf'
doc = SimpleDocTemplate(output_dir,
                        pagesize = letter,
                        leftMargin = margin,
                        rightMargin = margin,
                        topMargin = margin,
                        bottomMargin = margin)
elements = []
add_work_section(elements, orcid_info, 'Research Publications', 'journal-article')
add_work_section(elements, orcid_info, 'Talks', 'lecture-speech')
add_work_section(elements, orcid_info, 'Preprints', 'preprint')
doc.build(elements)
