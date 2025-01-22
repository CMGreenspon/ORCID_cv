import os
import xmltodict
import json
import requests
from copy import deepcopy
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import Paragraph, Spacer, Table, SimpleDocTemplate, Image
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from datetime import datetime
package_directory = os.path.dirname(os.path.abspath(__file__))


class HyperlinkedImage(Image, object):
    # https://stackoverflow.com/questions/18114820/is-it-possible-to-get-a-flowables-coordinate-position-once-its-rendered-using
    def __init__(self, filename, hyperlink=None, width=None, height=None, kind='direct', mask='auto', lazy=1):
        super(HyperlinkedImage, self).__init__(filename, width, height, kind, mask, lazy)
        self.hyperlink = hyperlink

    def drawOn(self, canvas, x, y, _sW=0):
        if self.hyperlink:  # If a hyperlink is given, create a canvas.linkURL()
            x1 = x
            y1 = y
            x2 = x1 + self._width
            y2 = y1 + self._height
            canvas.linkURL(url=self.hyperlink, rect=(x1, y1, x2, y2), thickness=0, relative=1)
        super(HyperlinkedImage, self).drawOn(canvas, x, y, _sW)


def load_xml(xml_path):
    with open(xml_path, encoding='utf-8') as xd:
        xml_dict = xmltodict.parse(xd.read())
    
    # Remove top_level
    if len(xml_dict.keys()) == 1:
        return list(xml_dict.values())[0]
    else:
        raise Exception('XML dictionary has more than one top-level key')


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
    # Extract common info
    affiliation_dict = {'organization': get_recursive_key(affiliation_xml, 'common:organization', 'common:name'),
                        'department': get_recursive_key(affiliation_xml, 'common:department-name'),
                        'role': get_recursive_key(affiliation_xml, 'common:role-title'),
                        'start_date': get_recursive_key(affiliation_xml, 'common:start-date', 'common:year'),
                        'end_date': get_recursive_key(affiliation_xml, 'common:end-date', 'common:year')}
    # Make date str
    if affiliation_dict['end_date'] == '':
        affiliation_dict['date_range'] = affiliation_dict['start_date'] + ' - present'
    else:
        affiliation_dict['date_range'] = affiliation_dict['start_date'] + ' - ' + affiliation_dict['end_date']

    return affiliation_dict


def load_work(work_path):
    # Convert .xml to dictionary. Generically loads fields, styles determine later display
    in_work_dict = load_xml(work_path)
    out_work_dict = {'type': in_work_dict['work:type'],  # This field is used for sorting entries and defines other behavior
                     'title': get_recursive_key(in_work_dict, 'work:title', 'common:title'),
                     'subtitle': '',
                     'journal': get_recursive_key(in_work_dict, 'work:journal-title'),
                     'doi': get_recursive_key(in_work_dict, 'common:url'),
                     'year': get_recursive_key(in_work_dict, 'common:publication-date', 'common:year'),
                     'month': get_recursive_key(in_work_dict, 'common:publication-date', 'common:month'),
                     'authors': get_recursive_key(in_work_dict, 'work:contributors', 'work:contributor')}
    # Check empty date for later sorting
    if out_work_dict['year'] == '':
        out_work_dict['year'] = 0
    if out_work_dict['month'] == '':
        out_work_dict['month'] = 0
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
                if 'rxiv' in out_work_dict['journal']:
                    out_work_dict['journal'] = out_work_dict['journal'].replace('rxiv', 'Rxiv')
            except:
                print('Could not lookup preprint: ' + out_work_dict['title'])
    elif out_work_dict['type'] == 'software':
        out_work_dict['subtitle'] = get_recursive_key(in_work_dict, 'work:title', 'common:subtitle')
    return out_work_dict


def load_funding(funding_path):
    in_funding_dict = load_xml(funding_path)
    out_funding_dict = {'title': get_recursive_key(in_funding_dict, 'funding:title', 'common:title'),
                        'role': get_recursive_key(in_funding_dict, 'funding:organization-defined-type'),
                        'org': get_recursive_key(in_funding_dict, 'common:organization', 'common:name'),
                        'start_year': get_recursive_key(in_funding_dict, 'common:start-date', 'common:year'),
                        'end_year': get_recursive_key(in_funding_dict, 'common:end-date', 'common:year'),
                        'value': get_recursive_key(in_funding_dict, 'funding:amount', '#text')}
    
    return out_funding_dict
    
def load_review(review_path):
    in_review_dict = load_xml(review_path)
    # Parse IISN
    issn = in_review_dict['peer-review:review-group-id'][5:]
    r = requests.get(f"https://portal.issn.org/resource/ISSN/{issn}")
    potential_name = ''
    for l in r.iter_lines():
        line = str(l)
        if issn in line:
            potential_name = line
            idx = potential_name.find('|')
            potential_name = potential_name[idx+2:-5]
            break
    if potential_name == '':
        print(f"Could not identify ISSN {issn}")

    # Make dict
    out_review_dict = {'year': get_recursive_key(in_review_dict, 'peer-review:review-completion-date', 'common:year'),
                       'role': get_recursive_key(in_review_dict, 'peer-review:review-type'),
                       'org': potential_name.title()}
    
    return out_review_dict

def extract_orcid_info(orcid_dir):
    # First check if there is an ORCID.json file
    if os.path.isfile(os.path.join(orcid_dir, 'ORCID.json')):
        print('Loading ORCID dict from local json.')
        with open(os.path.join(orcid_dir, 'ORCID.json')) as f:
            return json.load(f)
    
    # Personal info
    personal_info = load_xml(os.path.join(orcid_dir, "person.xml"))
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

    # Parse XML folders to make dictionaries
    employment_dict = folder_to_dict(os.path.join(orcid_dir, 'affiliations', 'employments'), load_affiliation)
    education_dict = folder_to_dict(os.path.join(orcid_dir, 'affiliations', 'educations'), load_affiliation)
    work_dict = folder_to_dict(os.path.join(orcid_dir, 'works'), load_work)
    funding_dict = folder_to_dict(os.path.join(orcid_dir, 'fundings'), load_funding)
    review_dict = folder_to_dict(os.path.join(orcid_dir, 'peer_reviews'), load_review)

    # Combine
    out_dict = {'personal': personal,
                'work': work_dict,
                'employment': employment_dict,  
                'education': education_dict,
                'funding': funding_dict,
                'reviews': review_dict}

    # Write the json
    print('Saving local json.')
    with open(os.path.join(orcid_dir, 'ORCID.json'), 'w') as fp:
        json.dump(out_dict, fp, indent = 4)

    return out_dict


def folder_to_dict(path, load_fun: callable) -> dict:
    _dict = {}
    xml_list = os.listdir(path)
    for x in xml_list:
        _dict[x[:-4]] = load_fun(os.path.join(path, x))
    
    return _dict


def dict_to_list(input_dict: dict) -> list:
    return [input_dict[k] for k in input_dict.keys()]


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
            if person['fullname'] in a:
                embolden = True
            elif person['name-short'] in a:
                embolden = True
            elif person['firstname'] + ' ' + person['lastname'] in a:
                embolden = True
            elif person['firstname'][0] + '. ' + person['lastname'] in a:
                embolden = True
            else:
                print('Did not embolden: ' + a)

            if embolden:
                author_list[i] = ''.join(['<b>', a, '</b>'])

    return author_list


def add_equal_author(author_list: list[str], num_first: int = 0, num_last: int = 0):
    num_authors = len(author_list)
    for i in range(1, len(author_list) + 1):
        if i <= num_first:
            author_list[i-1] += '*'
        if i > num_authors-num_last:
            author_list[i-1] += '*'


def get_column_widths(config, section_type):
    ratio = 1
    table_width = int(config['pagesize'][0] - (config['margin'] * 2))
    if config['style'] == 'greenspon-default':
        if section_type == 'work':
            ratio = 7
        elif section_type == 'affiliation':
            ratio = 6
        elif section_type == 'person':
            ratio = 3.5
        else:
            ValueError('what')
    else:
        ValueError('Invalid style')

    right_col_width = round(table_width / ratio)
    return [table_width - right_col_width, right_col_width]


def make_affiliation_table(config, affiliation, section_heading = ''):
    if config['style'] == 'greenspon-default':
        if section_heading == '':
            table_data = [[Paragraph(affiliation['role'], style = config['item_title_style']), Paragraph(affiliation['date_range'], style = config['item_date_style'])],
                          [Paragraph(affiliation['organization'] + ', ' + affiliation['department'], style = config['item_body_style']), '']]
            table_style = [('VALIGN', (0, 0), (-1, -1), 'TOP'),
                           ('NOSPLIT', (0, 0), (-1, -1))]  # ('GRID', (0,0), (-1, -1), 0.5, colors.gray)
        else:
            table_data = [[Paragraph(section_heading, style = config['section_style']), ''],
                          ['', ''],  # Padding for large config['section_style']
                          [Paragraph(affiliation['role'], style = config['item_title_style']), Paragraph(affiliation['date_range'], style = config['item_date_style'])],
                          [Paragraph(affiliation['organization'] + ', ' + affiliation['department'], style = config['item_body_style']), '']]
            table_style = [('SPAN', (0, 0), (-1, 0)),
                           ('LINEBELOW', (0, 1), (-1, 1), 2, colors.gray),
                           ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                           ('NOSPLIT', (0, 0), (-1, -1))]
    else:
        ValueError('Invalid style')

    return table_data, table_style


def make_work_table(config, work_title, work_body, work_date, section_heading = ''):
    # Show as empty
    if work_date == 0:
        work_date = ''

    if config['style'] == 'greenspon-default':
        if section_heading == '':
            table_data = [[Paragraph(work_title, style = config['item_title_style']), Paragraph(work_date, style = config['item_date_style'])],
                          [work_body, '']]
            table_style = [('VALIGN', (0, 0), (-1, -1), 'TOP'),
                           ('NOSPLIT', (0, 0), (-1, -1))]  # ('GRID', (0,0), (-1, -1), 0.5, colors.gray)
        else:
            table_data = [[Paragraph(section_heading, style = config['section_style']), ''],
                          ['', ''],  # Padding for large config['section_style']
                          [Paragraph(work_title, style = config['item_title_style']), Paragraph(work_date, style = config['item_date_style'])],
                          [work_body, '']]
            table_style = [('SPAN', (0, 0), (-1, 0)),
                           ('LINEBELOW', (0, 1), (-1, 1), 2, colors.gray),
                           ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                           ('NOSPLIT', (0, 0), (-1, -1))]
    else:
        ValueError('Invalid style')

    return table_data, table_style


def make_funding_table(config, fund, section_heading = ''):
    if config['style'] == 'greenspon-default':
        if section_heading == '':
            table_data = [[Paragraph(fund['title'], style = config['item_title_style']), Paragraph(fund['start_year'], style = config['item_date_style'])],
                          [fund['org'], '']]
            table_style = [('VALIGN', (0, 0), (-1, -1), 'TOP'),
                           ('NOSPLIT', (0, 0), (-1, -1))]  # ('GRID', (0,0), (-1, -1), 0.5, colors.gray)
        else:
            table_data = [[Paragraph(section_heading, style = config['section_style']), ''],
                          ['', ''],  # Padding for large config['section_style']
                          [Paragraph(fund['title'], style = config['item_title_style']), Paragraph(fund['start_year'], style = config['item_date_style'])],
                          [fund['org'], '']]
            table_style = [('SPAN', (0, 0), (-1, 0)),
                           ('LINEBELOW', (0, 1), (-1, 1), 2, colors.gray),
                           ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                           ('NOSPLIT', (0, 0), (-1, -1))]
    else:
        ValueError('Invalid style')

    return table_data, table_style


def process_external_links(link_dict):
    link_list = []
    for (k, v) in link_dict.items():
        im_path = os.path.join(package_directory, 'external_link_img', k + '.png')
        link_list.append(HyperlinkedImage(im_path, hyperlink = v, height = 15, width = 15))
    return link_list


def add_person_section(elements, orcid_dict, config):
    if config['style'] == 'greenspon-default':
        # Format table
        column_widths = get_column_widths(config, 'person')
        info = dict_to_list(orcid_dict['employment'])
        info = sorted(info, key = lambda v: int(v['start_date']), reverse = True)
        fullname = orcid_dict['personal']['fullname']
        person_summary = ('<br/>' +
                          info[0]['role'] + '<br/>' +
                          info[0]['organization'] + '<br/>' +
                          orcid_dict['personal']['email'])
        table_data = [[Paragraph(fullname, style = config['person_title_style']), Paragraph(person_summary, style = config['person_summary_style'])]]
        table_style = [('NOSPLIT', (0, 0), (-1, -1)),
                       ('VALIGN', (0, 0), (-1, -1), 'TOP')]
        # Convert to table
        t = Table(table_data, colWidths = column_widths)
        t.setStyle(table_style)
        # Append
        elements.append(t)
        elements.append(Spacer(0, -15))

        # Add links
        link_list = process_external_links(orcid_dict['personal']['links'])
        t = Table([link_list], colWidths = [20] * 1, hAlign='LEFT')
        t.setStyle(table_style)
        elements.append(t)
        elements.append(Spacer(0, config['item_spacing']))
    else:
        ValueError('Nope')


def add_affiliation_section(elements, orcid_dict: dict, config, heading, affiliation_type):
    # Compute column size
    column_widths = get_column_widths(config, 'affiliation')

    # Get correct affiliation type
    if not orcid_dict.__contains__(affiliation_type):
        return ValueError('Dict does not contain affiliation type: ' + affiliation_type)
    
    # Sort the affiliations
    affiliations = dict_to_list(orcid_dict[affiliation_type])
    affiliations = sorted(affiliations, key = lambda v: int(v['start_date']), reverse = True)
    
    # Iterate through affiliations and make tables
    is_heading = True
    for af in affiliations:
        # prepare table
        if is_heading:
            table_data, table_style = make_affiliation_table(config, af, section_heading = heading)
            is_heading = False
        else:
            table_data, table_style = make_affiliation_table(config, af)

        # Convert to table
        t = Table(table_data, colWidths = column_widths)
        t.setStyle(table_style)
        # Append
        elements.append(t)
        elements.append(Spacer(0, config['item_spacing']))


def add_work_section(elements, orcid_dict, config, heading, search_str):
    # Compute column size
    column_widths = get_column_widths(config, 'work')
    if isinstance(search_str, str):
        search_str = [search_str]

    # Get subset of publications
    works = dict_to_list(orcid_dict['work'])
    works = [w for w in works if w['type'] in search_str]
    if works == []:
        return ValueError('No matching works for: ' + search_str)
    works = sorted(works, key = lambda v: int(v['year']) * 1000 + int(v['month']), reverse = True)  # Sort by year then month

    # Iterate through list and make tables
    is_heading = True
    for work in works:
        work_date = work['year']
        work_journal = work['journal']
        # Process title
        work_title = work['title']
        if '‐' in work_title:  # Replace bad character
            work_title = work_title.replace('‐', '-')
        # Swap journal and date for software
        if work['type'] == 'software':
            work_journal = work['subtitle']
            work_date = work['journal']
            work['subtitle'] = '' # Prevent duplications from custom subtitle fields

        # Process author
        author_cat = ''
        if not work['authors'] == []:
            author_list = work['authors']
            if config['initalize_authors']:
                author_list = [initalize_name(i) for i in author_list]
            if config['embolden_author']:
                author_list = embolden_authors(orcid_dict['personal'], author_list)
            if len(author_list) > 1:
                author_list[-1] = 'and ' + author_list[-1]
            if len(author_list) > 2:
                author_cat = ', '.join(author_list)
        if '‐' in author_cat:  # Replace bad character
            author_cat = author_cat.replace('‐', '-')

        # Process DOI/link
        doi_str = work['doi']
        if doi_str == '':  # Remove "," and go straight to author line
            work_str = work_journal
        else:
            if 'doi.org/' in doi_str:
                idx = doi_str.find('doi.org/')
                short_doi = doi_str[idx + 8:]
                doi_str = '<link href="' + 'https://www.doi.org/' + short_doi + '">' + 'DOI: <u>' + short_doi + ' </u></link>'
            elif 'github.com/' in doi_str:
                idx = doi_str.find('github.com/')
                short_doi = doi_str[idx + 11:]
                doi_str = '<link href="' + 'https://www.github.com/' + short_doi + '">' + 'GitHub: <u>' + short_doi + ' </u></link>'
            else:
                doi_str = '<link href="' + doi_str + '"><u>' + doi_str + ' </u></link>'
            work_str = work_journal + ', ' + doi_str

        # Add subtitle
        if work['subtitle'] != '':
            work_str = work_str + ', ' + work['subtitle']
        
        # Add formatting to work str
        work_body = Paragraph(work_str + '<br/>' + author_cat, style = config['item_body_style'])
        # prepare table
        if is_heading:
            table_data, table_style = make_work_table(config, work_title, work_body, work_date, section_heading = heading)
            is_heading = False
        else:
            table_data, table_style = make_work_table(config, work_title, work_body, work_date, section_heading = '')

        # Convert to table
        t = Table(table_data, colWidths = column_widths)
        t.setStyle(table_style)
        # Append
        elements.append(t)
        elements.append(Spacer(0, config['item_spacing']))


def add_funding_section(elements, orcid_dict, config, heading):
    # Compute column size
    column_widths = get_column_widths(config, 'affiliation')
    
    # Sort the fund
    fund = dict_to_list(orcid_dict['funding'])
    fund = sorted(fund, key = lambda v: int(v['start_year']), reverse = True)
    
    # Iterate through fund and make tables
    is_heading = True
    for f in fund:
        # prepare table
        if is_heading:
            table_data, table_style = make_funding_table(config, f, section_heading = heading)
            is_heading = False
        else:
            table_data, table_style = make_funding_table(config, f)

        # Convert to table
        t = Table(table_data, colWidths = column_widths)
        t.setStyle(table_style)
        # Append
        elements.append(t)
        elements.append(Spacer(0, config['item_spacing']))


class FooterCanvas(canvas.Canvas):
    left_str: str = ""

    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self.pages = []

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self.pages)
        for page in self.pages:
            self.__dict__.update(page)
            self.draw_canvas(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_canvas(self, page_count):
        page = "Page %s of %s" % (self._pageNumber, page_count)
        x = 100
        self.saveState()
        self.setStrokeColorRGB(0, 0, 0)
        self.setLineWidth(0.5)
        self.line(50, 78, letter[0] - 50, 78)
        self.setFont('Helvetica', 9)
        self.drawString(letter[0]-x, 65, page)
        self.drawString(50, 65, datetime.today().strftime("%d-%b-%Y"))
        self.restoreState()


def make_document_config(style):
    if style.lower() == 'greenspon-default':
        # Enable TTF fonts
        # pdfmetrics.registerFont(TTFont('GillSans', 'GIL_____.TTF'))
        # pdfmetrics.registerFont(TTFont('GillSansBold', 'GILB____.TTF'))
        # pdfmetrics.registerFontFamily('GillSans', normal = 'GillSans', bold = 'GillSansBold')
        config = {'style': style.lower(),
                  'pagesize': letter,
                  'margin': 50,
                  'item_spacing': 5,
                  'initalize_authors': True,
                  'embolden_author': True,
                  'initalize_primary_author': True,
                  'person_title_style': ParagraphStyle('PersonTitle', alignment = TA_LEFT, fontSize = 22, fontName = 'Helvetica-Bold'),
                  'person_summary_style': ParagraphStyle('PersonSummary', alignment = TA_RIGHT, fontSize = 9, fontName = 'Helvetica'),
                  'section_style': ParagraphStyle('SectionTitle', alignment = TA_LEFT, fontSize = 18, fontName = 'Helvetica-Bold'),
                  'item_title_style': ParagraphStyle('ItemTitle', alignment = TA_LEFT, fontSize = 11, fontName = 'Helvetica-Bold'),
                  'item_date_style': ParagraphStyle('ItemDate', alignment = TA_RIGHT, fontSize = 9, fontName = 'Helvetica-Bold'),
                  'item_body_style': ParagraphStyle('ItemBody', alignment = TA_LEFT, fontSize = 9, fontName = 'Helvetica', underlineWidth = 1,
                                                    underlineOffset = '-0.1*F')}
    else:
        ValueError('Invalid style')

    return config


def quick_build(orcid_dir, output_fname, style = 'greenspon-default'):
    orcid_dict = extract_orcid_info(orcid_dir)
    config = make_document_config(style)
    doc_title = orcid_dict['personal']['fullname'] + ' - CV'
    doc = SimpleDocTemplate(output_fname,
                            pagesize = letter,
                            leftMargin = config['margin'],
                            rightMargin = config['margin'],
                            topMargin = config['margin'],
                            bottomMargin = config['margin'],
                            title = doc_title)
    elements = []
    add_person_section(elements, orcid_dict, config)
    add_affiliation_section(elements, orcid_dict, config, 'Employment', 'employment')
    add_affiliation_section(elements, orcid_dict, config, 'Education', 'education')
    add_work_section(elements, orcid_dict, config, 'Research Publications', 'journal-article')
    add_work_section(elements, orcid_dict, config, 'Talks', 'lecture-speech')
    add_work_section(elements, orcid_dict, config, 'Preprints', 'preprint')
    doc.build(elements)
    print('Success!')
