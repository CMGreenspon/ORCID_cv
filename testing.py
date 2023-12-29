import os
import xmltodict
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import Paragraph, Spacer, Image, Table, SimpleDocTemplate, Drawing
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.graphics.shapes import Drawing, Line
from reportlab.lib import colors
from math import ceil

# Enable TTF fonts
pdfmetrics.registerFont(TTFont('GillSans', 'GIL_____.ttf'))
pdfmetrics.registerFont(TTFont('GillSansBold', 'GILB____.ttf'))

# Document config
output_dir = r'C:\GitHub\misc\ORCIDResume' + r'\test_resume.pdf'
margin = 50 # proportion of page size

#%% Load XML files
orcid_dir = r"C:\Users\Somlab\Downloads\0000-0002-6806-3302"

# Personal info
with open(os.path.join(orcid_dir, "person.xml")) as fd:
    personal_info = xmltodict.parse(fd.read())
# Extract name
personal_info['name'] = (personal_info['person:person']['person:name']['personal-details:given-names'] + " " +
                         personal_info['person:person']['person:name']['personal-details:family-name'])
# Extract links
personal_info['links'] = {'ORCID': 'https://orcid.org/' + personal_info['person:person']['person:name']['@path']}
if 'researcher-url:researcher-urls' in personal_info['person:person'].keys():
    # Assign so it's not insanely long
    link_list = personal_info['person:person']['researcher-url:researcher-urls']['researcher-url:researcher-url']
    for i in range(len(link_list)):
        personal_info['links'][link_list[i]['researcher-url:url-name']] = link_list[i]['researcher-url:url']
    
# Works
work_list = []
work_xml_list = os.listdir(os.path.join(orcid_dir, 'works'))
for i in work_xml_list:
    with open(os.path.join(orcid_dir, 'works', i), encoding='utf-8') as fd:
        in_work_dict = xmltodict.parse(fd.read())
    # Extract relevant info
    in_work_dict = in_work_dict['work:work']
    try:
        if in_work_dict['work:type'] == 'journal-article':
            out_work_dict = {'title': in_work_dict['work:title']['common:title'],
                             'journal': in_work_dict['work:journal-title'],
                             'type': in_work_dict['work:type'],
                             'doi': in_work_dict['common:url'],
                             'year': in_work_dict['common:publication-date']['common:year'],
                             'month': in_work_dict['common:publication-date']['common:month'],
                             'authors': [i['work:credit-name'] for i in in_work_dict['work:contributors']['work:contributor']]}
        elif in_work_dict['work:type'] == 'preprint':
            out_work_dict = {'title': in_work_dict['work:title']['common:title'],
                             'journal': 'preprint',
                             'type': in_work_dict['work:type'],
                             'doi': in_work_dict['common:url'],
                             'year': in_work_dict['common:publication-date']['common:year'],
                             'month': in_work_dict['common:publication-date']['common:month'],
                             'authors': [i['work:credit-name'] for i in in_work_dict['work:contributors']['work:contributor']]}
        elif in_work_dict['work:type'] == 'book-chapter':
            out_work_dict = {'title': in_work_dict['work:title']['common:subtitle'],
                             'journal': in_work_dict['work:title']['common:title'],
                             'type': in_work_dict['work:type'],
                             'doi': in_work_dict['common:url'],
                             'year': in_work_dict['common:publication-date']['common:year'],
                             'month': in_work_dict['common:publication-date']['common:month'],
                             'authors': [i['work:credit-name'] for i in in_work_dict['work:contributors']['work:contributor']]}
        else:
             print(in_work_dict['work:type'] + ' is not recognized')
             continue
        work_list.append(out_work_dict)
    except:
        print('Issue parsing work: ' + in_work_dict['work:title']['common:title'] + ' (' + in_work_dict['@put-code'] + ')')
    
#%% Create document
doc = SimpleDocTemplate(output_dir,
                        pagesize = letter,
                        leftMargin = margin,
                        rightMargin = margin,
                        topMargin = margin,
                        bottomMargin = margin)

ncols = 6
column_width  = round((doc.pagesize[0] - (margin * 2)) / ncols)
column_widths = [column_width] * ncols
row_height = 20
elements = []

sectionStyle = ParagraphStyle('Section', alignment = TA_LEFT, fontSize = 20, fontName = 'Helvetica-Bold')
itemTitleStyle = ParagraphStyle('Section', alignment = TA_LEFT, fontSize = 11, fontName = 'Helvetica-Bold')
itemDateStyle = ParagraphStyle('Section', alignment = TA_RIGHT, fontSize = 11, fontName = 'Helvetica-Bold')
itemBodyStyle = ParagraphStyle('Section', alignment = TA_LEFT, fontSize = 11, fontName = 'Helvetica')

header_string = 'Research Publications'
elements.append(Paragraph(header_string, style = sectionStyle))
elements.append(Spacer(0, 10))
d = Drawing(100, 1)
d.add(Line(0,0, doc.pagesize[0] - (margin * 2),0))
elements.append(d)
elements.append(Spacer(0, 5))

works = [i for i in work_list if i['type'] == 'journal-article']
for i in range(len(works)):
    work = works[i]
    author_cat = ', '.join(work['authors'])
    data = [[Paragraph(work['title'], style = itemTitleStyle), '', '', '', '', Paragraph(work['year'], style = itemDateStyle)],
            [Paragraph(work['journal'] + ', ' + work['doi'], style = itemBodyStyle), '', '', '', '', ''],
            [Paragraph(author_cat, style = itemBodyStyle), '', '', '', '', '']]
    style = [('SPAN', (0,0), (ncols-2,0)),
             ('SPAN', (0,1), (ncols-2,1)),
             ('SPAN', (0,2), (ncols-2,2))]
    t = Table(data, colWidths = column_widths) #, rowHeights = [row_height * ceil(len(work['title']) / 85), row_height, row_height * ceil(len(author_cat)/ 85)]
    t.setStyle(style)
    # Append
    elements.append(t)
    elements.append(Spacer(0, 5))
# Build
doc.build(elements)
