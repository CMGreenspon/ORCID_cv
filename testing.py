# -*- coding: utf-8 -*-
"""
Created on Thu Dec 28 16:54:05 2023

@author: Somlab
"""

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
import os
import xmltodict

# Load XML files
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
    try:
        if in_work_dict['work:work']['work:type'] == 'journal-article':
            out_work_dict = {'title': in_work_dict['work:work']['work:title']['common:title'],
                             'journal': in_work_dict['work:work']['work:journal-title'],
                             'type': in_work_dict['work:work']['work:type'],
                             'doi': in_work_dict['work:work']['common:url'],
                             'year': in_work_dict['work:work']['common:publication-date']['common:year'],
                             'authors': [i['work:credit-name'] for i in in_work_dict['work:work']['work:contributors']['work:contributor']]}
        elif in_work_dict['work:work']['work:type'] == 'preprint':
            out_work_dict = {'title': in_work_dict['work:work']['work:title']['common:title'],
                             'journal': 'preprint',
                             'type': in_work_dict['work:work']['work:type'],
                             'doi': in_work_dict['work:work']['common:url'],
                             'year': in_work_dict['work:work']['common:publication-date']['common:year'],
                             'authors': [i['work:credit-name'] for i in in_work_dict['work:work']['work:contributors']['work:contributor']]}
        elif in_work_dict['work:work']['work:type'] == 'book-chapter':
            out_work_dict = {'title': in_work_dict['work:work']['work:title']['common:subtitle'],
                             'journal': in_work_dict['work:work']['work:title']['common:title'],
                             'type': in_work_dict['work:work']['work:type'],
                             'doi': in_work_dict['work:work']['common:url'],
                             'year': in_work_dict['work:work']['common:publication-date']['common:year'],
                             'authors': [i['work:credit-name'] for i in in_work_dict['work:work']['work:contributors']['work:contributor']]}
        else:
             print(in_work_dict['work:work']['work:type'] + ' is not recognized')
             continue
        work_list.append(out_work_dict)
    except:
        print('Issue parsing work: ' + in_work_dict['work:work']['work:title']['common:title'] + ' (' + in_work_dict['work:work']['@put-code'] + ')')
    
# Enable TTF fonts
pdfmetrics.registerFont(TTFont('GillSans', 'GIL_____.ttf'))
pdfmetrics.registerFont(TTFont('GillSansBold', 'GILB____.ttf'))

# Section spacing
canvas.setLineWidth(1)

# Create document
canvas = canvas.Canvas("form.pdf", pagesize=letter)
canvas.setFont('GillSans', 11)
canvas.drawString(30,750,'Charles M. Greenspon')
canvas.save()


