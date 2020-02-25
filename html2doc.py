import os
import re
from bs4 import BeautifulSoup
from pathlib import Path
from urllib.request import urlopen
from docx import Document
from docx.enum.text import WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor
from docx.shared import Pt
from docx.shared import Cm


def sorted_alphanumeric(data):
    def convert(text): return int(text) if text.isdigit() else text.lower()
    def alphanum(key): return [convert(c) for c in re.split('([0-9]+)', key)]
    return sorted(data, key=alphanum)


directory = 'D:/reading/yechen/raw/'

document = Document()
firstPagebreak = True  # set flag for pagebreak

section = document.sections[0]
section.left_margin = Cm(2.25)
section.top_margin = Cm(2.0)

styles = document.styles
new_heading_style = styles.add_style('New Heading', WD_STYLE_TYPE.PARAGRAPH)
new_heading_style.base_style = styles['Heading 1']
font = new_heading_style.font
font.name = 'Calibri Light'
font.size = Pt(12)
font.bold = False
font.color.rgb = RGBColor(0, 0, 0)

for filename in sorted_alphanumeric(os.listdir(directory)):
    if filename.endswith('.html'):
        extName = os.path.join(directory, filename)

        # if extracting from local html
        path = 'file:///'+extName
        filename = Path(path).stem
        html = urlopen(path)

        # if scrape from web
        # html = urlopen('https://www.biqupa.com/7_7817/6864689.html')

        content = html.read().decode('gb2312', 'ignore')
        soup = BeautifulSoup(content, 'lxml')

        if(firstPagebreak):
            firstPagebreak = False
        else:
            pagebreak = document.add_paragraph().add_run()
            pagebreak.add_break(WD_BREAK.PAGE)

        # if scrape from web
        # text = soup.find('div', attrs={'id':'content'}).get_text()
        # test = text.split('章')

        # if scrape from web
        # document.add_paragraph(test[0].strip()+'章',  style='New Heading')

        # if extracting from local html
        heading = document.add_paragraph(filename, style='New Heading')
        heading.paragraph_format.space_before = Pt(0)

        paragraph = document.add_paragraph()
        paragraph.paragraph_format.line_spacing = 1.5

        # if scrape from web
        # run = paragraph.add_run(test[1])

        # if extracting from local html
        run = paragraph.add_run(soup.get_text(separator='\n'))

        font = run.font
        font.name = 'Arial'
        font.size = Pt(10)

document.save('D:\\reading\\yechen\\raw\\combined.docx')
