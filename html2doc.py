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


document = Document()
firstPagebreak = True  # set flag for pagebreak

# formatting for whole document
section = document.sections[0]
section.left_margin = Cm(2.25)
section.top_margin = Cm(2.0)

# formatting for heading of the chapter
styles = document.styles
new_heading_style = styles.add_style('New Heading', WD_STYLE_TYPE.PARAGRAPH)
new_heading_style.base_style = styles['Heading 1']
font = new_heading_style.font
font.name = 'Calibri Light'
font.size = Pt(12)
font.bold = False
font.color.rgb = RGBColor(0, 0, 0)

directory = 'D:/reading/yechen/raw/'

for filename in sorted_alphanumeric(os.listdir(directory)):
    if filename.endswith('.html'):
        extName = os.path.join(directory, filename)

        # if extracting from local html
        path = 'file:///' + extName
        filename = Path(path).stem
        html = urlopen(path)
        content = html.read().decode('gb2312', 'ignore')

        # clean up html in preparations for processing
        regex1 = '^(.*?)<div id=\"content\" class=\"showtxt\">'
        regex2 = '(?<=</div>).*'
        regex3 = '^(.*?)章'
        regex4 = '<br />\r<br />&nbsp;'
        replace_txt4 = '<br />&nbsp;'

        cleaned_txt = re.sub(regex1, '', content, flags=re.DOTALL)
        cleaned_txt = re.sub(regex2, '', cleaned_txt, flags=re.DOTALL)
        cleaned_txt = re.sub(regex3, '', cleaned_txt, flags=re.DOTALL)
        cleaned_txt = cleaned_txt.replace(regex4, replace_txt4)

        # remove nonsense words
        useless_words = [
            "更新最快",
            "手机端一秒住槟提供精彩\\小fx。",
            "首发",
            "78中文首发",
            "叶辰萧初然来源：",
            "叶辰萧初然"
            ]

        for index, word in enumerate(useless_words):
            cleaned_txt = cleaned_txt.replace(useless_words[index], "")

        soup = BeautifulSoup(cleaned_txt, 'lxml')

        if(firstPagebreak):
            firstPagebreak = False
        else:
            pagebreak = document.add_paragraph().add_run()
            pagebreak.add_break(WD_BREAK.PAGE)

        # Add heading to the top of the page, using filename of HTML
        heading = document.add_paragraph(filename, style='New Heading')

        # formatting for content
        heading.paragraph_format.space_before = Pt(0)
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.line_spacing = 1.5

        # if extracting from local html
        # add the cleaned up content into word doc
        run = paragraph.add_run(soup.get_text(separator='\n'))
        run.font.name = 'Arial'
        run.font.size = Pt(10)

document.save(directory + 'combined.docx')
