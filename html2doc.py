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

directory = 'D:/reading/yechen/raw/'

for filename in sorted_alphanumeric(os.listdir(directory)):
    if filename.endswith('.html'):
        extName = os.path.join(directory, filename)

        # if extracting from local html
        path = 'file:///'+extName
        filename = Path(path).stem
        html = urlopen(path)

        # if scrape from web
        # html = urlopen('https://www.biqupa.com/7_7817/6864689.html')
        # html = urlopen('http://blog.weimengclass.com/index.php/2020/02/25/%e3%80%8a%e4%b8%8a%e9%97%a8%e9%be%99%e5%a9%bf%e3%80%8b1051-1100/')

        content = html.read().decode('gb2312', 'ignore')
        # clean up before adding into word doc
        regex1 = r"^(.*?)<div id=\"content\" class=\"showtxt\">"
        regex2 = r'(?<=</div>).*'
        regex3 = r'^(.*?)章'
        regex4 = '<br />\r<br />&nbsp;'
        replace_txt4 = '<br />&nbsp;'

        cleaned_txt = re.sub(regex1, '', content, flags=re.DOTALL).strip()
        cleaned_txt = re.sub(regex2, '', cleaned_txt, flags=re.DOTALL).strip()
        cleaned_txt = re.sub(regex3, '', cleaned_txt, flags=re.DOTALL).strip()
        cleaned_txt = cleaned_txt.replace(regex4, replace_txt4)

        # remove nonsense words
        useless_words = ["更新最快", "手机端一秒住槟提供精彩\\小fx。"]
        for index, word in enumerate(useless_words):
            cleaned_txt = cleaned_txt.replace(useless_words[index], "")

        soup = BeautifulSoup(cleaned_txt, 'lxml')

        if(firstPagebreak):
            firstPagebreak = False
        else:
            pagebreak = document.add_paragraph().add_run()
            pagebreak.add_break(WD_BREAK.PAGE)

        # if scrape from web or local html from http://blog.weimengclass.com/
        # unfinished code - to change to individual methods

        # text = soup.find('div', attrs={'class': 'entry-content'}).get_text()
        # content_split = re.split(r'第(\d+)章', text)
        # print(content_split[39])

        # for x in range(len(content_split)):
            # print(x)

        # header = content_split[0].strip()+'章'
        # heading = document.add_paragraph(header, style='New Heading')

        # if extracting from local html
        heading = document.add_paragraph(filename, style='New Heading')

        # formatting
        heading.paragraph_format.space_before = Pt(0)

        paragraph = document.add_paragraph()
        paragraph.paragraph_format.line_spacing = 1.5

        # if scrape from web
        # run = paragraph.add_run(content_split[1])

        # if extracting from local html
        # add into word doc
        run = paragraph.add_run(soup.get_text(separator='\n'))
        # run = paragraph.add_run(cleaned_txt)

        font = run.font
        font.name = 'Arial'
        font.size = Pt(10)

document.save('D:\\reading\\yechen\\raw\\combined.docx')
