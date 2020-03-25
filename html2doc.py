import os
import re
import win32com.client
from bs4 import BeautifulSoup
from urllib.request import urlopen
from docx import Document
from docx.enum.text import WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor
from docx.shared import Pt
from docx.shared import Cm
# from pathlib import Path


def sorted_alphanumeric(data):
    def convert(text): return int(text) if text.isdigit() else text.lower()
    def alphanum(key): return [convert(c) for c in re.split('([0-9]+)', key)]
    return sorted(data, key=alphanum)


def getHTML(url):
    mainHtml = urlopen(url)
    mainContent = mainHtml.read().decode('gb2312', 'ignore')
    return mainContent


def getLatestChapter(mainContent):
    latestChpt = re.search(r'最新章节：(.*?)</p>', mainContent).group(1)
    chptNum = re.search(r'>第(.*?)章<', latestChpt).group(1)
    return chptNum


def getChapters(dir, url):
    mainContent = getHTML(url)
    latest = getLatestChapter(mainContent)
    doneTill = 1
    links = {}

    for filename in sorted_alphanumeric(os.listdir(dir)):
        if filename.endswith('.completed'):
            doneTill = os.path.splitext(filename)[0]

    latest = int(latest)
    doneTill = int(doneTill)
    iterations = latest - doneTill
    for x in range(iterations):
        doneTill += 1
        regex = r'<dd><a href ="(.*?)">第' + str(doneTill)
        chptUrl = re.search(regex, mainContent).group(1)
        links[doneTill] = chptUrl
    return links


def updateLatestChapter(directory, latest):
    for filename in os.listdir(directory):
        if filename.endswith('.completed'):
            full_filename = os.path.join(directory, filename)
            full_latestChapter = os.path.join(directory, str(latest))
            os.rename(full_filename, full_latestChapter)


def cleanHTML(content):
    regex1 = '^(.*?)<div id=\"content\" class=\"showtxt\">'
    regex2 = '(?<=</div>).*'
    regex3 = '^(.*?)章'
    regex4 = '<br />\r<br />&nbsp;'
    replace_txt4 = '<br />&nbsp;'

    cleaned_txt = re.sub(regex1, '', content, flags=re.DOTALL)
    cleaned_txt = re.sub(regex2, '', cleaned_txt, flags=re.DOTALL)
    cleaned_txt = re.sub(regex3, '', cleaned_txt, flags=re.DOTALL)
    cleaned_txt = cleaned_txt.replace(regex4, replace_txt4)
    return cleaned_txt


def removeNonsense(content):
    useless_words = [
        "更新最快",
        "手机端一秒住槟提供精彩\\小fx。",
        "78中文首发",
        "叶辰萧初然来源：",
        "叶辰萧初然",
        "首发"
    ]

    for index, word in enumerate(useless_words):
        cleaned_txt = content.replace(useless_words[index], "")

    return cleaned_txt


def update_toc(docx_file):
    word = win32com.client.DispatchEx("Word.Application")
    doc = word.Documents.Open(docx_file)
    doc.TablesOfContents(1).Update()
    doc.Close(SaveChanges=True)
    word.Quit()


# init directory and website location
directory = 'D:/reading/yechen/'
base_url = 'https://www.biqupa.com'
topic = '7_7817'
url = '%s/%s/' % (base_url, topic)
combinedDocName = 'YeChenXiaoChuRanXiaoShuo.docx'
fullpath_combineDName = directory + combinedDocName

# init document creation and start formatting
document = Document(fullpath_combineDName)
firstPagebreak = True  # set flag for pagebreak

# formatting for whole document
section = document.sections[0]
section.left_margin = Cm(2.25)
section.top_margin = Cm(0)

section = document.sections[1]
section.left_margin = Cm(2.25)
section.top_margin = Cm(2.0)

# formatting for heading of the chapter
chptValid = False

list_styles = [s for s in document.styles if s.type == WD_STYLE_TYPE.PARAGRAPH]
for style in list_styles:
    if(style.name == 'chpt'):
        chptValid = True

if (chptValid is False):
    styles = document.styles
    new_heading_style = styles.add_style('chpt', WD_STYLE_TYPE.PARAGRAPH)
    new_heading_style.base_style = styles['Heading 1']
    font = new_heading_style.font
    font.name = 'Calibri Light'
    font.size = Pt(12)
    font.bold = False
    font.color.rgb = RGBColor(0, 0, 0)

# check webpage for latest chapter and get array of links
chapters = getChapters(directory, url)

for chapter in chapters:
    chapterURL = '%s/%s' % (base_url, chapters[chapter])
    html = urlopen(chapterURL)
    content = html.read().decode('gb2312', 'ignore')

    # clean up html in preparations for processing
    cleaned_txt = cleanHTML(content)

    # remove nonsense words
    cleaned_txt = removeNonsense(cleaned_txt)

    soup = BeautifulSoup(cleaned_txt, 'lxml')

    pagebreak = document.add_paragraph().add_run()
    pagebreak.add_break(WD_BREAK.PAGE)

    # Add heading to the top of the page
    heading = document.add_paragraph("第"+str(chapter)+"章", style='chpt')

    # formatting for content
    heading.paragraph_format.space_before = Pt(0)
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.line_spacing = 1.5

    # if extracting from local html
    # add the cleaned up content into word doc
    run = paragraph.add_run(soup.get_text(separator='\n'))
    run.font.name = 'Arial'
    run.font.size = Pt(10)

    # update to latest after completion
    updateLatestChapter(directory, str(chapter) + ".completed")

"""
# local extraction
for filename in sorted_alphanumeric(os.listdir(directory)):
    if filename.endswith('.html'):
        extName = os.path.join(directory, filename)

        # if extracting from local html
        path = 'file:///' + extName
        filename = Path(path).stem
        html = urlopen(path)
        content = html.read().decode('gb2312', 'ignore')

        # clean up html in preparations for processing
        cleaned_txt = cleanHTML(content)

        # remove nonsense words
        cleaned_txt = removeNonsense(cleaned_txt)

        soup = BeautifulSoup(cleaned_txt, 'lxml')

        if(firstPagebreak):
            firstPagebreak = False
        else:
            pagebreak = document.add_paragraph().add_run()
            pagebreak.add_break(WD_BREAK.PAGE)

        # Add heading to the top of the page, using filename of HTML
        heading = document.add_paragraph(getCurrChapter(content), style='chpt')

        # formatting for content
        heading.paragraph_format.space_before = Pt(0)
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.line_spacing = 1.5

        # if extracting from local html
        # add the cleaned up content into word doc
        run = paragraph.add_run(soup.get_text(separator='\n'))
        run.font.name = 'Arial'
        run.font.size = Pt(10)
"""

if(bool(chapters)):
    document.save(fullpath_combineDName)
    update_toc(fullpath_combineDName)
