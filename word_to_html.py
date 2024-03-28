import zipfile
import xml.etree.ElementTree as ET
import re
import string


# Path to the DOCX file
docx_file_path = 'test.docx'

with zipfile.ZipFile(docx_file_path, 'r') as zip_ref:
    with zip_ref.open('word/document.xml') as xml_file:
        # Read the XML content
        xml_content = xml_file.read()

root = ET.fromstring(xml_content)

paragraphs = []
for paragraph in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
    text = ''
    for run in paragraph.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
        if run.text is not None:
            text += run.text
    paragraphs.append(text)


# remove ""
paragraphs = [x for x in paragraphs if x != ""]


def has_punctuation(paragraph):
    for char in paragraph:
        if char in string.punctuation:
            return True
    return False


def test_heading(paragraph: str):
    return not has_punctuation(paragraph=paragraph)


def test_bullets(paragraph: str):
    if ":" in paragraph:
        return True
    return False


def get_lists(paragraphs: list):
    is_list = False
    lists = []

    for index, paragraph in enumerate(paragraphs):
        if is_list:
            # only for 1. some_words, 2. some_words, .....
            pattern = r'^\d+\.\s*\w+'
            if re.match(pattern, paragraph):
                # lists.append(paragraph)
                lists[len(lists)-1]['bullets'].append(paragraph)
            else:
                lists[len(lists)-1]['end'] = index
                # return lists  #return for only allowing 1st list to return
                is_list = False

        else:
            is_list = test_bullets(paragraph=paragraph)
            if is_list:
                lists.append(
                    {"title": paragraph, "start": index, "bullets": []})

    return lists


def change_to_list(paragraphs: list):
    is_list = False
    for index, paragraph in enumerate(paragraphs):
        if is_list:
            # only for 1. some_words, 2. some_words, .....
            pattern = r'^\d+\.\s*\w+'
            if re.match(pattern, paragraph):
                paragraphs[index] = f"<li>{paragraph}</li>"
            else:
                # no list present then suppose 1st mai
                # something: chha tara tala aru chhainan list 1. aru
                if "<ul>" in paragraphs[index-1]:
                    paragraphs[index-1] = paragraphs[index -
                                                     1].replace("<ul>", "")
                    continue

                is_list = False
                paragraphs[index-1] += '</ul>'

        else:
            is_list = test_bullets(paragraph=paragraph)
            if is_list:
                paragraphs[index] = f"<ul>{paragraph}"

    return paragraphs


def change_to_head(paragraphs: list):
    for index, paragraph in enumerate(paragraphs):
        if test_heading(paragraph=paragraph):
            paragraphs[index] = f'<h1>{paragraph}</h1>'


def change_to_paragraph(paragraphs: list):
    for index, paragraph in enumerate(paragraphs):
        if "<" in paragraph:
            continue
        paragraphs[index] = f'<p>{paragraph}</p>'


def change_to_ques_ans(paragraphs: list):
    for index, paragraph in enumerate(paragraphs):
        if paragraph.startswith("Q. "):
            paragraphs[
                index] = f"<h4 style='color: #37474F; font-size: 20px; font-weight: bold; margin-bottom: 10px;'>{paragraph}</h4>"
        elif paragraph.startswith("A. "):
            paragraphs[index] = f"<h4 style='color: #4CAF50;'>{paragraph}</h4>"


change_to_head(paragraphs=paragraphs)
change_to_list(paragraphs=paragraphs)
change_to_ques_ans(paragraphs=paragraphs)
# do at last
change_to_paragraph(paragraphs=paragraphs)


print(paragraphs)


# convert to html

def convert_docx_to_html(paragraphs: list):

    # Create HTML content
    html_content = '<html><head></head><body>'

    for paragraph in paragraphs:
        # Add paragraph tags
        html_content += paragraph

    # Close HTML tags
    html_content += '</body></html>'

    return html_content


html_content = convert_docx_to_html(paragraphs=paragraphs)
with open("test.html", 'w', encoding='utf-8') as file:
    file.write(html_content)
