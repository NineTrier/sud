from docx import Document
from docx.oxml import OxmlElement, ns
from docx.shared import Inches, Pt, Mm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_COLOR_INDEX

doc = Document('Решение.docx')

month = {"01": "февраля", "02": "января", "03": "марта", "04": "апреля", "05": "мая"
    , "06": "июня", "07": "июля", "08": "августа", "09": "сентября", "10": "октября"
    , "11": "ноября", "12": "декабря"}


def create_element(name):
    return OxmlElement(name)


def create_attribute(element, name, value):
    element.set(ns.qn(name), value)


def add_page_number(run):
    fldChar1 = create_element('w:fldChar')
    create_attribute(fldChar1, 'w:fldCharType', 'begin')

    instrText = create_element('w:instrText')
    create_attribute(instrText, 'xml:space', 'preserve')
    instrText.text = "PAGE"

    fldChar2 = create_element('w:fldChar')
    create_attribute(fldChar2, 'w:fldCharType', 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)


# Функция для форматирования текст
def Format():
    # Настройка отступов
    section = doc.sections[-1]
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.left_margin = Mm(15)
    section.right_margin = Mm(15)

    # Настройка междустрочного интервала и убираем выделение корректором
    for p in doc.paragraphs:
        p.paragraph_format.line_spacing = Pt(15)
        p.runs[0].font.highlight_color = WD_COLOR_INDEX.AUTO
    # Настройка шрифта и размера текста
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)


# функция редактирование текста
def Redact():
    for p in doc.paragraphs:
        text = p.text
        flag = False
        if " N " in p.text:
            text = text.replace(' N ', ' № ')
            flag = True
        if ' "' in p.text and '" ' in p.text:
            text = text.replace(' "', ' «')
            text = text.replace('" ', '» ')
            flag = True
        if '“' in p.text and '”' in p.text:
            text = text.replace('“', ' «')
            text = text.replace('”', '» ')
            flag = True
        if ' - ' in p.text:
            text = text.replace(' - ', ' – ')
            flag = True
        if ' т.е. ' in p.text:
            text = text.replace(' т.е. ', ' то есть ')
            flag = True
        if flag:
            style = p.style
            p.text = text
            p.style = style

    if "Дело" in doc.paragraphs[5].text:
        text = str(doc.paragraphs[5].text)
        flag = False
        if "." in text[0:10]:
            monthNumb = text[3:5]
            if monthNumb in month.keys():
                text = text.replace(text[0:10], f"{text[0:10]} года")
                text = text.replace(f".{monthNumb}.", f" {month.get(monthNumb)} ")
                flag = True
        elif "г.":
            text = text.replace("г.", "года ")
            flag = True
        if flag:
            style = p.style
            doc.paragraphs[5].text = text
            p.style = style
    doc.save('test.docx')


add_page_number(doc.sections[0].header.paragraphs[0].add_run())
doc.sections[0].header.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
Format()
Redact()
