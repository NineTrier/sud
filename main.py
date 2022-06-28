from docx import Document
from docx.shared import Inches, Pt

doc = Document('Решение.docx')

month = {1: "февраля", 2: "января", 3: "марта", 4: "апреля", 5: "мая"
     ,6: "июня", 7: "июля", 8: "августа", 9: "сентября", 10: "октября"
    ,11: "ноября", 12: "декабря"}

# def Format():
#     # for p in doc.paragraphs:
#     #
#     #     doc.paragraph_format.line_spacing = Pt(30)

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
        if "." in text[0:10]:
            monthNumb = int(text[3:5])
            if monthNumb in month.keys():
                text = text.replace(f".{monthNumb}.", f" {month.get(monthNumb)} ")
        elif "г.":
            text = text.replace("г.", "года ")
        style = p.style
        doc.paragraphs[5].text = text
        p.style = style
    doc.save('test.docx')

Redact()
