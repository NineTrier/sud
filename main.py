from docx import Document
from docx.shared import Inches, Pt

doc = Document('Решение.docx')


def ReplaceN():
    for p in doc.paragraphs:
        if " N " in p.text:
            text = p.text.replace(' N ',' № ')
            style = p.style
            p.text = text
            p.style = style
        if ' "' in p.text and '" ' in p.text:
            text = p.text.replace(' "',' «')
            text1 = text.replace('" ','» ')
            style = p.style
            p.text = text1
            p.style = style
        if '“' in p.text and '”' in p.text:
            text = p.text.replace('“',' «')
            text1 = text.replace('”','» ')
            style = p.style
            p.text = text1
            p.style = style
    doc.save('test.docx')

ReplaceN()
