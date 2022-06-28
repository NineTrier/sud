from docx import Document
from docx.shared import Inches, Pt

doc = Document('Решение.docx')

for p in doc.paragraphs:
    if " N " in p.text:
        text = p.text.replace(' N ',' № ')
        style = p.style
        p.text = text
        p.style = style
    if '"' in p.text:
        text = p.text.replace('"', '')
        style = p.style
        p.text = text
        p.style = style

    doc.save('test.docx')
