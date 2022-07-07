import os

from docx import Document
from docx.oxml import OxmlElement, ns
from docx.shared import Pt, Mm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.text import WD_COLOR_INDEX


class Formatter:

    settings: dict

    month = {"01": "февраля", "02": "января", "03": "марта", "04": "апреля", "05": "мая",
             "06": "июня", "07": "июля", "08": "августа", "09": "сентября", "10": "октября",
             "11": "ноября", "12": "декабря"}

    def __init__(self, path, settings):
        self.doc = Document(path)
        self.path = path
        self.settings = settings

    def create_element(self, name):
        return OxmlElement(name)

    def create_attribute(self, element, name, value):
        element.set(ns.qn(name), value)

    def add_page_number(self, run):
        fldStart = self.create_element('w:fldChar')
        self.create_attribute(fldStart, 'w:fldCharType', 'begin')

        instrText = self.create_element('w:instrText')
        self.create_attribute(instrText, 'xml:space', 'preserve')
        instrText.text = "PAGE"

        fldChar1 = self.create_element('w:fldChar')
        self.create_attribute(fldChar1, 'w:fldCharType', 'separate')

        fldChar2 = self.create_element('w:t')
        fldChar2.text = "2"

        fldEnd = self.create_element('w:fldChar')
        self.create_attribute(fldEnd, 'w:fldCharType', 'end')

        run._r.append(fldStart)
        run._r.append(instrText)
        run._r.append(fldChar1)
        run._r.append(fldChar2)
        run._r.append(fldEnd)

    def get_textInput(self, paragraph):

        run = paragraph.add_run()
        self.create_attribute(run._r, 'w:rsidRPr', '00921D4A')
        rPr = self.create_element('w:rPr')
        rPr1 = self.create_element('w:szCs')
        self.create_attribute(rPr1, 'w:val', '26')
        rPr2 = self.create_element('w:highlight')
        self.create_attribute(rPr2, 'w:val', 'default')
        rPr.append(rPr1)
        rPr.append(rPr2)
        run._r.append(rPr)

        fldStart = self.create_element('w:fldChar')
        self.create_attribute(fldStart, 'w:fldCharType', 'begin')
        ffdata = self.create_element('w:ffData')
        name = self.create_element('w:name')
        self.create_attribute(name, 'w:val', 'Хуй')
        ffdata.append(name)
        enabled = self.create_element('w:enabled')
        ffdata.append(enabled)
        calc = self.create_element('w:calcOnExit')
        self.create_attribute(calc, 'w:val', '0')
        ffdata.append(calc)
        textInput = self.create_element('w:textInput')
        default = self.create_element('w:default')
        self.create_attribute(default, 'w:val', "Жопа")
        textInput.append(default)
        ffdata.append(textInput)
        fldStart.append(ffdata)
        run._r.append(fldStart)

        run2 = paragraph.add_run()
        self.create_attribute(run2._r, 'w:rsidRPr', '00921D4A')
        rPrN = self.create_element('w:rPr')
        rPrN1 = self.create_element('w:szCs')
        self.create_attribute(rPr1, 'w:val', '26')
        rPrN.append(rPrN1)
        run._r.append(rPrN)
        instrText = self.create_element('w:instrText')
        self.create_attribute(instrText, 'xml:space', 'preserve')
        instrText.text = " FORMTEXT "
        run2._r.append(instrText)

        run3 = paragraph.add_run()
        self.create_attribute(run3._r, 'w:rsidRPr', '00921D4A')
        rPrNN = self.create_element('w:rPr')
        fldChar1 = self.create_element('w:szCs')
        self.create_attribute(fldChar1, 'w:val', '26')
        rPrNN.append(fldChar1)
        run3._r.append(rPrNN)

        run4 = paragraph.add_run()
        self.create_attribute(run4._r, 'w:rsidRPr', '00921D4A')
        rPrNNN = self.create_element('w:rPr')
        fldChar2 = self.create_element('w:szCs')
        self.create_attribute(fldChar2, 'w:val', '26')
        rPrNNN.append(fldChar2)
        run4._r.append(rPrNNN)
        fldCharSep = self.create_element('w:fldChar')
        self.create_attribute(fldCharSep, 'w:fldCharType', 'separate')
        run4._r.append(fldCharSep)

        run5 = paragraph.add_run()
        self.create_attribute(run5._r, 'w:rsidRPr', '00921D4A')
        rPrNNNN = self.create_element('w:rPr')
        fldChar22 = self.create_element('w:szCs')
        self.create_attribute(fldChar22, 'w:val', '26')
        rPrNNNN.append(fldChar22)
        run5._r.append(rPrNNNN)
        fldCharText = self.create_element('w:t')
        fldCharText.text = "Жопа"
        run5._r.append(fldCharText)

        run6 = paragraph.add_run()
        self.create_attribute(run6._r, 'w:rsidRPr', '00921D4A')
        fldEnd = self.create_element('w:fldChar')
        self.create_attribute(fldEnd, 'w:fldCharType', 'end')
        run6._r.append(fldEnd)

    def set_textInput(self, paragraph):

        run = paragraph.add_run()
        for bad in run._r.xpath('//w:textInput'):
            bad.getparent().getparent().getparent().getparent().remove(bad.getparent().getparent().getparent())
        # for bad in run._r.xpath('//w:ffData'):
        #     bad.getparent().remove(bad)
        # run = paragraph.add_run()
        for bad in run._r.xpath("//w:instrText[text()=' FORMTEXT ']"):
            bad.getparent().getparent().remove(bad.getparent())
        for bad in run._r.xpath("//w:instrText[text()='FORMTEXT ']"):
            bad.getparent().getparent().remove(bad.getparent())

    # нумерация
    def numbering(self):
        print("Заголовок____________________________", self.doc.sections[0].header.paragraphs[0].text)
        if self.doc.sections[0].header.paragraphs[0].text == "":
            self.add_page_number(self.doc.sections[0].header.paragraphs[0].add_run())
            self.doc.sections[0].header.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # выравниваем по центру
            self.doc.sections[0].different_first_page_header_footer = True  # особый колонтитул для первой страницы - вкл
            #self.doc.sections[0].header.paragraphs[0].paragraph_format.space_after = Mm(100)  # отступ колонитула от вверхнего края
            sectPr = self.doc.sections[0]._sectPr  # хер его знает, стоило бы узнать
            print("Трогаю")
            pgNumType = OxmlElement('w:pgNumType')
            pgNumType.set(ns.qn('w:start'), "1")  # 1 это с какой страницы начинается отсчёт
            sectPr.append(pgNumType)

    # Функция для форматирования текст
    def Format(self):
        # Настройка отступов
        section = self.doc.sections[-1]
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(15)
        section.right_margin = Mm(15)
        section.header_distance = Mm(10)
        # отступ от нижнего края страницы до
        # нижнего края нижнего колонтитула
        section.footer_distance = Mm(10)

        # Настройка междустрочного интервала и убираем выделение корректором
        for p in self.doc.paragraphs:
            p.add_run()
            #print(p.text)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            p.runs[0].font.highlight_color = WD_COLOR_INDEX.AUTO
        # Настройка шрифта и размера текста
        style = self.doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(12)

    # функция редактирование текста
    def Redact(self):
        self.numbering()
        self.Format()
        for p in self.doc.paragraphs:  # проходим все абзацы в документе на поиск ошибок, и заменяем их
            #self.get_textInput(p)
            self.set_textInput(p)
            text = p.text
            if p.text == "":
                continue
            #print(text, "------------------", p.style.name, p.style.font.bold, p.style.font.size)
            # флаг, который отвечает, есть ли ошибка в абзаце
            # если есть, то правим и заменяем текс, если нет, то нет
            flag = False
            if list(self.settings.values())[0]:
                if " N " in p.text:  # проверяем, если ли N, если есть заменяем на №
                    text = text.replace(' N ', ' № ')
                    flag = True
            if list(self.settings.values())[2]:
                if ' "' in p.text and '" ' in p.text:  # проверяем, если ли "...", если есть заменяем на «...»
                    text = text.replace(' "', ' «')
                    text = text.replace('" ', '» ')
                    flag = True
                if '“' in p.text and '”' in p.text:  # проверяем, если ли “...”, если есть заменяем на «...»
                    text = text.replace('“', ' «')
                    text = text.replace('”', '» ')
                    flag = True
            if list(self.settings.values())[4]:
                if ' - ' in p.text:  # проверяем, если ли -, если есть заменяем на –
                    text = text.replace(' - ', ' – ')
                    flag = True
            if list(self.settings.values())[3]:
                if ' т.е. ' in p.text:  # проверяем, если ли т.е., если есть заменяем на то есть
                    text = text.replace(' т.е. ', ' то есть ')
                    flag = True
            if 'решение Арбитражный суд' in p.text:
                text = text.replace('решение Арбитражный суд','решение Арбитражного суда')
                flag = True
            if 'определение Арбитражный суд' in p.text:
                flag = True
                text = text.replace('определение Арбитражный суд','определение Арбитражного суда')
            # если есть хотя бы одна ошибка в абзаце, меняем на исправленный вариант
            if flag:
                style = p.style
                p.text = text
                p.style = style
        # редактирование даты, например 12.03.2021 или 12 октября 2021 г.
        # в 12 октября 2021 года
        if list(self.settings.values())[1]:
            if "Дело" in self.doc.paragraphs[5].text:
                text = str(self.doc.paragraphs[5].text)
                flag = False
                if "." in text[0:10]:
                    monthNumb = text[3:5]
                    if monthNumb in self.month.keys():
                        text = text.replace(text[0:10], f"{text[0:10]} года")
                        text = text.replace(f".{monthNumb}.", f" {self.month.get(monthNumb)} ")
                        flag = True
                elif "г.":
                    text = text.replace("г.", "года ")
                    flag = True
                if flag:
                    style = p.style
                    self.doc.paragraphs[5].text = text
                    p.style = style
        self.doc.save("test1.docx")
        os.startfile("test1.docx")
