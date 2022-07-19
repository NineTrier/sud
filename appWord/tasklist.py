
def get_all_Date_Time(text):
    dates = []
    times = []
    for i in range(len(text)):
        try:
            boolDat, Dat = is_Date(text[i:i+10])
            boolTime, Time = is_Time(text[i:i+5])
        except:
            continue
        if boolDat:
            dates.append(Dat)
        if boolTime:
            times.append(Time)
    return dates, times

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
        self.create_attribute(name, 'w:val', 'Тратата')
        ffdata.append(name)
        enabled = self.create_element('w:enabled')
        ffdata.append(enabled)
        calc = self.create_element('w:calcOnExit')
        self.create_attribute(calc, 'w:val', '0')
        ffdata.append(calc)
        textInput = self.create_element('w:textInput')
        default = self.create_element('w:default')
        self.create_attribute(default, 'w:val', "Текст")
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
        fldCharText.text = "Текст"
        run5._r.append(fldCharText)

        run6 = paragraph.add_run()
        self.create_attribute(run6._r, 'w:rsidRPr', '00921D4A')
        fldEnd = self.create_element('w:fldChar')
        self.create_attribute(fldEnd, 'w:fldCharType', 'end')
        run6._r.append(fldEnd)


def is_Date(text: str):
    if text[2] == '.' and text[5] == '.':
        for i in range(len(text)):
            if i == 2 or i ==   5:
                continue
            if not ('0' <= text[i] <= '9'):
                if i == 8 and text[9] == " ":
                    return True, f"{text[:8]}"
                return False, ""
        return True, text
    else:
        return False, ""

def is_Time(text: str):
    if text[2] == ':':
        for i in range(len(text)):
            if i == 2:
                continue
            if not ('0' <= text[i] <= '9'):
                return False, ""
        return True, text
    else:
        return False, ""



