import pdfplumber as plmr

filePath = "C:/Users/JoeFeatherstone/Documents/Python Projects/D&L Meeting Decision Interpreter/pdf folder/209 Ware Road.pdf"

def PDFtoText(filePath):
        
        #opens a pdf, selects the first page, then extracts the text
        pdf = plmr.open(filePath)
        page = pdf.pages[0]
        return page.extract_text()

txt = PDFtoText(filePath)
uppertxt = txt.upper()
print(' '.join(uppertxt.splitlines()))

#print(txt)

if "Grant Planning Permission Subject to Conditions".upper() in ''.join(uppertxt.splitlines()):
    print("Text found")
else:
    print("Text not Found")