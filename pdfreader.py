import pdfplumber as plmr
import re

#Class that will read PDFs and Gather information about them
class pdfInfo:

    #The fikePath s just where the pdf is on the computer.
    def __init__(self, filePath):
        self.filePath = filePath
        #data that needs to be extracted from the PDF
        self.location = ""
        self.appNo = ""
        self.decision = ""
    
    #The PDF file in plaintext, ready for processing information
    

    #This will open the pdf specified when initiating the pdfInfo object.
    def getInfo(self):
        #opens a pdf, selects the first page, then extracts the text
        pdf = plmr.open(self.filePath)
        page = pdf.pages[0]
        self.text = page.extract_text()

    #testing to make sure getInfo works properly
    def printText(self):
        print(self.text)
    
    #Use regular expressions to take the information out of the string. These may all need to be tweaked depending on how the files are layed out. 
    def extractInfo(self):
        #this line will search the text string for a certain term then return all characters after it and before a new line
        #the .group part at the end of the line isolates the string returned by the expression.
        self.appNo = re.search(r'(?<=Hertford Town Council Our Ref: ).*', self.text).group()
        #print(self.appNo)
        #find the line that contains the full address (identified by everything after "AT:")
        self.locationFull = re.search(r'(?<=AT: ).*?\n', self.text).group()
        self.location = re.search(r'.+?(?= Hertford)', self.locationFull).group()
        #print(self.location)
        #find the decision by looking for key terms.

#class for setting up and entering information into a table in a word document.
class createDocx:
    #filepath is the location of the word template.
    def __init__(self, filePath, appNo, location, decision, comment):
        selffilePath = self.filePath
        appNo = a



#pdf1 = pdfInfo('E:\\Python Projects\\D&L Meeting Decision Interpreter\\304A Ware Road.pdf')
#pdf1.getInfo()
#pdf1.printText()
#pdf1.extractInfo()