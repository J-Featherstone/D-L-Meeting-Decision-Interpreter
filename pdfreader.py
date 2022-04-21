import pdfplumber as plmr
import re
import shutil
from docx import Document

#Class that will read PDFs and Gather information about them
class pdfInfo:

    #The fikePath s just where the pdf is on the computer.
    def __init__(self, filePath):
        self.filePath = filePath
        #data that needs to be extracted from the PDF
        self.location = ""
        self.appNo = ""
        self.ehDecision = ""
        #creates the decisionPhrasesEH list with
        with open('EHDC Decision Phrases.txt', 'r') as f:
            self.decisionPhrasesEH = [line.strip() for line in f]

    #The PDF file in plaintext, ready for processing information
    

    #This will open the pdf specified when initiating the pdfInfo object.
    def PDFtoText(self):
        
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
        self.extractDecision()
        #add to the decision list for extraction
        self.decisionList = [self.appNo, self.location, self.ehDecision]

    #the job of extracting the decision from the document is trickier, so I have made an extra function for it.
    #Will search for phrases from a list in a text document that can be updated.
    def extractDecision(self):
        for decisionPhrase in self.decisionPhrasesEH:
            if decisionPhrase in self.text:
                self.ehDecision = decisionPhrase
                break
            #In case the phrase is not in the TXT document, you will need to add it manually
            else:
                self.ehDecision = "Search East Herts Decision Manually"

    #This will run the above functions so you need only call one. Returns a list with the appNo, location and east herts Decision
    def getInfo(self):
        self.PDFtoText()
        #self.printText()
        self.extractInfo()
        return self.decisionList

#class for setting up and entering information into a table in a word document.
class createDocx:
    #filepath is the location of the word template.
    #meetingDate will be a string in the format "25 March 2022" - need a way to get this from the user.
    def __init__(self, filePath, decisionList, meetingDate):
        self.decisionList = decisionList
        self.meetingDate = meetingDate
        self.fileName = "Paper D " + meetingDate + ".doc"
        self.filePath = filePath + self.fileName
    
#--------------------------------------------------------------------------------------------------------------------------#
#FilePaths will need to be altered for live use
    def copyTemplate(self):
        original = r'C:\\Users\\JoeFeatherstone\\Documents\\Python Projects\\D&L Meeting Decision Interpreter\\Paper D Template.doc'
        target = r'C:\\Users\\JoeFeatherstone\\Documents\\Python Projects\\D&L Meeting Decision Interpreter\\Test documents\\' + self.fileName
        shutil.copyfile(original, target)

    def changeDate(self):




pdf1 = pdfInfo('C:\\Users\\JoeFeatherstone\\Documents\\Python Projects\\D&L Meeting Decision Interpreter\\304A Ware Road.pdf')
#pdf1.printText()
List1 = pdf1.getInfo()
#print(List1)
filePath = r'C:\\Users\\JoeFeatherstone\\Documents\\Python Projects\\D&L Meeting Decision Interpreter\\Test documents\\'
date = "25 March 2022"
paperD = createDocx(filePath, List1 , date)
paperD.copyTemplate()