import pdfplumber as plmr
import re
import shutil
import docx
import os
import PletstrepReader as ps

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
        self.ehDecision = self.extractDecision()
        #add to the decision list for extraction
        self.decisionList = [self.appNo, self.location, self.ehDecision]

    #the job of extracting the decision from the document is trickier, so I have made an extra function for it.
    #Will search for phrases from a list in a text document that can be updated.
    def extractDecision(self):
        #print(self.decisionPhrasesEH)((.|\n)*)
        #print(self.text)
        txtUpper = self.text.upper()
        txtOneLine = ''.join(txtUpper.splitlines())
        for decisionPhrase in self.decisionPhrasesEH:
            if decisionPhrase.upper() in txtOneLine:
                return decisionPhrase
            #In case the phrase is not in the TXT document, you will need to add it manually
        #self.ehDecision = "Search East Herts Decision Manually"

    #This will run the above functions so you need only call one. Returns a list with the appNo, location and east herts Decision
    def getInfo(self):
        self.PDFtoText()
        #self.printText()
        self.extractInfo()
        return self.decisionList

#A class to allow formatting for multiple pdfs in a certain folder. folderPath is a directory in a string with the pdfs in them
class pdfFolder:
    def __init__(self, folderPath, meetingDate):
        self.folderPath = folderPath
        self.firstDecisionsList = []
        self.meetingDate = meetingDate
    
    #this function will iterate through the pdf files in the given folder
    def getInfoFromFolder(self):
        for file in os.listdir(self.folderPath):
            print(file)
            if not file.endswith(".pdf"):
                continue
            filePath = self.folderPath + file
            print(filePath)
            pdfFile = pdfInfo(filePath)
            self.firstDecisionsList.append(pdfFile.getInfo())

        pletstrep1 = ps.pletstrep(self.firstDecisionsList, self.meetingDate)
        return pletstrep1.getHTC()


#class for setting up and entering information into a table in a word document.
class createDocx:
    #filepath is the location of the word template.
    #meetingDate will be a string in the format "25 March 2022" - need a way to get this from the user.
    def __init__(self, allDecisionList, meetingDate):
        self.allDecisionList = allDecisionList
        self.meetingDate = meetingDate
        self.fileName = r'Paper D ' + meetingDate + r'.docx'
        #self.filePath = filePath + self.fileName
    
#--------------------------------------------------------------------------------------------------------------------------#
#FilePaths will need to be altered for live use
    def copyTemplate(self):
        original = r'C:/Users/JoeFeatherstone/Documents/Python Projects/D&L Meeting Decision Interpreter/Paper D Template.docx'
        target = r'C:/Users/JoeFeatherstone/Documents/Python Projects/D&L Meeting Decision Interpreter/Test documents/' + self.fileName
        shutil.copyfile(original, target)
        self.filePath = target

    def changeDate(self):
        self.doc = docx.Document(self.filePath)
        for p in self.doc.paragraphs:
            if '*DATE*' in p.text:
                inline = p.runs
                # Loop added to work with runs (strings with same style)
                for i in range(len(inline)):
                    if '*DATE*' in inline[i].text:
                        text = inline[i].text.replace('*DATE*', self.meetingDate)
                        inline[i].text = text
                #print(p.text)
                self.doc.save(self.filePath)

    def appendTable(self):
        self.doc.tables
        #print("Retrieved value: " + self.doc.tables[0].cell(0, 0).text)
        #self.doc.tables[0].cell(1, 0).text = "new value1"
        #self.doc.tables[0].add_row() #ADD ROW HERE
        #self.doc.tables[0].cell(2, 1).text = "new value2"
        #self.doc.tables[0].add_row()
        #self.doc.tables[0].cell(3, 2).text = "new value3"
        row = 1
        column = 0
        for decisionList in self.allDecisionList:
            for s in decisionList:
                self.doc.tables[0].cell(row, column).text = s
                column += 1
            row += 1
            column = 0
            self.doc.tables[0].add_row()
        self.doc.save(self.filePath)




#pdf1 = pdfInfo(r'C:/Users/JoeFeatherstone/Documents/Python Projects/D&L Meeting Decision Interpreter/304A Ware Road.pdf')
#pdf1.printText()
#List1 = pdf1.getInfo()
#print(List1)
filePath = r'C:/Users/JoeFeatherstone/Documents/Python Projects/D&L Meeting Decision Interpreter/Test documents/'
date = r'25 March 2022'
folder = pdfFolder('C:/Users/JoeFeatherstone/Documents/Python Projects/D&L Meeting Decision Interpreter/pdf folder/', '25.04.2022')
allDecisionsList = folder.getInfoFromFolder()
print(allDecisionsList)
paperD = createDocx(allDecisionsList, date)
paperD.copyTemplate()
paperD.changeDate()
paperD.appendTable()