import docx
import re

#This class will create a file based on the date of the meeting and format it
#infoList is in the format [[application number, location of application, EHDC decision, HTC comments], app2, app3...]
class paperD:
    def __init__(self, infoList):
        self.infoList = infoList
    
    #asks user for the meeting date, stores it in the format dd/mm/yy in the string self.meetingDate
    def getDate(self):
        dateMatch = re.compile("^[0-9]{2}/[0-9]{2}/[0-9]{2}$")
        self.meetingDate = input("Enter the date of the meeting dd/mm/yy: ")
        while bool(re.search(dateMatch, self.meetingDate)) == False:
            self.meetingDate = input("Please enter the date in the format dd/mm/yy: ")
        print(re.search(dateMatch, self.meetingDate))
        print("date successfully stored: " + self.meetingDate)

    def addToTable(self):


newPaper = paperD([])
newPaper.getDate()