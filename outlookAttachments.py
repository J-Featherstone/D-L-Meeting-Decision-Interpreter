#import datetime
import os
import win32com.client

class emails:
    def __init__(self, meetingDate):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.acc = self.outlook.Folders.Item("joe.featherstone@hertford.gov.uk")
        self.working = self.acc.folders("Decision Working")
        self.msgs = self.working.Items
        self.meetingDate = meetingDate
        #self.msgs = self.inbox.GetLast()
        #self.meetingFolder = meetingDate

    #print msgs
    def printMsgs(self):
        for msg in self.msgs:
            print(msg.Subject)
        #print(self.msgs.subject)

    #meeting date in format "25.05.2020"
    def getFolder(self):
        self.meetingFolder = "F:/D & L Committee/PLANNING SUB/PLANS/" + self.meetingDate + "/decisions"


    def saveAttachments(self):
        self.getFolder()
        for message in self.msgs:
                for attachment in message.Attachments:
                    attachment.SaveAsFile(os.path.join(self.meetingFolder, str(attachment)))
                    #if message.Subject == subject and message.Unread:
                    #    message.Unread = False
                    #break
            #print(account.DeliveryStore.DisplayName)



#emailTest = emails("28.06.2022")
#emailTest.saveAttachments()
#emailTest.checkFolderNums()
#emailTest.PrintAccounts()
#emailTest.printMsgs()