import datetime
import os
import win32com.client

class emails:
    def __init__(self, meetingDate):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.folder = self.outlook.Folders.Item('joe.featherstone@hertford.gov.uk')
        self.decicisionWorkingFolder = self.folder.Folders.Item("Decision Working")
        self.decisionArchiveFolder = self.folder.Folders.Item("Decision Archive")
        self.msg = self.inbox.Items
        self.msgs = self.inbox.GetLast()
        #self.meetingFolder = meetingDate

        self.subject = "Planning Application Decision"

    #print msgs
    def printMsgs(self):
        print(self.msgs)
        print(self.msgs.subject)

    #meeting date in format "25.05.2020"
    def getfolder(self):
        self.meetingFolder = "F:/D & L Committee/PLANNING SUB/PLANS/" + self.meetingDate + "/decisions"


    def saveAttachments(self):
        for message in self.msgs:
            
                body_content = message.body
                print(body_content)
                break
                #attachments = message.Attachments
                #attachment = attachments.Item(1)
                #for attachment in message.Attachments:
                #    attachment.SaveAsFile(os.path.join(self.meetingFolder, str(attachment)))
                #    if message.Subject == subject and message.Unread:
                #        message.Unread = False
                #    break
            print(account.DeliveryStore.DisplayName)



emailTest = emails("29.04.20")
#emailTest.saveAttachments()
#emailTest.checkFolderNums()
#emailTest.PrintAccounts()
emailTest.printMsgs()