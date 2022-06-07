import datetime
import os
import win32com.client

class emails:
    def __init__(self, meetingDate):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.inbox = self.outlook.GetDefaultFolder(6)
        #Folders('admin@hertford.gov.uk').Folders('Inbox')
        self.messages = self.inbox.Items
        self.meetingFolder = meetingDate

        self.folder = self.outlook.Folders.Item("Mailbox Name")
        self.inbox = self.folder.Folders.Item("Inbox")

        self.subject = "Planning Application Decision"

    #meeting date in format "25.05.2020"
    def getfolder(self):
        self.meetingFolder = "F:/D & L Committee/PLANNING SUB/PLANS/" + self.meetingDate + "/decisions"

    def saveAttachments(self):
        for message in self.messages:
            
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
    
    def checkFolderNums(self):
        for i in range(50):
            try:
                box = self.outlook.GetDefaultFolder(i)
                name = box.Name
                print(i, name)
            except:
                pass
        

    def PrintAccounts(self):
        for account in self.outlook.Accounts:
            print(account.DeliveryStore.DisplayName)


emailTest = emails("29.04.20")
#emailTest.saveAttachments()
#emailTest.checkFolderNums()
emailTest.PrintAccounts()