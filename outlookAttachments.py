import datetime
import os
import win32com.client

class emails:
    def __init__(self, meetingDate):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.inbox = self.outlook.Folders('admin@hertford.gov.uk').Folders('Inbox')
        self.messages = self.inbox.Items
        self.meetingFolder = meetingDate
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


emailTest = emails("29.04.20")
emailTest.saveAttachments()