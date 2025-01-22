from win32com.client import Dispatch
import os
import re

outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
print(inbox)
messages = inbox.items
message = messages.GetLast()
name = str(message.subject)
#to eliminate any special charecters in the name
name = re.sub('[^A-Za-z0-9]+', '', name)+'.msg'
#to save in the current working directory
message.SaveAs(os.getcwd()+'//'+name)