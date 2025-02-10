from win32com.client import Dispatch
import extract_msg
import weasyprint
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

f = r'REProofreview525AYand26AYemails.msg'  # currently hardcoded
msg = extract_msg.Message(f)
msg_sender = msg.sender
msg_date = msg.date
msg_subj = msg.subject
msg_message = msg.body
msg.close()

print('Sender: {}'.format(msg_sender))
print('Sent On: {}'.format(msg_date))
print('Subject: {}'.format(msg_subj))
print('Body: {}'.format(msg_message))

html = "<html><head></head><body><p>" + msg.body.replace('\n', "<br>") + "</p></body></html>"

text_file = open("Output.html", "w")
text_file.write(html)
text_file.close()

pdf = weasyprint.HTML('hOutput.html').write_pdf()