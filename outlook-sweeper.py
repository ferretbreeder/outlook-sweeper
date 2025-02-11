from win32com.client import Dispatch
import extract_msg
from xhtml2pdf import pisa
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

f = name
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

def convert_html_to_pdf(source_html, output_filename):
    # open output file for writing (truncated binary)
    result_file = open(output_filename, "w+b")

    # convert HTML to PDF
    pisa_status = pisa.CreatePDF(
            source_html,                # the HTML to convert
            dest=result_file)           # file handle to recieve result

    # close output file
    result_file.close()                 # close output file

    # return False on success and True on errors
    return pisa_status.err

# Define your data
source_html = open('Output.html')
output_filename = f'{name}.pdf'
convert_html_to_pdf(source_html, output_filename)