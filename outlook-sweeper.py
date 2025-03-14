from win32com.client import Dispatch
import extract_msg
from xhtml2pdf import pisa
import os
import re
import tkinter as tk
import datetime

def main():

    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook.Folders.Item(2)
    sweeper_folder = folder.Folders['Sweeper Testing'] # will need to be updated to account for the specific production environment; probably set this up to take user input for folder selection
    messages = sweeper_folder.items

    current_date = datetime.datetime.now()
    formatted_date = str(current_date.year) + str(current_date.month) + str(current_date.day)

    for message in messages:
        name = str(message.subject)
        #to eliminate any special charecters in the name
        name = re.sub('[^A-Za-z0-9]+', '', name)+'.msg'
        #to save in the current working directory
        message.SaveAs(os.getcwd()+'//'+name)

        f = name
        msg = extract_msg.Message(f)
        msg_sender = msg.sender.strip('\"').strip('\"')
        msg_recip = msg.to
        msg_date = msg.date
        msg_subj = msg.subject
        msg_message = msg.body
        msg.close()

        print('Sender: {}'.format(msg_sender))
        print('Sent On: {}'.format(msg_date))
        print('Subject: {}'.format(msg_subj))
        print('Body: {}'.format(msg_message))

        msg_message = re.sub(r'(\n\s*)+\n', '\n\n', msg_message)

        breaks = msg_message.replace('\r', '<br>')

        stu_uid = uid_label_entry.get()
        stu_fname = fname_label_entry.get()
        stu_lname = lname_label_entry.get()

        html = "<html><head></head><body><div style='font-size: 18px; font-family: sans-serif; padding-bottom: 10px;'>{} {}, {}, - Reason</div>".format(stu_uid, stu_lname, stu_fname) + "<div style='font-family: sans-serif;'><strong>From:</strong> {}<br><strong>To:</strong> {}<br><strong>Subject:</strong> {}<br><strong>Date:</strong> {}<br><hr style='height:2px; color:black; background-color: black;'></div><p style='font-size:14px;'>".format(msg_sender, msg_recip, msg.subject, msg_date) + breaks + "</p></body></html>"

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
        output_filename = f'{stu_lname},{stu_fname}_{formatted_date}.pdf'
        convert_html_to_pdf(source_html, output_filename)
        os.remove(name)

# Create the main window
root = tk.Tk()
root.title("Alison's Email Sweeper")

# Create entry fields for UTM parameters
fname_label = tk.Label(root, text="Student First Name:")
fname_label.pack()
fname_label_entry = tk.Entry(root)
fname_label_entry.pack()

lname_label = tk.Label(root, text="Student Last Name:")
lname_label.pack()
lname_label_entry = tk.Entry(root)
lname_label_entry.pack()

uid_label = tk.Label(root, text="Student UID:")
uid_label.pack()
uid_label_entry = tk.Entry(root)
uid_label_entry.pack()

# Create a button to select the HTML file
select_button = tk.Button(root, text="Run", command=main)
select_button.pack()

# Display a message to guide the user
message_label = tk.Label(root, text="Lol")
message_label.pack()

root.mainloop()