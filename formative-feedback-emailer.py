#!/usr/bin/env python3


"""
Simple utility to draft a batch of emails based on anonymized file names.
Emails are not sent automatically but only opened as drafts:
any email has to be sent out with manual confirmation.
Assumes Outlook is running in the background.

To adapt the email template, edit the html string in the function
"create_email_body" below.

Gregor Boes, 2021, gregor.boes@kcl.ac.uk
"""

import datetime
import win32com.client
import sys
import os
import re

pathbase = r"D:\gdrive\boox\review\Epistemology Marking\\"
files = os.listdir(pathbase)
files = [f for f in files if re.search(r".docx|.pdf",f)]
# student_ids = re.findall(r"[kK][0-9]+", ' '.join(files))

try:
    outlook = win32com.client.Dispatch('Outlook.Application')
except:
    outlook = win32com.client.GetActiveObject('Outlook.Application')

def create_email_body():
    body = \
    f"""<html><body>Hello, <br><br>
    Please find attached the feedback to your formative essay. The grading happened anonymously, if you removed identifying information from your file and filename.
     <br><br>
    Please let me know if you want to book a slot in the feedback session, which I will schedule for 2pm, Friday 17th December. This can be attended in person or remotely.
    If you find the feedback particularly helpful or insufficient, I would also be happy to hear about it.
    <br><br>
    All the best,<br>
    Gregor
    </body></html>
    """
    return body

# for student_id in student_ids:
for f in files:
    try:
        student_id = re.search(r"[kK][0-9]+", f)[0]
    except TypeError:
        student_id = "unknown"
    mail = outlook.CreateItem(0)
    mail.HTMLBody = create_email_body()
    # breakpoint()
    mail.To = f"{student_id}@kcl.ac.uk"
    mail.Subject = "Formative Essay Feedback for Epistemology I"
    mail.Attachments.Add(pathbase+f)
    mail.Display(True)
