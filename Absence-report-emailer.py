"""
Simple utility to draft emails to tutors of absent students.
Currently set to not send, but only draft these.
"""
# import pyOutlook as out
import datetime
import win32com.client
import sys
import pandas as pd



try:
    outlook = win32com.client.Dispatch('Outlook.Application')
except:
    outlook = win32com.client.GetActiveObject('Outlook.Application')

# Collect Student Names
student_names = []
t_add = pd.read_csv("./tutor_addresslist.csv")

while True:
    newinput = input("Add student to reports - or type 'DONE' to continue to tutors, 'DELETE' to remove last addition, or a local path to a .csv or .txt file with student names and tutor names. \n")
    if newinput[-4::] == ".txt":
        df = pd.read_table(newinput)
        student_names = df.Name.tolist()
        print("Reporting absence for:\n", *student_names, sep="\n")
        break
    if newinput[-4::] == ".csv":
        df = pd.read_csv(newinput)
        student_names = df.Name.values.tolist()
        print("Reporting absence for:\n", *student_names, sep="\n")
        break
    if newinput.upper() == "DONE":
        break
    if newinput == "DELETE":
        print(f"Deleted {student_names.pop()} from students")
    else:
        name = ' '.join([k.capitalize() for k in newinput.split()])
        student_names.append(name)
        print (f"Added {name} to list of absent students.")

# Collect Tutor Names and Emails
tutor_names = []
tutor_addresses = []
for student in student_names:
    tutor_name = input(f"Name of Tutor for {student}\n")
    tutor_names.append(tutor_name)
    if tutor_name in t_add.Name.values:
        tutor_address = t_add.loc[t_add.Name==tutor_name]["Mail"].values.tolist()[0]
        tutor_addresses.append(tutor_address)
        print(f"Found address for {tutor_name}:\t'{tutor_address}'")
        continue
    default_address = (".".join(tutor_name.split())).lower() + "@kcl.ac.uk"
    tutor_address = input(f"Email address for {tutor_name} (RETURN for {default_address})\n")
    if tutor_address == "":
        tutor_addresses.append(default_address)
        t_add = t_add.append({"Name":tutor_name, "Mail":default_address}, ignore_index=True)
        t_add.to_csv("./tutor_addresslist.csv", index=False)
        print(f"Added{tutor_name}:\t'{default_address}' to address list.")
    else:
        tutor_addresses.append(tutor_address)
        t_add = t_add.append({"Name":tutor_name, "Mail":tutor_address}, ignore_index=True)
        t_add.to_csv("./tutor_addresslist.csv", index=False)
        print(f"Added{tutor_name}:\t'{tutor_address}' to address list.")

# Collect Seminar Name
seminar_name = input("Enter Seminar Name (e.g. 'Methodology')\n")

# Collect Seminar Time
seminar_time = input("Enter Seminar Time (e.g. 'Wednesday, 11:00-12:00')\n")

# Draft Outlook Emails to Tutors
def create_email_body(student_name, tutor_name, seminar_name, seminar_time):
    body = \
    f"""<html><body>Dear {tutor_name.split()[0]}, <br><br>

    I am reporting the absence of <b>{student_name}</b> in two consecutive seminars. I just want to check they are alright.
    <br>

    <br>
    This concerns the seminar <b>'{seminar_name}'</b>, {seminar_time}.</par>
    <br><br>

    Best,<br>
    Gregor</body></html>
    """
    return body


# =============================================================================
# Create Emails
# =============================================================================
for i in range(len(student_names)):
    # Open a Message window per student
    mail = outlook.CreateItem(0)
    mail.HTMLBody = create_email_body(student_names[i], tutor_names[i], seminar_name,
                                  seminar_time)
    mail.To = tutor_addresses[i]
    mail.Subject = f"{student_names[i]}: Two Consecutive Absences in '{seminar_name}''"
    mail.Display(True)


# Save report of reported absences and .csv list of tutor/student information
date = str(datetime.datetime.today())
student_namestring = "\n".join(student_names)
report = \
f"""
Reported Absences

on {date}:

{student_namestring}
"""
print(report)
