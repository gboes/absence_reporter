"""
Simple utility to draft emails to tutors of absent students.
Currently set to not send, but only draft these.
"""
import pyOutlook
import datetime

# Collect Student Names
student_names = []
while True:
    newinput = input("Add student to reports - or type 'DONE' to continue to tutors, DELETE to remove last addition")
    if newinput == "DONE":
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
    tutor_name = input("Name of Tutor for {student}")
    tutor_names.append(tutor_name)
    default_address = (".".join(tutor_name.split())) + "@kcl.ac.uk"
    tutor_address = input("Email address for {tutor_name} (RETURN for {default_address})")
    if tutor_address == "":
        tutor_addresses.append(default_address)
    else:
        tutor_addresses.append(tutor_address)

# Collect Seminar Name
seminar_name = input("Enter Seminar Name (e.g. 'Methodology')")

# Collect Seminar Time
seminar_time = input("Enter Seminar Time (e.g. 'Wednesay, 11:00-12:00')")

# Draft Outlook Emails to Tutors
def create_email_body(student_name, tutor_name, seminar_name, seminar_time):
    template =
    f"""
    Dear {tutor_name.split()[0]},

    I am reporting the absence of <b>{student_name}</b> in two consecutive seminars. I just want to check they are alright.

    This concerns the seminar '{seminar_name}', {seminar_time}.

    Best,
    Gregor
    """

# Save report of reported absences
date = str(datetime.datetime.today())
report =
f"""
Reported Absences

on {date}

{"\nl".join(student_names)}
"""
