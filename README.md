# absence_reporter

Small Utility for reporting absent students in seminar groups. 
Open Outlook, select a folder in the email account you want to email from, and run from the command line: 
```python3 Absence-report-emailer.py```

Connects to Outlook, drafts emails with a template text and basic information.
To change the template, open ```Absence-report-emailer.py``` file with a text editor (e.g. Notepad). 
Tutor emails are auto-completed if they have occurred before (stored in ```tutor_addresslist.csv```). 
If you have a list of students for whom you want to send an absentee report, they can be entered in the table "absences.csv", with each name on a new line.
Absentee and instructor information can be entered directly into the .csv table or interactively on the command line (and they will be saved in the table afterwards). 
