# Python Calendar 📆

Create an iCalendar (\*.ics) from a XLSX file with Python.
The script searches for a specific string, in this case "Paolo", and extracts the cell with its date.
Then it converts the object to a calendar in the \*ics format, so that it can be imported into Google Calendar or similar software.

Change the row and column index if you want to use this script on your own XLSX file (and the string to search for).

### How to run 🏃🏻

1. `python3 -m venv env`
1. `source env/bin/activate`
1. `python3 -m pip install -r requirements.txt`
1. `python3 script.py`
