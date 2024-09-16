from datetime import datetime
import pandas as pd
from ics import Calendar, Event

# Load the Excel file
file_path = 'lessons.xlsx'
df = pd.read_excel(file_path)

# Find all cells that contain "Paolo"
matches = df.isin(['Paolo'])

lessons = []

for row_idx, col_idx in zip(*matches.to_numpy().nonzero()):
    # First cell of the same row (assuming it's column 0)
    hours = df.iloc[row_idx, 0]
    
    day = ""

    # Up to row and col with the date related to the selected hour
    if hours == "9.00-10.00":
        day = df.iloc[row_idx - 2, col_idx - 1]
    elif hours == "10.00-11.00":
        day = df.iloc[row_idx - 3, col_idx - 1]
    elif hours == "11.00-12.00":
        day = df.iloc[row_idx - 4, col_idx - 1]
    elif hours == "12.00-13.00":
        day = df.iloc[row_idx - 5, col_idx - 1]
    elif hours == "14.00-15.00":
        day = df.iloc[row_idx - 7, col_idx - 1]
    elif hours == "15.00-16.00":
        day = df.iloc[row_idx - 8, col_idx - 1]
    elif hours == "16.00-17.00":
        day = df.iloc[row_idx - 9, col_idx - 1]
    elif hours == "17.00-18.00":
        day = df.iloc[row_idx - 10, col_idx - 1]

    lessons.append({'hours': hours, 'day': day})

# convert array to dataframe
df_hours = pd.DataFrame(lessons)

# Step 2: Group by the 'Date' field
grouped = df_hours.groupby('day')

# Print group
# for date, group in grouped:
#     print(f"\nDate: {date}")
#     print(group.hours)

# Create calendar
calendar = Calendar()

for date, group in grouped:
    for item in group.hours:
        start_time, end_time = item.split('-')
        
        # Convert times to datetime objects and handle fractional hours
        start_datetime = datetime.strptime(f"{date.date()} {start_time}", "%Y-%m-%d %H.%M")
        end_datetime = datetime.strptime(f"{date.date()} {end_time}", "%Y-%m-%d %H.%M")

        # Calculate the event duration
        duration = end_datetime - start_datetime

        # Create an event
        event = Event()
        event.name = "Lezione DAITA19"  
        event.begin = start_datetime 
        event.duration = duration 
        
        # Add the event to the calendar
        calendar.events.add(event)
        
# Save the calendar to an .ics file
with open('lessons.ics', 'w') as f:
    f.writelines(calendar)
