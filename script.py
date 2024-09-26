from datetime import datetime, timedelta
import pandas as pd
from ics import Calendar, Event
from pytz import timezone

# Define Rome time zone (automatically handles CET/CEST based on date)
rome_tz = timezone('Europe/Rome')

# Load the Excel file
file_path = 'lessons.xlsx'
df = pd.read_excel(file_path)

# Find all cells that contain "Paolo"
matches = df.isin(['Paolo'])

lessons = []

for row_idx, col_idx in zip(*matches.to_numpy().nonzero()):
    # First cell of the same row (assuming it's column 0)
    hours = df.iloc[row_idx, 0]

    lessons_subject = df.iloc[row_idx, col_idx - 1]
    
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

    lessons.append({'hours': hours, 'day': day, "subject": lessons_subject})

# convert array to dataframe
df_hours = pd.DataFrame(lessons)

# Step 2: Group by the 'Date' field
grouped = df_hours.groupby('day')

# Print group
# for date, group in grouped:
#     print(f"\nDate: {date}")
#     print(group)

# Create calendar
calendar = Calendar()

for date, group in grouped:
    data = group.to_dict('records')

    # save the first start_time useful
    starting_time = []

    for item in data:
        start_time, end_time = item['hours'].split('-')

        if(len(data) > data.index(item)+1 and end_time in data[data.index(item)+1]['hours']):
            starting_time.append(start_time)
            continue
        
        # Convert times to datetime objects and handle fractional hours
        start_datetime = rome_tz.localize(datetime.strptime(f"{date.date()} {starting_time[0]}", "%Y-%m-%d %H.%M"))
        end_datetime = rome_tz.localize(datetime.strptime(f"{date.date()} {end_time}", "%Y-%m-%d %H.%M"))

        # Calculate the event duration
        duration = end_datetime - start_datetime

        # Create an event
        event = Event()
        event.name = item['subject'] 
        event.begin = start_datetime 
        event.duration = duration 
    
        # Add the event to the calendar
        calendar.events.add(event)

        # Clear starting time array
        starting_time = []
        
# Save the calendar to an .ics file
with open('lessons.ics', 'w') as f:
    f.writelines(calendar)
