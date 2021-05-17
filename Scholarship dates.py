#!/usr/bin/env python
# coding: utf-8

# In[101]:


#https://medium.com/@vince.shields913/reading-google-sheets-into-a-pandas-dataframe-with-gspread-and-oauth2-375b932be7bf
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import numpy as np

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/calendar']

#credentials = ServiceAccountCredentials.from_json_keyfile_name(
         #r'C:\Users\timkh\OneDrive\Desktop\Programming\Projects\Scholarship dates\scholarship-dates-63e61857065e.json', scope)
credentials = ServiceAccountCredentials.from_json_keyfile_name(
         r'scholarship-dates-63e61857065e.json', scope)


gc = gspread.authorize(credentials)

wks = gc.open("Application Organizer").sheet1

data = wks.get_all_values()
headers = data.pop(0)

scholarships = pd.DataFrame(data, columns=headers)
scholarships = scholarships.replace(r'^\s*$', np.nan, regex=True) #replace whitespace with NaN
scholarships['Winner \nAnnounced'] = pd.to_datetime(scholarships['Winner \nAnnounced'])
scholarships['Deadline'] = pd.to_datetime(scholarships['Deadline'])
#scholarships.iloc[0]['Deadline']


# In[145]:


#https://developers.google.com/calendar/v3/reference/events

from apiclient import discovery

def add_day(date): #add one day
    from datetime import timedelta
    return (date+timedelta(days=1)).strftime('%Y-%m-%d')
    
def create_event(name,date): #date is yyyy-mm-dd
    event = {
  'summary': name,
  'start': {'date': date.strftime('%Y-%m-%d')},
  'end': {'date':add_day(date)}, #for some reason, using endTimeUnspecified returns a 403 forbidden error. 
 #Found it out by experimenting and removing that parameter in the "try it" section of https://developers.google.com/calendar/create-events

  'reminders': {
    'useDefault': False,
    'overrides': [{'method': 'email', 'minutes': 48 * 60}] #remind 2 days before
  },
}
    service = discovery.build('calendar', 'v3', credentials=credentials)
    return service.events().insert(calendarId='timkhaiet@gmail.com', body=event).execute()
    #events were being created in the service account's calendar if I used the 'primary' keyword as the calendarId
    #found it out by using the "try it" section of the API's 'get' method. 
    #I manually signed into my main account, and through the 'get' method, found out that the calendarId is just the email

def push_dates(row):
    if pd.isna(row['Event created']):
        if pd.isna(row['Date\n Submitted']): #if it hasn't been submitted yet, add an event name+deadline and name+winners announced
            create_event(row['Scholarship Name'] + ' Deadline', row['Deadline'])
            create_event(row['Scholarship Name'] + ' winners announced', row['Winner \nAnnounced'])
df = scholarships.copy()
df.apply(push_dates,axis=1)

last = len(wks.get_all_values()) #last row before the first empty row
#get_all_values returns a 2D list of the sheet's data. Each nested list is a row, so the length of the 2D list is the number of rows that has any data.

cell_list = wks.range(f'K2:K{last}') #Event Created column. Don't forget to change the letter K to something else if I remove some columns
#every cell has to have a value - dates cannot be empty!
for cell in cell_list:
    cell.value = 'Y'
wks.update_cells(cell_list) #change all values in the Event Created column to 'Y'

#create_event(df.iloc[0]['Scholarship Name'],df.iloc[0]['Deadline'])
#add_day(scholarships.iloc[0]['Deadline'])
#for some reason, doesn't work on pandas 1.0.2 and later - returns DLL ImportError. Still works on Jupyter though
#fixed by developers - just needed to pip install --pre pandas==1.1.0rc0




