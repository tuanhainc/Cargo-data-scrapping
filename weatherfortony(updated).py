import sys, time
from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup
import openpyxl



# Create an Excel file for data storing
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Cargo Data"
headers = ["DATE", "STT", "VESSEL", "NATIONALITY", "CALLSIGN", "DWT", "GRT", "LOA", "DRAFT", "FROM", "TO", "POB", "NOTE", "NAME", "RANK"]
sheet.append(headers)

# Create a list of dates from start date to end date
def generate_dates(start_date_str, end_date_str):
    # Convert input date strings to datetime objects
    start_date = datetime.strptime(start_date_str, "%d-%m-%Y")
    end_date = datetime.strptime(end_date_str, "%d-%m-%Y")

    # Generate dates between start and end dates (inclusive)
    current_date = start_date
    while current_date <= end_date:
        yield current_date.strftime("%d-%m-%Y")
        current_date += timedelta(days=1)

start_date_input = input("Enter start date (DD-MM-YYYY): ")
end_date_input = input("Enter end date (DD-MM-YYYY): ")

for date in generate_dates(start_date_input, end_date_input):
    print("Getting data from ", date)
    url ='https://vungtaupilot.com/Ke-Hoach-Dan-Tau/?date=%s' % (date)   # The %s acts as a placeholder

    response = requests.get(url)
    response.encoding = 'utf-8'
    response.raise_for_status()
    html_content = response.text
    # print(html_content) # (use this to see the html content)
    soup = BeautifulSoup(html_content, 'html.parser')
    # Find all table table_rows
    rows = soup.find_all('tr')

    for row in rows:
        cols = row.find_all('td')
        if len(cols) >= 14: # To skip row with missing column (total 14 columns)
            STT = cols[0].text.strip()
            VESSEL = cols[1].text.strip()
            NATIONALITY = cols[2].text.strip()
            CALLSIGN = cols[3].text.strip()
            DWT = cols[4].text.strip()
            GRT = cols[5].text.strip()
            LOA = cols[6].text.strip()
            DRAFT = cols[7].text.strip()
            FROM = cols[8].text.strip()
            TO = cols[9].text.strip()
            POB = cols[10].text.strip()
            NOTE = cols[11].text.strip()
            NAME = cols[12].text.strip()
            RANK = cols[13].text.strip()

            dataforexcelsheet = [date, STT, VESSEL, NATIONALITY, CALLSIGN, DWT, GRT, LOA, DRAFT, FROM, TO, POB, NOTE, NAME, RANK]
            sheet.append(dataforexcelsheet)
        else:
            continue # This one right here because of the headers in HTML content that has less than 14 columns


workbook.save('data.xlsx')
