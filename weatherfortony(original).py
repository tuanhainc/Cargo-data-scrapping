import sys, time
from datetime import datetime, timedelta
from selenium import webdriver
import requests
from bs4 import BeautifulSoup
import openpyxl

#excelfile
workbook = openpyxl.Workbook()
sheet = workbook.active
headers = ["Date", "STT", "VESSEL", "NATIONALITY", "CALLSIGN", "GT", "DWT", "LOA", "DRAFT", "CARGO", "LOCTO", "TIME", "AGENT"]
sheet.append(headers)

# Date
def generate_dates(start_date_str, end_date_str):
    # Convert input date strings to datetime objects
    start_date = datetime.strptime(start_date_str, "%d/%m/%Y")
    end_date = datetime.strptime(end_date_str, "%d/%m/%Y")

    # Generate dates between start and end dates (inclusive)
    current_date = start_date
    while current_date <= end_date:
        yield current_date.strftime("%d/%m/%Y")
        current_date += timedelta(days=1)

start_date_input = input("Enter start date (DD/MM/YYYY): ")
end_date_input = input("Enter end date (DD/MM/YYYY): ")

browser = webdriver.Firefox()
browser.get('http://cangvuhanghaivungtau.gov.vn/Index.aspx?page=khddt&d=0')

for date in generate_dates(start_date_input, end_date_input):
    print("Getting data from ", date)
    original_date = datetime.strptime(date, "%d/%m/%Y")
    transformed_date = original_date.strftime("%m/%d/%Y")
    dateElem = browser.find_element_by_id('ctl03_rdpNgay_dateInput')
    dateElem.send_keys(date)
    browser.find_element_by_id('ctl03_btmSubmit').click()
    time.sleep(2)

    html_content = browser.page_source
    soup = BeautifulSoup(html_content, 'html.parser')

    for n in range (0,100):
        row_id = "ctl03_rgTauDen_ctl00__"+str(n)
        table_rows = soup.find_all('tr', id=row_id)
        distinct_set = set()
        for row in table_rows:
            cell = row.find_all('td')
            STT = cell[0].text.strip()
            if STT not in distinct_set:
                distinct_set.add(STT)
                VESSEL = cell[1].text.strip()
                NATIONALITY = cell[2].text.strip()
                CALLSIGN = cell[3].text.strip()
                GT = cell[4].text.strip()
                DWT = cell[5].text.strip()
                LOA = cell[6].text.strip()
                DRAFT = cell[7].text.strip()
                CARGO = cell[8].text.strip()
                LOCTO = cell[9].text.strip()
                TIME = cell[10].text.strip()
                AGENT = cell[13].text.strip()
                data = [transformed_date, STT, VESSEL, NATIONALITY, CALLSIGN, GT, DWT, LOA, DRAFT, CARGO, LOCTO, TIME, AGENT]
                sheet.append(data)
            else:
                break

workbook.save('data.xlsx')
