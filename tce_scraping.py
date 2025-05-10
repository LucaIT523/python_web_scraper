import requests
import os
import re
import httplib2
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from apiclient import discovery
import oauth2client
from oauth2client import file
from oauth2client import client
from oauth2client import tools
from googleapiclient.errors import HttpError

# TCE Information
tce_login_url = "http://tceorder.ro"
tce_scraping_url = "http://tceorder.ro/orders?page="
tce_login_user = "dyna"
tce_login_pw = "tcedyna"

# TCE ORDERS Item
extracting_info = {13 : 'Nume destinatar', 14 : 'Țară destinatar', 25 : 'Descriere marfă', 27 : 'Referință expeditor', 21 : 'Valoare ramburs', }

# Google Sheet Information
CLIENT_SECRET_FILE = 'creds.json'
SpreadSheetID = '18uw230JWuCSw9wSAzfNXbneYEltITGmFG4zETvDrqo4'

# Google URL Information
APPLICATION_NAME = 'Google Sheets API Python'
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
GOOGLE_DISCOVERY = 'https://sheets.googleapis.com/$discovery/rest?version=v4'

# Extracting Information
mydata_id = []
mydata_row = []

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

def get_credentials():
    store = oauth2client.file.Storage('tokens.json')
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
    return credentials

def tce_login():
    driver.get(tce_login_url)
    driver.find_element(By.NAME, "username").send_keys(tce_login_user)
    driver.find_element(By.NAME, "password").send_keys(tce_login_pw)
    driver.find_element(By.CSS_SELECTOR, ".btn-default.btn-lg").click()
    time.sleep(1)

def get_tce_page_num(driver):
    #
    href_info = "http://tceorder.ro/orders"
    driver.get(href_info)
    first_page_soup = BeautifulSoup(driver.page_source, 'html.parser')

    for link in first_page_soup.find_all('a'):
        href_info = link.get('href')
        # print(href_info)

    str_len = len(href_info)
    total_num = href_info[str_len-2:]
    return total_num


def get_extract_all(driver, page_url):
    driver.get(page_url)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    table = soup.find('table')

    for j in table.find_all('tr')[1:]:
        th_data = j.find_all('th')
        row_id = re.search(r'>(.*?)<', str(th_data[0])).group(1)
        mydata_id.append(row_id)

        row_data = j.find_all('td')
        row_extract_info = []
        for k, col_num in enumerate(extracting_info):
            row_extract_info.append(row_data[col_num].text)

        mydata_row.append(row_extract_info)

def set_googlesheets_header(google_service):
    rangeName = 'A1:F1'
    Headers = {'values': [['ID', 'Nume destinatar', 'Țară destinatar', 'Descriere marfă' , 'Referință expeditor', 'Valoare ramburs']]}

    google_service.spreadsheets().values().update(
                    spreadsheetId=SpreadSheetID, range=rangeName,
                    valueInputOption='RAW',
                    body=Headers).execute()


def get_googlesheets_IDList(google_service):

    IDList = []
    rangeName = 'A2:A'

    result = google_service.spreadsheets().values().get(spreadsheetId=SpreadSheetID, range=rangeName).execute()
    values = result.get('values', [])
    for row in values:
        IDList.append(row[0])

    return IDList

def set_googlesheets_data(google_service, ID, data):
    rangeName = 'A1:F1'
    row_data = {'values': [[ID, data[0], data[1], data[2] , data[3], data[4]]]}

    google_service.spreadsheets().values().append(
                    spreadsheetId=SpreadSheetID, range=rangeName,
                    valueInputOption='RAW',
                    body=row_data).execute()


if __name__ == '__main__':
    # Init webdriver
    driver = webdriver.Chrome()
    driver.minimize_window()
    print('OK ... Init Webdriver')

    # login to TCE
    tce_login()
    print('OK ... Login to TCE Order')

    # web scraping
    total_num = get_tce_page_num(driver)
    for num in range(1, int(total_num) + 1):
        url = tce_scraping_url + str(num)
        # page
        get_extract_all(driver, url)
        print(f'extracted page - {num}')

    # TEST CODE
    # get_extract_all(driver, 'D:/zzz/TCE Worldwide Web Services.html')
    # # test print
    # for num, id in enumerate(mydata_id):
    #     print(f'{id} , {mydata_row[num]}')
    # print(f'total row number = {len(mydata_id)}')

    # Close webdriver
    driver.quit()

    try:
        # Init of Google Sheets
        credentials = get_credentials()
        http = credentials.authorize(httplib2.Http())
        google_service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=GOOGLE_DISCOVERY)
        print('OK ... Authentication Google Sheets')


        # Init Header of Google Sheets
        set_googlesheets_header(google_service)

        # Get ID List of Google Sheets
        IDList = get_googlesheets_IDList(google_service)
        # check ID
        for num, tce_id in enumerate(mydata_id):
            bFind = False
            for num_g, sheets_id in enumerate(IDList):
                if tce_id == sheets_id :
                    bFind = True
                    break
            if bFind == False:
                # Write into Google Sheets
                set_googlesheets_data(google_service, tce_id, mydata_row[num])
                print(f'append data({tce_id} , {mydata_row[num]}) into google sheets')
                time.sleep(1)

        print('OK ... Update Google Sheets')

    except HttpError as err:
        print(f'Error ... Google Sheets ({err})')

