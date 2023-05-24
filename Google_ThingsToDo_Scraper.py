from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService 
from subprocess import CREATE_NO_WINDOW 
from datetime import datetime
import time
import warnings
import pycountry
import os
import re
import sys
import unidecode
import sys
# google API dependencies
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
# UI dependencies
import tkinter
from tkinter import ttk
from ttkthemes import ThemedTk
import threading
from tkinter import messagebox
import tkinter.scrolledtext as ScrolledText
import logging
import pickle

# disabling warnings
warnings.filterwarnings('ignore')
# disabling google API messages
logging.getLogger("googleapiclient").setLevel(logging.WARNING)

def get_Google_API_creds():
    
    global running, root, dummy_driver

    dummy_driver = get_url('https://www.google.com', dummy_driver)

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']

    # API configuration
    credentials = None
    if os.path.exists('token.json'):
        credentials = Credentials.from_authorized_user_file('token.json', SCOPES)
    try:
        # If there are no (valid) credentials available, let the user log in.
        connected = False
        if not credentials or not credentials.valid:
            try:
                if credentials and credentials.expired and credentials.refresh_token:
                    dummy_driver = get_url('https://www.google.com', dummy_driver)
                    credentials.refresh(Request())
                    connected = True
            except:
                pass
            if not connected:
                dummy_driver = get_url('https://www.google.com', dummy_driver)
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                ports = [8080, 7100, 6100, 5100, 4100]               
                for port in ports:
                    try:
                        dummy_driver = get_url('https://www.google.com', dummy_driver)
                        credentials = flow.run_local_server(port=port, access_type='offline', include_granted_scopes='true')
                        connected = True
                        break
                    except Exception as err:
                        output_msg('The below error occurred while authenicating the Google API ....', 1)
                        output_msg(str(err), 0)
                if not connected:
                    output_msg('Failed to authenicate the Google API, could not find an empty port, exiting ....', 1)
                    sys.exit()
                # Save the credentials for the next run
                with open('token.json', 'w') as token:
                    token.write(credentials.to_json())

    except Exception as err:
        output_msg('The below error occurred while authenicating the Google API ....', 1)
        output_msg(str(err), 0)
        sys.exit()

    time.sleep(2)

    return credentials


def process_sheet(url):

    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    credentials = get_Google_API_creds()
    sheets_service = build('sheets', 'v4', credentials=credentials, cache_discovery=False)

    # getting all the destinations infor from the template sheet
    dummy_driver = get_url('https://www.google.com', dummy_driver)
    sheet_row_count = get_sheet_row_count(sheets_service, url)
    dummy_driver = get_url('https://www.google.com', dummy_driver)
    rows = read_range(sheets_service, url, sheet_row_count)

    # trimming the list by the first empty row in the sheet
    dests = []
    for row in rows:
        try:
            if row[0] != '' and row[0] != 'nan' and row[0] != None:
                if len(row) == 3:
                    dests.append((row[0], row[1], row[2]))
                elif len(row) == 2:
                    dests.append((row[0], row[1], ''))
                else:
                    dests.append((row[0], '', ''))
            else:
                break
        except:
            break

    return dests

def get_url(url, driver):

    err_msg = False
    while True:
        try:
            driver.get(url)
            if err_msg:
                output_msg('The internet connection is restored, resuming the bot ...', 1)
                output_msg('-'*75, 0)
            break
        except:
            if not err_msg:
                output_msg('Warning: failed to load a page, waiting for the internet connection ...', 1)
                err_msg = True
            time.sleep(60)

    return driver

def initialize_bot():

    # Setting up chrome driver for the bot
    chrome_options  = webdriver.ChromeOptions()
    # suppressing output messages from the driver
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--window-size=1920,1080')
    # adding user agents
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument("--incognito")
    # running the driver with no browser window
    chrome_options.add_argument('--headless')
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    chrome_service.creationflags = CREATE_NO_WINDOW
    # configuring the driver
    driver = webdriver.Chrome(driver_path, options=chrome_options, service=chrome_service)
    driver.set_page_load_timeout(60)
    driver.maximize_window()

    return driver

def get_sheet_row_count(sheets_service, url):

    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    spreadsheet_id = url.split('/')[5]
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    dummy_driver = get_url('https://www.google.com', dummy_driver)
    sheets = sheet_metadata.get('sheets', '')
    for sheet in sheets:
        name_sheet = sheet["properties"]["title"]
        if name_sheet == "Sheet1":
            return sheet["properties"]["gridProperties"]["rowCount"]


def read_range(sheets_service, url, row_count_max):

    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    range_name = f'Sheet1!A2:C{row_count_max}'
    spreadsheet_id = url.split('/')[5]
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_name).execute()
    rows = result.get('values', [])
    return rows


def write_status(row_number, value, url):

    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    credentials = get_Google_API_creds()
    sheets_service = build('sheets', 'v4', credentials=credentials, cache_discovery=False)
    spreadsheet_id = url.split('/')[5]
    range_name = f'Sheet1!B{row_number}:B{row_number}'
    value_input_option = 'USER_ENTERED'
    body = {'values': [[value]]}
    dummy_driver = get_url('https://www.google.com', dummy_driver)
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=range_name,
        valueInputOption=value_input_option, body=body).execute()

def output_msg(msg, newline):

    if newline == 1:
        logging.info('-'*75)
    logging.info(msg)
    stamp = datetime.now().strftime("%Y-%m-%d %I:%M %p")
    with open('session_log.log', 'a', newline='', encoding='UTF-8') as f:
        if newline == 1:
            f.write('-'*75)
            f.write('\n')
        f.write(f'{stamp} - {msg}\n')

def search_destinations(driver, dest, limit, url, k, folder):

    global dummy_driver

    start_time = time.time()
    doc_id, nwiki, nflickr, ngoogle_official, ngoogle_user = 0, 0, 0, 0, 0
    country = 'Other'

    # navigating to google travel
    driver = get_url('https://www.google.com/travel/things-to-do/see-all', driver)
    time.sleep(2)
    # sending dummy text to initialize the textbox class
    search = wait(driver, 1).until(EC.presence_of_element_located((By.XPATH, "//input[@class='II2One j0Ppje zmMKJ LbIaRd']")))
    dummy_driver = get_url('https://www.google.com', dummy_driver)
    search.send_keys(' ')
    # sending the destination
    search = wait(driver, 1).until(EC.presence_of_element_located((By.XPATH, "//input[@class='II2One j0Ppje zmMKJ LbIaRd']")))
    dummy_driver = get_url('https://www.google.com', dummy_driver)
    search.clear()
    search.send_keys(dest[0])
    time.sleep(1)
    search.send_keys(Keys.ENTER)
    time.sleep(5)      
    dummy_driver = get_url('https://www.google.com', dummy_driver)
    link = driver.current_url
    # condition for no results in Google to do 
    if 'https://www.google.com/travel/things-to-do' not in link:
        write_status(k+2, 'No Results', url)
        output_msg(f'No results were found for destination: {dest[0].title()}', 1)
      
        return ''

    output_msg(f'Scraping information for destination {k+1}: {dest[0].title()}', 1)

    try:
        headers = wait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, "//h2[@class='osfY2d HVJNrc']")))
        destination = False
        for header in headers:
            if header.text == 'Top sights':
                destination = True
                break

        if not destination:
            # no top sights, then it is an attraction not a destination
            write_status(k+2, 'Error: attraction and not a destination', url)
            output_msg(f'{dest[0].title()} is an attraction not a destination, skipping ...', 1)
            # output status to the log file
            time.sleep(3)
            return 'Error: attraction and not a destination'
    except:
        # no top sights, then it is an attraction not a destination
        write_status(k+2, 'Error: attraction and not a destination', url)
        output_msg(f'{dest[0].title()} is an attraction not a destination, skipping ...', 1)
        time.sleep(3)
        return 'Error: attraction and not a destination'

    # getting the top attractions
    sights_menu = wait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[@class='XzK3Bf' and @aria-label='Top sights']")))
    top_sights = wait(sights_menu, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.f4hh3d"))) 
    nsights_top = len(top_sights)

    # reordering the top reviews
    top_reviews = {}
    for sight in top_sights:
        text = sight.text
        name = sight.text.split('\n')[0]
        try:
            text = wait(sight, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.jdzyld.XLC8M"))).text 
            nrev = text[text.find('(')+1:text.find(')')].replace(',', '')
            nrev = int(nrev)
        except:
            nrev = 0

        if nrev > 0:
            if top_reviews.get(nrev, -1) == -1:
                top_reviews[nrev] = [name]
            else:
                top_reviews[nrev].append(name)

    # getting the number of top attractions with reviews
    keys = list(top_reviews.keys())
    ntop = 0
    for key in keys:
        for name in top_reviews[key]:
            ntop += 1

    # condition to check if there is a need to scrape non top attractions
    if limit > ntop:
        # clicking on "see all top sights" button
        buttons = wait(driver, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "button")))
        for button in buttons:
            if 'See all top sights' in button.text:
                dummy_driver = get_url('https://www.google.com', dummy_driver)
                driver.execute_script("arguments[0].click();", button)
                time.sleep(3)
                break
        # looping over all the sights
        sights_menu = wait(driver, 1).until(EC.presence_of_element_located((By.XPATH, "//div[@class='XzK3Bf']")))
        sights = wait(sights_menu, 1).until(EC.presence_of_all_elements_located((By.XPATH, "//div[@class='f4hh3d']")))
        nsights = len(sights)
        # check the existance of the additonal sights
        if nsights > nsights_top:
            # looping through non top sights only to order them by the number of reviews
            reviews = {}
            for sight in sights[nsights_top:]:
                #iterating 2 times in case error occurred
                for _ in range(2):
                    try:
                        try:
                            text = sight.text
                            name = sight.text.split('\n')[0]
                            text = wait(sight, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.jdzyld.XLC8M"))).text 
                            nrev = text[text.find('(')+1:text.find(')')].replace(',', '')
                            nrev = int(nrev)
                        except:
                            # for sights with no reviews
                            nrev = 0

                        # only considering sights with reviews
                        if nrev > 0:
                            if reviews.get(nrev, -1) == -1:
                                reviews[nrev] = [name]
                            else:
                                reviews[nrev].append(name)
                        break
                    except:
                        output_msg("Warning: Page didn't reload correctly, refreshing the page ...", 1)
                        driver.refresh()
                        time.sleep(1)
                        continue

            # ordering the sights in Ascending order
            keys = list(reviews.keys())
            keys.sort(reverse=False)
            n = 0
            # getting the number of attractions
            ind = ntop
            done = False
            for key in keys:
                if done: break
                for _ in reviews[key]:
                    ind += 1
                    if ind == limit:
                        done = True
                        break
            
            total = ind
            # in case no attractions with reviews are available
            if total == 0:
                write_status(k+2, 'No attractions With Reviews', url)
                output_msg(f'No attractions With Reviews were found for destination: {dest[0].title()}, skipping ...', 1)
               
                return '' 
            
            # creating google doc for the destination
            for _ in range(5):
                doc_id, end_ind = create_google_doc(dest[0].title(), dest[2], ind, folder)
                if doc_id != '-1':
                    break

            if doc_id == '-1':
                write_status(k+2, 'Failure in creating Google doc', url)
                return 'Failure in creating Google doc' 

            if ind > 1:
                title = f"\nThe {ind} Most Popular Things To Do In {dest[0].title()}"
            else:
                title = f"\nThe Most Popular Thing To Do In {dest[0].title()}"

            end_ind = add_title_to_google_doc(title, doc_id, end_ind)

            if end_ind == -1:
                write_status(k+2, 'Failure in adding title to Google doc', url)
                return 'Failure in adding title to Google doc'

            # considering the remaining attractions within the limit
            n = ntop
            done = False
            for key in keys:
                if done: break
                for name in reviews[key]:
                    found = False
                    if done: break
                    success = False
                    for _ in range(1):
                        if success: break
                        # closing all opened tab except the first one
                        for handle in driver.window_handles[1:]:
                            dummy_driver = get_url('https://www.google.com', dummy_driver)
                            driver.switch_to.window(handle)
                            driver.close()
                        time.sleep(2)
                        dummy_driver = get_url('https://www.google.com', dummy_driver)
                        driver.switch_to.window(driver.window_handles[0])
                        sights_menu = wait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//div[@class='XzK3Bf']")))
                        sights = wait(sights_menu, 3).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.f4hh3d")))
                        for sight in sights[nsights_top:]:
                            if sight.text.split('\n')[0] == name:
                                try:
                                    found = True
                                    # name of the attraction
                                    name = sight.text.split('\n')[0]
                                    # for handling different attr with the same name
                                    try:
                                        # rating of the attraction
                                        output_msg('Scraping attraction rating ...', 0)
                                        rating = wait(sight, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.KFi5wf.lA0BZ"))).text
                                        # no of reviews of the attraction
                                        output_msg('Scraping attraction number of reviews ...', 0)
                                        text = wait(sight, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.jdzyld.XLC8M"))).text 
                                        nrev = text[text.find('(')+1:text.find(')')].replace(',', '')
                                        nrev = int(nrev)
                                        if nrev != key:
                                            continue
                                    except:
                                        continue

                                    # clicking on the sight card
                                    button = wait(sight, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.Ld2paf")))
                                    dummy_driver = get_url('https://www.google.com', dummy_driver)
                                    driver.execute_script("arguments[0].click();", button) 
                                    time.sleep(2)
                                    # clicking on "Web results about this place" button
                                    sight_div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.U4rdx")))
                                    button = wait(sight_div, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))[-1]
                                    dummy_driver = get_url('https://www.google.com', dummy_driver)
                                    driver.execute_script("arguments[0].click();", button) 
                                    time.sleep(1)
                                    # switching to the next tab of the browser window
                                    dummy_driver = get_url('https://www.google.com', dummy_driver)
                                    driver.switch_to.window(driver.window_handles[1])
                                    time.sleep(1)
                                    output_msg('Scraping attraction address, tel and website ...', 0)
                                    attr = get_attraction_info(driver, name, rating, nrev)
                                    # closing all opened tab except the first one
                                    for handle in driver.window_handles[1:]:
                                        dummy_driver = get_url('https://www.google.com', dummy_driver)
                                        driver.switch_to.window(handle)
                                        driver.close()
                                    time.sleep(2)
                                    dummy_driver = get_url('https://www.google.com', dummy_driver)
                                    driver.switch_to.window(driver.window_handles[0])
                                    try:
                                        output_msg('Scraping attraction image ...', 0)
                                        nwiki, nflickr, ngoogle_official, ngoogle_user = get_attraction_image(driver, attr, dest[0].title(), nwiki, nflickr, ngoogle_official, ngoogle_user)  
                                    except:
                                        # closing all opened tab except the first one
                                        output_msg('Warning: Failed to scrape the attraction image', 0)
                                        for handle in driver.window_handles[1:]:
                                            dummy_driver = get_url('https://www.google.com', dummy_driver)
                                            driver.switch_to.window(handle)
                                            driver.close()
                                            dummy_driver = get_url('https://www.google.com', dummy_driver)
                                        driver.switch_to.window(driver.window_handles[0])
                                        time.sleep(2)
                                        attr['image'] = ''
                                        attr['image_name'] = ''
                                        attr['image_url'] = ''
                                        attr['credit_url'] = ''
                                        attr['credit_name'] = ''
                                        attr['license_url'] = ''
                                        attr['license_name'] = ''
                                    if len(attr['country']) > 0 and country == 'Other':
                                        country = attr['country']
                            ################################################
                                    output_msg(f"Attraction Number: {ind}/{total}", 0)
                                    output_msg(f"Destination Name: {dest[0].title()}", 0)
                                    output_msg(f"Attraction Name: {attr['name']}", 0)
                                    output_msg(f"Attraction Rating: {attr['rating']}", 0)
                                    output_msg(f"Attraction No of Reviews: {attr['reviews']}", 0)
                                    if len(attr['website']) > 0:
                                        output_msg(f"Attraction Website: {attr['website']}", 0)
                                    if len(attr['address']) > 0:
                                        output_msg(f"Attraction Address: {attr['address']}", 0)
                                    if len(attr['phone']) > 0:
                                        output_msg(f"Attraction Phone Number: {attr['phone']}", 0)
                                    if len(attr['closed']) > 0:
                                        output_msg(f"Attraction Opening Status: {attr['closed']}", 0)                        
                                    if len(attr['image_name']) > 0:
                                        output_msg(f"Attraction Image Source: {attr['image_name']}", 0) 
                                    # exporting attraction to google doc
                                    end_ind = export_ensight(doc_id, attr, ind, end_ind)

                                    if end_ind == -1:
                                        output_msg(f"Exporting the attraction to Google doc", 1) 
                                        write_status(k+2, 'Failure in exporting attraction to Google doc', url)
                                        return 'Failure in exporting attraction to Google doc' 
                                    if not isinstance(end_ind, int): 
                                        return end_ind
                                    ind -= 1   
                                    success = True
                                    n += 1
                                    if n == limit:
                                        done = True
                                    break
                                except:
                                    output_msg("Warning: Page didn't reload correctly, refreshing the page ...", 1)
                                    dummy_driver = get_url('https://www.google.com', dummy_driver)
                                    driver.refresh()
                                    time.sleep(2)
                                    break
                    if not found:
                        output_msg(f"Warning: could not find attraction: {name}", 0)

            if ntop > 0 and total > 14:
                # exporting top sights title
                title = f"\nOur Top picks"

                end_ind = add_title_to_google_doc(title, doc_id, end_ind)

                if end_ind == -1:
                    write_status(k+2, 'Failure in adding title to Google doc', url)
                    return 'Failure in adding title to Google doc'

    # top sights export
    # if no document created yet
    if doc_id == 0:
        total = ntop
        ind = ntop
        if ind == 0:
            write_status(k+2, 'No attractions With Reviews', url)
            output_msg(f'No attractions With Reviews were found for destination: {dest[0].title()}, skipping ...', 1)

            return '' 
        # creating google doc for the destination
        for _ in range(5):
            doc_id, end_ind = create_google_doc(dest[0].title(), dest[2], ind, folder)
            if doc_id != '-1':
                break

        if doc_id == '-1':
            write_status(k+2, 'Failure in exporting Google doc', url)
            return 'Failure in exporting Google doc' 

        if ind > 1:
            title = f"\nThe {ind} Most Popular Things To Do In {dest[0].title()}"
        else:
            title = f"\nThe Most Popular Thing To Do In {dest[0].title()}"

        end_ind = add_title_to_google_doc(title, doc_id, end_ind)

        if end_ind == -1:
            write_status(k+2, 'Failure in adding title to Google doc', url)
            return 'Failure in adding title to Google doc'

        try:
            sights_menu = wait(driver, 1).until(EC.presence_of_element_located((By.XPATH, "//div[@class='XzK3Bf' and @aria-label='Top sights']")))
        except:
            # if there are no top sights
            sights_menu = wait(driver, 1).until(EC.presence_of_element_located((By.XPATH, "//div[@class='XzK3Bf']")))
        sights = wait(sights_menu, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.f4hh3d")))

    # ordering the sights in Ascending order
    keys = list(top_reviews.keys())
    keys.sort(reverse=False)

    for key in keys:
        for name in top_reviews[key]:
            success = False
            found = False
            for _ in range(1):
                if success: break
                # closing all opened tab except the first one
                for handle in driver.window_handles[1:]:
                    dummy_driver = get_url('https://www.google.com', dummy_driver)
                    driver.switch_to.window(handle)
                    driver.close()
                time.sleep(2)
                dummy_driver = get_url('https://www.google.com', dummy_driver)
                driver.switch_to.window(driver.window_handles[0])
                try:
                    sights_menu = wait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//div[@class='XzK3Bf' and @aria-label='Top sights']")))
                except:
                    # if there are no top sights
                    sights_menu = wait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//div[@class='XzK3Bf']")))
                sights = wait(sights_menu, 3).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.f4hh3d")))
                for sight in sights[:nsights_top+1]:
                    if sight.text.split('\n')[0] == name:
                        try:
                            found = True
                            # name of the attraction
                            name = sight.text.split('\n')[0]
                            # for handling different attr with the same name
                            try:
                                # rating of the attraction
                                output_msg('Scraping attraction rating ...', 0)
                                rating = wait(sight, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.KFi5wf.lA0BZ"))).text
                                # no of reviews of the attraction
                                output_msg('Scraping attraction number of reviews ...', 0)
                                text = wait(sight, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.jdzyld.XLC8M"))).text 
                                nrev = text[text.find('(')+1:text.find(')')].replace(',', '')
                                nrev = int(nrev)
                                if nrev != key:
                                    continue
                            except:
                                continue

                            # clicking on the sight card
                            button = wait(sight, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.Ld2paf")))
                            dummy_driver = get_url('https://www.google.com', dummy_driver)
                            driver.execute_script("arguments[0].click();", button) 
                            time.sleep(1)
                            # clicking on "Web results about this place" button 
                            sight_div = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.U4rdx")))
                            button = wait(sight_div, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))[-1]
                            dummy_driver = get_url('https://www.google.com', dummy_driver)
                            driver.execute_script("arguments[0].click();", button) 
                            time.sleep(1)
                            # switching to the next tab of the browser window
                            dummy_driver = get_url('https://www.google.com', dummy_driver)
                            driver.switch_to.window(driver.window_handles[1])
                            time.sleep(1)
                            output_msg('Scraping attraction address, tel and website ...', 0)
                            attr = get_attraction_info(driver, name, rating, nrev)
                            # closing all opened tab except the first one
                            for handle in driver.window_handles[1:]:
                                dummy_driver = get_url('https://www.google.com', dummy_driver)
                                driver.switch_to.window(handle)
                                driver.close()
                            time.sleep(2)
                            dummy_driver = get_url('https://www.google.com', dummy_driver)
                            driver.switch_to.window(driver.window_handles[0])
                            time.sleep(2)
                            try:
                                output_msg('Scraping attraction image ...', 0)
                                nwiki, nflickr, ngoogle_official, ngoogle_user = get_attraction_image(driver, attr, dest[0].title(), nwiki, nflickr, ngoogle_official, ngoogle_user)  
                            except:
                                # closing all opened tab except the first one
                                output_msg('Warning: Failed to scrape the attraction image', 0)
                                for handle in driver.window_handles[1:]:
                                    dummy_driver = get_url('https://www.google.com', dummy_driver)
                                    driver.switch_to.window(handle)
                                    driver.close()
                                time.sleep(2)
                                dummy_driver = get_url('https://www.google.com', dummy_driver)
                                driver.switch_to.window(driver.window_handles[0])   
                                attr['image'] = ''
                                attr['image_name'] = ''
                                attr['image_url'] = ''
                                attr['credit_url'] = ''
                                attr['credit_name'] = ''
                                attr['license_url'] = ''
                                attr['license_name'] = ''
                            if len(attr['country']) > 0 and country == 'Other':
                                country = attr['country']
                ###################################################
                            output_msg(f"Attraction Number: {ind}/{total}", 0)
                            output_msg(f"Destination Name: {dest[0].title()}", 0)
                            output_msg(f"Attraction Name: {attr['name']}", 0)
                            output_msg(f"Attraction Rating: {attr['rating']}", 0)
                            output_msg(f"Attraction No of Reviews: {attr['reviews']}", 0)
                            if len(attr['website']) > 0:
                                output_msg(f"Attraction Website: {attr['website']}", 0)
                            if len(attr['address']) > 0:
                                output_msg(f"Attraction Address: {attr['address']}", 0)
                            if len(attr['phone']) > 0:
                                output_msg(f"Attraction Phone Number: {attr['phone']}", 0)
                            if len(attr['closed']) > 0:
                                output_msg(f"Attraction Opening Status: {attr['closed']}", 0)    
                            output_msg(f"Attraction country: {attr['country']}", 0)
                            if len(attr['image_name']) > 0:
                                output_msg(f"Attraction Image Source: {attr['image_name']}", 0)
                            # exporting the attraction to google doc
                            output_msg(f"Exporting the attraction to Google doc", 1)
                            end_ind = export_ensight(doc_id, attr, ind,end_ind)

                            if end_ind == -1:
                                write_status(k+2, 'Failure in exporting attraction to Google doc', url)
                                return 'Failure in exporting attraction to Google doc' 

                            if not isinstance(end_ind, int): 
                                return end_ind
                            ind -= 1
                            success = True
                            break
                        except:
                            output_msg("Warning: Page didn't reload correctly, refreshing the page ...", 1)
                            dummy_driver = get_url('https://www.google.com', dummy_driver)
                            driver.refresh()
                            time.sleep(1)
                            break
            if not found:
                output_msg(f"Warning: could not find attraction: {name}", 1)

    # checking if the attractions index is ended at 1
    if ind != 0:
        return 'Missing Attractions in the Google doc' 

    # adding the final paragraph
    # new line for the next attraction
    end_text = True
    text = '\n\n\n'
    start_ind = end_ind 
    end_ind =  start_ind + len(text) 
    success = False
    for _ in range(10):
        try:
            add_text_via_API(doc_id, text, start_ind, end_ind, 'Martel', 13, True, False, 0, 0, 0)  
            success = True
            break
        except:
            start_ind -= 1

    if not success:
        write_status(k+2, 'Failure in exporting final text to Google doc', url)
        output_msg('Failure in exporting final text to Google doc, please add manually', 1)
        end_text = False


    text = 'Final Words:\n'
    start_ind = end_ind
    end_ind = start_ind + len(text)
    success = False
    for _ in range(10):
        try:
            add_heading_via_API(doc_id, text, start_ind, end_ind, 'Martel', 13, True, False, 0, 0, 0, 'HEADING_3') 
            success = True
            break
        except:
            start_ind -= 1

    if not success:
        write_status(k+2, 'Failure in exporting final text to Google doc', url)
        output_msg('Failure in exporting final text to Google doc, please add manually', 1)
       
        end_text = False

    text = f'Thank you for reading our list of the best things to do in {dest[0].title()}! We hope it helped you plan the perfect itinerary for your next trip to this fantastic destination. Happy travels!\n'
    start_ind = end_ind
    end_ind = start_ind + len(text)
    success = False
    for _ in range(10):
        try:
            add_text_via_API(doc_id, text, start_ind, end_ind, 'Martel', 13, False, False, 0, 0, 0)
            success = True
            break
        except:
            start_ind -= 1

    if not success:
        write_status(k+2, 'Failure in exporting final text to Google doc', url)
        output_msg('Failure in exporting final text to Google doc, please add manually', 1)
        
        end_text = False

    text = 'Join the TouristWire community\n'
    start_ind = end_ind
    end_ind = start_ind + len(text)
    success = False
    for _ in range(10):
        try:
            add_heading_via_API(doc_id, text, start_ind, end_ind, 'Martel', 13, True, False, 0, 0, 0, 'HEADING_3')
            success = True
            break
        except:
            start_ind -= 1

    if not success:
        write_status(k+2, 'Failure in exporting final text to Google doc', url)
        output_msg('Failure in exporting final text to Google doc, please add manually', 1)
        
        end_text = False

    text = 'At TouristWire, our team works tirelessly to bring you the most comprehensive itineraries and reviews to help you plan your next trip. Don’t miss out, join our mailing list and our reviews delivered directly to your mailbox!\n'
    start_ind = end_ind
    end_ind = start_ind + len(text)
    success = False
    for _ in range(10):
        try:
            add_text_via_API(doc_id, text, start_ind, end_ind, 'Martel', 13, False, False, 0, 0, 0)
            success = True
            break
        except:
            start_ind -= 1

    if not success:
        write_status(k+2, 'Failure in exporting final text to Google doc', url)
        output_msg('Failure in exporting final text to Google doc, please add manually', 1)
        
        end_text = False
    # moving the destination document to the right location on the drive
    move_doc(doc_id, country, dest[0].title(), folder)
    if end_text:
        # report success for the destination
        write_status(k+2, 'Scraped', url)
        output_msg(f'{dest[0].title()} is processed successfully!', 1)
    output_msg(f'Number of Google Maps Official Images: {ngoogle_official}', 0)
    output_msg(f'Number of Google Maps User Images: {ngoogle_user}', 0)
    output_msg(f'Number of Wikimedia Images: {nwiki}', 0)
    output_msg(f'Number of Flickr Images: {nflickr}', 0)
    dest_time = round((time.time() - start_time)/60, 2)
    output_msg(f'{dest[0].title()} is completed in: {dest_time} mins', 1)
    
    return ''

def add_text_via_API(doc_id, text, start_ind, end_ind, font, size, bold, italic, blue, green, red):

    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    credentials = get_Google_API_creds()
    service = build('docs', 'v1', credentials=credentials, cache_discovery=False)

    requests = [{'insertText': {'location': {'index': start_ind},'text': text}}, {"updateParagraphStyle": {"range": {"startIndex": start_ind,"endIndex": end_ind},"paragraphStyle": {"alignment": "START"},"fields": "alignment"}}, {'updateTextStyle': {'range': {'startIndex': start_ind,'endIndex':end_ind},'textStyle': {'weightedFontFamily': {'fontFamily': font}, 'bold': bold, 'italic': italic,'fontSize': {'magnitude': size,'unit': 'PT'}, 'foregroundColor': {'color': {'rgbColor': {'blue': blue,'green': green,'red': red}}}},'fields':'foregroundColor, bold, italic, weightedFontFamily, fontSize'}}]
    dummy_driver = get_url('https://www.google.com', dummy_driver)
    service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
    
def add_heading_via_API(doc_id, text, start_ind, end_ind, font, size, bold, italic, blue, green, red, heading):

    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    credentials = get_Google_API_creds()
    service = build('docs', 'v1', credentials=credentials, cache_discovery=False)

    requests = [{'insertText': {'location': {'index': start_ind},'text': text}}, {"updateParagraphStyle": {"range": {"startIndex": start_ind,"endIndex": end_ind},"paragraphStyle": {"namedStyleType": heading, "alignment": "START"},"fields": "alignment, namedStyleType"}}, {'updateTextStyle': {'range': {'startIndex': start_ind,'endIndex':end_ind},'textStyle': {'weightedFontFamily': {'fontFamily': font}, 'bold': bold, 'italic': italic,'fontSize': {'magnitude': size,'unit': 'PT'}, 'foregroundColor': {'color': {'rgbColor': {'blue': blue,'green': green,'red': red}}}},'fields':'foregroundColor, bold, italic, weightedFontFamily, fontSize'}}]
    dummy_driver = get_url('https://www.google.com', dummy_driver)
    service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
    
def add_hyperlink_via_API(doc_id, text, link, start_ind, end_ind, style_start_ind, style_end_ind, font, size, bold, italic, blue, green, red):

    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    credentials = get_Google_API_creds()
    service = build('docs', 'v1', credentials=credentials, cache_discovery=False)

    requests = [{'insertText': {'location': {'index': start_ind},'text': text}}, {'updateTextStyle': {'range': {'startIndex': start_ind,'endIndex':end_ind},'textStyle': {'weightedFontFamily': {'fontFamily': font}, 'bold': bold, 'italic': italic,'fontSize': {'magnitude': size,'unit': 'PT'}, 'foregroundColor': {'color': {'rgbColor': {'blue': blue,'green': green,'red': red}}}},'fields':'foregroundColor, bold, italic, weightedFontFamily, fontSize'}}, {"updateParagraphStyle": {"range": {"startIndex": start_ind,"endIndex": end_ind},"paragraphStyle": {"alignment": "START"}, "fields": "alignment"}}, {'updateTextStyle': {'range': {'startIndex': style_start_ind, 'endIndex': style_end_ind} ,'textStyle': {'link':{'url': link}},'fields': 'link'}}]
    dummy_driver = get_url('https://www.google.com', dummy_driver)
    service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()


def create_google_doc(dest, text, natt, folder_id):
    
    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    credentials = get_Google_API_creds()
    output_msg(f'Creating Google doc for {dest} ...', 1)

    drive_service = build('drive', 'v3', credentials=credentials, cache_discovery=False)

    try:
        title = f"{natt} Best and Fun Things To Do In {dest}"
        body = {'name': title, 'mimeType': 'application/vnd.google-apps.document', 'parents': [folder_id]}
        # checking destinations files in the input folder
        attr_files = {}
        query = f'mimeType != "application/vnd.google-apps.folder" and trashed = false and name = "{title}" and "{folder_id}" in parents'
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        response = drive_service.files().list(q = query).execute()
        for attr_file in response.get('files', []):
            attr_files[attr_file.get('name')] = attr_file.get('id')

        # if there is an old destination file with the same name then remove it to be replaced with the updated one
        if title in attr_files.keys():
            attr_file_id = attr_files[title]
            dummy_driver = get_url('https://www.google.com', dummy_driver)
            drive_service.files().delete(fileId=attr_file_id).execute()

        #creating the new document
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        doc = drive_service.files().create(body=body, fields='id').execute()
        doc_id = doc.get('id')
        # sharing the doc
        request = {'role':'writer', 'type':'anyone'}
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        drive_service.permissions().create(fileId=doc_id, body=request).execute()
        #link_response = drive_service.files().get(fileId=doc_id, fields='webViewLink').execute()
        #link = link_response['webViewLink']
    except HttpError as err:
        output_msg('The following error occurred while creating the Google doc', 1)
        output_msg(err, 0)

    try:
        # adding contents of the intro
        start_ind = 1
        end_ind = 1
        if len(text) > 0 and text != 'nan':
            end_ind =  len(text)
            text = text + '\n\n'
            success = False
            for _ in range(10):
                try:
                    add_text_via_API(doc_id, text, start_ind, end_ind, 'Martel', 13, False, False, 0, 0, 0)
                    success = True
                    break
                except:
                    start_ind -= 1

            if not success:
                output_msg(f'Failed to export the intro text to Google doc for destination: {dest}', 1)

                return '-1', '-1'
                


            # new lines for the title
            text = '\n\n'
            start_ind = end_ind 
            end_ind =  start_ind + len(text) 
            for _ in range(10):
                try:
                    add_text_via_API(doc_id, text, start_ind, end_ind, 'Martel', 20, False, False, 0, 0, 0)
                    break
                except:
                    start_ind -= 1


    except Exception as err:
        # failure in exporting the data to google doc
        output_msg('The following error occurred while creating the Google doc', 1)
       
    return doc_id, end_ind

def add_title_to_google_doc(text, doc_id, start_ind):

    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    try:
        # title 
        text = text + '\n\n'
        start_ind = start_ind
        end_ind =  start_ind + len(text) 
        success = False
        for _ in range(10):
            try:
                add_heading_via_API(doc_id, text, start_ind, end_ind, 'Reem Kufi', 30, True, False, 0.51, 0.23, 0.25, 'HEADING_2')
                success = True
                break
            except:
                start_ind -= 1

        if not success:
            output_msg('Failed to export the title text to Google doc', 1)
           
            return -1

    except Exception as err:
        # failure in exporting the data to google doc
        output_msg('The following error occurred while adding a title to the Google doc ...', 1)
        output_msg(err, 0)
        
    return end_ind


def export_ensight(doc_id, attr, attr_id, start_id):
    
    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    credentials = get_Google_API_creds()
    service = build('docs', 'v1', credentials=credentials, cache_discovery=False)

    try:
        # attraction name
        text =  str(attr_id) + '. ' + str(attr['name']) + '\n'
        if len(text) > 0:
            start_ind = start_id 
            end_ind =  start_ind + len(text)
            success = False
            for _ in range(10):
                try:
                    add_heading_via_API(doc_id, text, start_ind, end_ind, 'Reem Kufi', 30, True, False, 0.51, 0.23, 0.25, 'HEADING_3')
                    success = True
                    break
                except:
                    start_ind -= 1

            if not success:
                output_msg(f"Failed to export the attraction name '{attr['name']}' to Google doc", 0)
                
                return -1

    except Exception as err:
        return 'Error in adding attraction name:\n' + str(err)
    
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    try:
        # attraction rating 
        text = str(attr['rating']) + ' '
        if len(text) > 0:
            start_ind = end_ind 
            end_ind =  start_ind + len(text)   
            success = False
            for _ in range(10):
                try:
                    add_text_via_API(doc_id, text, start_ind, end_ind, 'Urbanist', 12, True, True, 0.51, 0.23, 0.25) 
                    success = True
                    break
                except:
                    start_ind -= 1

            if not success:
                output_msg(f"Failed to export the rating for attraction '{attr['name']}' to Google doc", 0)
                
                return -1

    except Exception as err:
        return 'Error in adding attraction rating:\n' + str(err)
    
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    try:
        # attraction rating stars 
        text = ''
        full_stars = int(float(attr['rating']))
        frac = int(str(attr['rating']).split('.')[-1])
        for _ in range(full_stars):
            text +=  '★'
        if frac >= 5 and full_stars < 5:
            text +=  '★'
        elif frac < 5 and full_stars < 5:
            text +=  '☆'
        nstars = 5 - len(text)
        for _ in range(nstars):
            text +=  '☆'
        if len(text) > 0:
            start_ind = end_ind 
            end_ind =  start_ind + len(text)  
            success = False
            for _ in range(10):
                try:
                    add_text_via_API(doc_id, text, start_ind, end_ind, 'Urbanist', 11, False, True, 0, 0.7, 1)
                    success = True
                    break
                except:
                    start_ind -= 1

            if not success: 
                output_msg(f"Failed to export the rating stars for attraction '{attr['name']}' to Google doc", 0)
                
                return -1

    except Exception as err:
        return 'Error in adding attraction rating stars:\n' + str(err)
    
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    try:
        # attraction number of reviews 
        if int(attr['reviews']) > 1:
            text = ' (' + format(attr['reviews'], ',d') + ' reviews)'
        else:
            text = ' (' + format(attr['reviews'], ',d') + ' review)'
        if len(text) > 0:
            start_ind = end_ind 
            end_ind =  start_ind + len(text)
            text = text + '\n\n'
            success = False
            for _ in range(10):
                try:
                    add_text_via_API(doc_id, text, start_ind, end_ind, 'Urbanist', 11, False, True, 0.32, 0.32, 0.32)
                    success = True
                    break
                except:
                    start_ind -= 1

            if not success:
                output_msg(f"Failed to export the reviews for attraction '{attr['name']}' to Google doc", 0)
                
                return  -1

    except Exception as err:
        return 'Error in adding attraction number of reviews:\n' + str(err)
    
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    # attraction image 
    img = False
    try:
        if isinstance(attr['image'], str) and len(attr['image']) > 0:  
            # new lines before the image
            text = '\n'
            start_ind = end_ind 
            end_ind =  start_ind + len(text) 
            success = False
            for _ in range(10):
                try:
                    add_text_via_API(doc_id, text, start_ind, end_ind, 'Martel', 20, False, False, 0, 0, 0)
                    success = True
                    break
                except:
                    start_ind -= 1

            img = attr['image']
            start_ind = end_ind 
            end_ind = start_ind + 1

            success = False
            for _ in range(10):
                try:
                    # adding image
                    requests = [{'insertInlineImage': {'location': {'index': start_ind},'uri':img,'objectSize': {'height':{'magnitude':400,'unit':'PT'},'width': {'magnitude': 400,'unit': 'PT'}}}}, {"updateParagraphStyle": {"range": {"startIndex": start_ind,"endIndex": end_ind},"paragraphStyle": {"alignment": "CENTER"},"fields": "alignment"}}]
                    dummy_driver = get_url('https://www.google.com', dummy_driver)
                    service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
                    img = True
                    success = True
                    break
                except:
                    start_ind -= 1

            if not success:
                output_msg(f"Failed to export the image for attraction '{attr['name']}' to Google doc", 0)
                
                return -1

    except Exception as err:
        return 'Error in adding attraction image:\n' + str(err)

    dummy_driver = get_url('https://www.google.com', dummy_driver)

    if img:
        try:
            # attraction image credit
            if len(attr['image_url']) > 0:
                url = attr['image_url']
                text = '\nCredit:'
                start_ind = end_ind 
                end_ind =  start_ind + len(text) 
                success = False
                for _ in range(10):
                    try:
                        add_hyperlink_via_API(doc_id, text, url, start_ind, end_ind, start_ind, end_ind, 'Urbanist Medium', 10, False, False, 0.51, 0.23, 0.25) 
                        success = True
                        break
                    except:
                        start_ind -= 1

                if not success:
                    output_msg(f"Failed to export the image credit for attraction '{attr['name']}' to Google doc", 0)
                    
                    return -1

                text = ' '
                start_ind = end_ind 
                end_ind =  start_ind + len(text) 
                success = False
                for _ in range(10):
                    try:
                        add_text_via_API(doc_id, text, start_ind, end_ind, 'Urbanist', 10, False, False, 0.51, 0.23, 0.25)
                        success = True
                        break
                    except:
                        start_ind -= 1

            if not success:
                output_msg(f"Failed to export the image credit for attraction '{attr['name']}' to Google doc", 0)
                
                return -1

            # adding credit hyperlinks
            if isinstance(attr['credit_name'], str) and len(attr['credit_name']) > 0:
                if len(attr['image_url']) > 0:
                    text = attr['credit_name']
                else:
                    text = '\nCredit: ' + attr['credit_name']
                url = attr['credit_url']
                if len(text) > 0:
                    start_ind = end_ind 
                    end_ind =  start_ind + len(text)
                    success = False
                    for _ in range(10):
                        try:
                            add_hyperlink_via_API(doc_id, text, url, start_ind, end_ind, start_ind, end_ind, 'Urbanist Medium', 10, False, False, 0.51, 0.23, 0.25)
                            success = True
                            break
                        except:
                            start_ind -= 1

            if not success:
                output_msg(f"Failed to export the image credit for attraction '{attr['name']}' to Google doc", 0)
                
                return  -1

        except Exception as err:
            return 'Error in adding attraction image credit:\n' + str(err)

        dummy_driver = get_url('https://www.google.com', dummy_driver)

        try:
        # adding license hyperlinks
            if isinstance(attr['license_name'], str) and len(attr['license_name']) > 0:
                url = attr['license_url']
                if len(attr['credit_name']) > 0:
                    text = ', ' + attr['license_name']
                    start_ind = end_ind 
                    end_ind =  start_ind + len(text) 
                    style_start_ind = start_ind+2
                    style_end_ind = end_ind
                else:
                    text = attr['license_name']
                    start_ind = end_ind 
                    end_ind =  start_ind + len(text) 
                    style_start_ind = start_ind
                    style_end_ind = end_ind
                success = False
                for _ in range(10):
                    try:
                        add_hyperlink_via_API(doc_id, text, url, start_ind, end_ind, style_start_ind, style_end_ind, 'Urbanist Medium', 10, False, False, 0.51, 0.23, 0.25)
                        success = True
                        break
                    except:
                        start_ind -= 1

            if not success:
                output_msg(f"Failed to export the image credit for attraction '{attr['name']}' to Google doc", 0)
                
                return -1

        except Exception as err:
            return 'Error in adding attraction image license:\n' + str(err)            

        dummy_driver = get_url('https://www.google.com', dummy_driver)

        try:        
            # adding image hyperlinks
            if isinstance(attr['image_name'], str) and len(attr['image_name']) > 0:
                url = attr['image_url']
                if len(attr['license_url']) > 0:
                    text = ', ' + attr['image_name']
                    start_ind = end_ind 
                    end_ind =  start_ind + len(text)
                    style_start_ind = start_ind+2
                    style_end_ind = end_ind
                else:
                    text = attr['image_name']
                    start_ind = end_ind 
                    end_ind =  start_ind + len(text)
                    style_start_ind = start_ind
                    style_end_ind = end_ind
                success = False
                for _ in range(10):
                    try:
                        add_hyperlink_via_API(doc_id, text, url, start_ind, end_ind, style_start_ind, style_end_ind, 'Urbanist Medium', 10, False, False, 0.51, 0.23, 0.25)
                        success = True
                        break
                    except:
                        start_ind -= 1

            if not success:
                output_msg(f"Failed to export the image link for attraction '{attr['name']}' to Google doc", 0)
                
                return -1

        except Exception as err:
            return 'Error in adding attraction image link:\n' + str(err) 

    dummy_driver = get_url('https://www.google.com', dummy_driver)

    address = False
    # Location text
    try:
        text = attr['address']
        if isinstance(text, str) and len(text) > 0:
            address = True
            text = '\n\nLocation: '
            start_ind = end_ind 
            end_ind =  start_ind + len(text) 
            success = False
            for _ in range(10):
                try:
                    add_text_via_API(doc_id, text, start_ind, end_ind, 'Urbanist', 13, True, False, 0.51, 0.23, 0.25)
                    success = True
                    break
                except:
                    start_ind -= 1

            if not success:
                output_msg(f"Failed to export the address for attraction '{attr['name']}' to Google doc", 0)
                
                return -1
            
            # address text
            text = attr['address']
            start_ind = end_ind 
            end_ind =  start_ind + len(text) 
            success = False
            for _ in range(10):
                try:
                    add_text_via_API(doc_id, text, start_ind, end_ind, 'Martel', 13, False, False, 0, 0, 0) 
                    success = True
                    break
                except:
                    start_ind -= 1

            if not success:
                output_msg(f"Failed to export the address for attraction '{attr['name']}' to Google doc", 0)
                
                return -1

    except Exception as err:
        return 'Error in adding attraction address:\n' + str(err) 

    dummy_driver = get_url('https://www.google.com', dummy_driver)

    phone = False
    try:
        text = attr['phone']
        if isinstance(text, str) and len(text) > 0:
            phone = True
            # reformat if no address is available
            if address:
                text = '\nTel:  '
            else:
                text = '\n\nTel:  '
            start_ind = end_ind 
            end_ind =  start_ind + len(text) 
            success = False
            for _ in range(10):
                try:
                    add_text_via_API(doc_id, text, start_ind, end_ind, 'Urbanist', 13, True, False, 0.51, 0.23, 0.25) 
                    success = True
                    break
                except:
                    start_ind -= 1

            if not success:
                output_msg(f"Failed to export the phone number for attraction '{attr['name']}' to Google doc", 0)
                
                return -1
            
            # phone text
            text = attr['phone']
            start_ind = end_ind 
            end_ind =  start_ind + len(text)
            success = False
            for _ in range(10):
                try:
                    add_text_via_API(doc_id, text, start_ind, end_ind, 'Martel', 13, False, False, 0, 0, 0)
                    success = True
                    break
                except:
                    start_ind -= 1

            if not success:
                output_msg(f"Failed to export the phone number for attraction '{attr['name']}' to Google doc", 0)
                
                return -1

    except Exception as err:
        return 'Error in adding attraction phone number:\n' + str(err) 

    dummy_driver = get_url('https://www.google.com', dummy_driver)

    try:
        text = attr['website']
        if isinstance(text, str) and len(text) > 0:
            # reformat if no address or phone are available
            if not address and not phone:
                text = '\n\nWeb Address:   '
            else:
                text = '\nWeb Address:   '
            start_ind = end_ind 
            end_ind =  start_ind + len(text) 
            success = False
            for _ in range(10):
                try:
                    add_text_via_API(doc_id, text, start_ind, end_ind, 'Urbanist', 13, True, False, 0.51, 0.23, 0.25)
                    success = True
                    break
                except:
                    start_ind -= 1

            if not success:
                output_msg(f"Failed to export the website for attraction '{attr['name']}' to Google doc", 0)
                
                return -1

            # website text
            text = attr['website']
            start_ind = end_ind 
            end_ind =  start_ind + len(text) 
            success = False
            for _ in range(10):
                try:
                    add_text_via_API(doc_id, text, start_ind, end_ind, 'Martel', 13, False, False, 0, 0, 0)
                    success = True
                    break
                except:
                    start_ind -= 1

            if not success:
                output_msg(f"Failed to export the website for attraction '{attr['name']}' to Google doc", 0)
                
                return -1

    except Exception as err:
        return 'Error in adding attraction website:\n' + str(err) 

    dummy_driver = get_url('https://www.google.com', dummy_driver)
    try:
        # new line for the next attraction
        text = '\n\n'
        start_ind = end_ind 
        end_ind =  start_ind + len(text) 
        for _ in range(10):
            try:
                add_text_via_API(doc_id, text, start_ind, end_ind, 'Martel', 15, False, False, 0, 0, 0)  
                break
            except:
                start_ind -= 1
    except Exception as err:
        return 'Error in adding attraction spacing:\n' + str(err) 

    output_msg(f'Attraction {attr_id} is exported to Google doc successfully', 1)


    return end_ind

def move_doc(doc_id, country, dest, folder_id):

    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    credentials = get_Google_API_creds()
    drive_service = build('drive', 'v3', credentials=credentials, cache_discovery=False)

    try:
        ## getting existing folders names in the google drive under the input folder
        date_folders = {}
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        response = drive_service.files().list(q = f'mimeType = "application/vnd.google-apps.folder" and trashed = false and "{folder_id}" in parents').execute()
        for folder in response.get('files', []):
            dummy_driver = get_url('https://www.google.com', dummy_driver)
            date_folders[folder.get('name')] = folder.get('id')
        # creating date folder inside the user input folder if not created yet
        folder_name = datetime.now().strftime("%d-%b-%y")
        if folder_name not in date_folders.keys():
            try:
                file_metadata = {'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [folder_id]}
                dummy_driver = get_url('https://www.google.com', dummy_driver)
                folder = drive_service.files().create(body=file_metadata,fields='id').execute()
                date_folder_id = folder.get('id')
            except Exception as err:
                output_msg('The following error occurred while moving the document to the user specified folder on Google drive', 1)
                output_msg(err, 0)
        else:
            # folder already created
            date_folder_id = date_folders[folder_name]

        # creating a state folder for US destinations
        if dest.find(',') > 0:
            country = dest.split(',')[-1].strip()

        # creating country folder inside the date one
        dest_folders = {}
        query = f'mimeType = "application/vnd.google-apps.folder" and trashed = false and "{date_folder_id}" in parents'
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        response = drive_service.files().list(q = query).execute()
        for folder in response.get('files', []):
            dummy_driver = get_url('https://www.google.com', dummy_driver)
            dest_folders[folder.get('name')] = folder.get('id')
        if country not in dest_folders.keys():
            file_metadata = {'name': country, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [date_folder_id]}
            dummy_driver = get_url('https://www.google.com', dummy_driver)
            folder = drive_service.files().create(body=file_metadata,fields='id').execute()
            dest_folder_id = folder.get('id')
        else:
            # folder already created
            dest_folder_id = dest_folders[country]

        # checking destinations files in the country folder
        attr_files = {}
        query = f'mimeType != "application/vnd.google-apps.folder" and trashed = false and "{dest_folder_id}" in parents'
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        response = drive_service.files().list(q = query).execute()
        for attr_file in response.get('files', []):
            dummy_driver = get_url('https://www.google.com', dummy_driver)
            attr_files[attr_file.get('name')] = attr_file.get('id')

        # if there is an old destination file with the same name then remove it to be replaced with the updated one
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        doc = drive_service.files().get(fileId=doc_id).execute()
        # getting the name from the doc dictionary
        title = doc.get('name', 'None')
        if title in attr_files.keys():
            attr_file_id = attr_files[title]
            dummy_driver = get_url('https://www.google.com', dummy_driver)
            drive_service.files().delete(fileId=attr_file_id).execute()

        # moving the new destination file to the country folder
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        doc = drive_service.files().get(fileId=doc_id,
                                    fields='parents').execute()
        previous_parents = ",".join(doc.get('parents'))
        doc = drive_service.files().update(fileId=doc_id, removeParents=previous_parents, addParents=dest_folder_id, fields='id, parents').execute()

    except Exception as err:
        # failure in exporting the data to google doc
        output_msg('The following error occurred while moving the document to the correct folder on Google drive', 1)
        output_msg(err, 0)
        
    return


def get_attraction_image(driver, attr, dest, nwiki, nflickr, ngoogle_official, ngoogle_user):

    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    exclude = ['from', 'to', '&', 'by', 'and', 'as', 'in', 'on', 'at', 'of', 'for', 'the', 'a', 'an', 'is', 'are', 'there', 'those', 'these', 'any', 'not', 'all', 'who', 'with', 'you', 'your', 'i', 'how', 'me', 'have', 'has', 'what', 'my', 'had', 'asap', 'him', 'her', 'his', 'about', 'new', 'would', 'will', 'should', 'could', 'be', 'can', 'could', 'it', 'its', "it's" 'just', 'by','do','does','did','done','that','he','she','or','one','our','that','their','was','we', 'if', 'so', 'they', 'when', 'up', 'after', 'but']

    # case1: Wikimedia search for the image
    output_msg('Checking for attraction image in Wikimedia ...', 0)
    try:
        # Open a new window
        driver.execute_script("window.open('');")
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(1)
        driver = get_url('https://commons.wikimedia.org/', driver)
        time.sleep(1)
        # preparing search text
        text = attr['name'] + ', ' + dest
        #driver.execute_script(f"document.getElementById('searchInput').value='{text}'")
        search = wait(driver, 1).until(EC.presence_of_element_located((By.ID, "searchInput")))
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        search.send_keys(text)
        time.sleep(1)
        button = wait(driver, 1).until(EC.presence_of_element_located((By.ID, "searchButton")))
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        driver.execute_script("arguments[0].click();", button) 
        time.sleep(2)

        # selecting large image size
        buttons = wait(driver, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "button")))
        for button in buttons:
            if 'Image size' in button.text:
                dummy_driver = get_url('https://www.google.com', dummy_driver)
                driver.execute_script("arguments[0].click();", button) 
                time.sleep(1)
                div = wait(driver, 1).until(EC.presence_of_element_located((By.XPATH, "//div[@class='sd-select-menu']")))
                dummy_driver = get_url('https://www.google.com', dummy_driver)
                wait(div, 1).until(EC.presence_of_element_located((By.XPATH, "//li[@id='fileres__listbox-item-3']"))).click()
                time.sleep(3)
                break

        try:
            # checking if there are search results
            res = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.sdms-search-results")))
            url = wait(res, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.sdms-image-result")))
        except:

             # selecting medium image size
            buttons = wait(driver, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "button")))
            for button in buttons:
                if 'Large' in button.text:
                    driver.execute_script("arguments[0].click();", button) 
                    time.sleep(1)
                    div = wait(driver, 1).until(EC.presence_of_element_located((By.XPATH, "//div[@class='sd-select-menu']")))
                    dummy_driver = get_url('https://www.google.com', dummy_driver)
                    wait(div, 1).until(EC.presence_of_element_located((By.XPATH, "//li[@id='fileres__listbox-item-2']"))).click()
                    time.sleep(3)
                    break
        # checking if there are search results
        res = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.sdms-search-results")))
        imgs = wait(res, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.sdms-image-result")))
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        res_url = driver.current_url
        n = len(imgs)
        for i in range(n):
            res = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.sdms-search-results")))
            url = wait(res, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.sdms-image-result")))[i].get_attribute('href')
            driver = get_url(url, driver)
            time.sleep(1)
            # getting the image url
            img_div = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.fullImageLink")))
            img_url = wait(img_div, 1).until(EC.presence_of_element_located((By.TAG_NAME, "img")))
            attr['image'] = img_url.get_attribute('src')
            attr['image_name'] = 'Wikimedia commons'
            attr['image_url'] = driver.current_url
            # getting the credit info
            credit_div = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.hproduct.commons-file-information-table")))
            trs = wait(credit_div, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "tr")))
            credit_name, credit_url = '', ''
            for tr in trs:
                if 'Author' in tr.text:
                    td1 = wait(tr, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "td")))[0]
                    if td1.text.strip() != 'Author': continue
                    td = wait(tr, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "td")))[1]
                    credit_name = td.text.strip()
                    words = credit_name.split(' ')
                    # excluding unwanted words
                    for word in words:
                        if word.lower() in exclude:
                            credit_name = credit_name.replace(' ' + word, '')
                    # limiting the name by three words at max
                    try:
                        trim_name = ' '.join(credit_name.split(' ')[:3])
                        credit_name = trim_name.strip()
                    except:
                        pass
                    try:
                        credit_url = wait(td, 1).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('href')
                        credit_name = wait(td, 1).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).text.strip()
                        words = credit_name.split(' ')
                        # excluding unwanted words
                        for word in words:
                            if word.lower() in exclude:
                                credit_name = credit_name.replace(' ' + word, '')
                        try:
                            trim_name = ' '.join(credit_name.split(' ')[:3])
                            credit_name = trim_name.strip()
                        except:
                            pass
                    except:
                        # author link is not available
                        credit_url = attr['image_url']
                    break
            attr['credit_url'] = credit_url
            attr['credit_name'] = credit_name
            attr['license_url'] = 'https://creativecommons.org/licenses/by-sa/4.0/deed.en'
            attr['license_name'] = 'License'
            if 'Internet Archive Book' in attr['credit_name']:
                driver = get_url(res_url, driver)
                time.sleep(2)
                continue
            try:
                # getting the license info
                license_div = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.layouttemplate.licensetpl.mw-content-ltr")))
                trs = wait(license_div, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "tr")))
                license_url = 'https://creativecommons.org/licenses/by-sa/4.0/deed.en'
                for tr in trs:
                    if 'This file is licensed under' in tr.text:
                        td = wait(tr, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "td")))[1]
                        licence = wait(td, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))[-1]
                        #license_name = licence.text.strip()
                        license_url = licence.get_attribute('href')
                        attr['license_url'] = license_url
                        #attr['license_name'] = license_name
                        break
            except:
                pass
            # closing the tab
            driver.close()
            # back to the first tab of the browser window
            dummy_driver = get_url('https://www.google.com', dummy_driver)
            driver.switch_to.window(driver.window_handles[0])
            nwiki += 1
            output_msg('Successfully scraped attraction image from Wikimedia', 0)
            return nwiki, nflickr, ngoogle_official, ngoogle_user
    except:
        # no search results
        output_msg('No valid image is found in Wikimedia ...', 0)
        pass

    # case2: Flickr search for the image
    output_msg('Checking for attraction image in Flicker ...', 0)
    try:
        driver = get_url('https://www.flickr.com/', driver)
        time.sleep(1)
        # preparing search text
        text = attr['name'] + ', ' + dest
        #driver.execute_script(f"document.getElementById('search-field').value='{text}'")
        search = wait(driver, 1).until(EC.presence_of_element_located((By.ID, "search-field")))
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        search.send_keys(text)
        time.sleep(1)
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        search.send_keys(Keys.ENTER)
        #driver.find_element_by_id("search-field").send_keys(Keys.ENTER)
        time.sleep(2)
        # filtering results by all creative commons license 
        button = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.dropdown-link.filter-license")))
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        driver.execute_script("arguments[0].click();", button)
        time.sleep(3)
        try:
            menu = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.droparound.menu")))
            lis = wait(menu, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "li")))
            for li in lis:
                if 'creative commons' in li.text:
                    dummy_driver = get_url('https://www.google.com', dummy_driver)
                    driver.execute_script("arguments[0].click();", li)
                    time.sleep(3)
                    break
        except:
            pass
        # checking if there are search results
        res = wait(driver, 1).until(EC.presence_of_element_located((By.ID, "search-unified-content")))
        imgs = wait(res, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.view.photo-list-photo-view.awake")))
        res_url = driver.current_url
        n = len(imgs)
        for i in range(n):
            res = wait(driver, 1).until(EC.presence_of_element_located((By.ID, "search-unified-content")))
            img = wait(res, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.view.photo-list-photo-view.awake")))[i]
            button = wait(img, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.overlay")))
            dummy_driver = get_url('https://www.google.com', dummy_driver)
            driver.execute_script("arguments[0].click();", button)
            time.sleep(3)
            link = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "img.main-photo")))
            attr['image'] = link.get_attribute('src')
            attr['image_name'] = 'Flickr'
            attr['image_url'] = driver.current_url
            credit = wait(driver, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.owner-name.truncate")))[-1]
            attr['credit_url'] = credit.get_attribute('href')
            attr['credit_name'] = credit.text.strip()
            # skipping images from internet archive book
            if 'Internet Archive Book' in attr['credit_name']:
                driver = get_url(res_url, driver)
                time.sleep(2)
                continue
            license = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.photo-license-url")))
            attr['license_url'] = license.get_attribute('href')
            attr['license_name'] = 'License'
            # closing the tab
            driver.close()
            # back to the first tab of the browser window
            driver.switch_to.window(driver.window_handles[0])
            nflickr += 1
            output_msg('Successfully scraped attraction image from Flicker', 0)
            return nwiki, nflickr, ngoogle_official, ngoogle_user
    except:
        # no search results
        output_msg('No valid image is found in Flicker', 0)
        pass

    # closing all opened tabs except the first one
    try:
        for handle in driver.window_handles[1:]:
            dummy_driver = get_url('https://www.google.com', dummy_driver)
            driver.switch_to.window(handle)
            driver.close()
        # back to the first tab of the browser window
        dummy_driver = get_url('https://www.google.com', dummy_driver)
        driver.switch_to.window(driver.window_handles[0])
        # getting all the attraction images
    except:
        pass
    try:
        imgs = wait(driver, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.QtzoWd")))
    except:
        # single image handling
        imgs = wait(driver, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "img.R1Ybne.YH2pd")))

    # case3: image with text that matches the attraction name
    output_msg('Checking for attraction image in Google official ...', 0)
    try:
        for img in imgs:
            if attr['name'] == img.text.strip():
                img_url = wait(img, 1).until(EC.presence_of_element_located((By.TAG_NAME, "img")))
                url = img_url.get_attribute('src')
                if isinstance(url, str) and len(url) > 0:
                    attr['image'] = url
                    attr['image_name'] = img_url.text.strip()
                    attr['image_url'] = ''
                else:
                    # no valid image is found
                    break
                credit = wait(img, 1).until(EC.presence_of_element_located((By.TAG_NAME, "a")))
                attr['credit_url'] = credit.get_attribute('href')
                attr['credit_name'] = attr['name'] + ', Google Maps'
                attr['license_url'] = ''
                attr['license_name'] = ''
                ngoogle_official += 1
                output_msg('Successfully scraped attraction image from Google official', 0)
                return nwiki, nflickr, ngoogle_official, ngoogle_user
    except:
        pass
    # case4: the attraction has an image with text that doesn't match
    output_msg('No valid image is found in Google official', 0)
    output_msg('Checking for attraction image in Google user ...', 0)
    # getting all the attraction images
    try:
        imgs = wait(driver, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.QtzoWd")))
    except:
        # single image handling
        imgs = wait(driver, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "img.R1Ybne.YH2pd")))

    try:
        for img in imgs:
            if len(img.text) > 0:
                img_url = wait(img, 1).until(EC.presence_of_element_located((By.TAG_NAME, "img")))
                attr['image'] = img_url.get_attribute('src')
                attr['image_name'] = ''
                attr['image_url'] = ''
                credit = wait(img, 1).until(EC.presence_of_element_located((By.TAG_NAME, "a")))
                attr['credit_url'] = credit.get_attribute('href')
                attr['credit_name'] = credit.text.strip() + ', Google Maps'
                attr['license_url'] = ''
                attr['license_name'] = ''
                ngoogle_user += 1
                output_msg('Successfully scraped attraction image from Google user', 0)
                return nwiki, nflickr, ngoogle_official, ngoogle_user

        output_msg('No valid image is found in Google user', 0)
        # case5: no image can bo found for the attraction
        attr['image'] = ''
        attr['image_name'] = ''
        attr['image_url'] = ''
        attr['credit_url'] = ''
        attr['credit_name'] = ''
        attr['license_url'] = ''
        attr['license_name'] = ''

        return nwiki, nflickr, ngoogle_official, ngoogle_user
    except:
        attr['image'] = ''
        attr['image_name'] = ''
        attr['image_url'] = ''
        attr['credit_url'] = ''
        attr['credit_name'] = ''
        attr['license_url'] = ''
        attr['license_name'] = ''
        return nwiki, nflickr, ngoogle_official, ngoogle_user

def get_attraction_info(driver, name, rating, nrev):

    global dummy_driver
    dummy_driver = get_url('https://www.google.com', dummy_driver)

    attr = {}
    # checking once if the displayed language is English
    try:
        buttons = wait(driver, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))
        for button in buttons:
            if button.text == 'Change to English':
                dummy_driver = get_url('https://www.google.com', dummy_driver)
                driver.execute_script("arguments[0].click();", button)
                time.sleep(1)
                break
    except:
        pass

    # Website info of the attraction
    site = ''
    try:
        buttons = wait(driver, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.ab_button")))
        for button in buttons:
            if button.text == 'Website':
                site = button.get_attribute('href')
                # removing what is after the "?" in the link
                if site.find('?') > 0:
                    site = site[:site.find('?')]
                # removing "www." from the link   
                site = site.replace('www.', '').replace('WWW.', '')
                break
    except:
        pass

    attr['country'] = ''
    # Address info of the attraction
    add = '' 
    try:
        divs = wait(driver, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.zloOqf.PZPZlf")))
        
        for div in divs:
            if 'Address' in div.text:
                add = wait(div, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.LrzXr"))).text

                # remove ascents from add
                add = unidecode.unidecode(add)
                # removing non ASCII characters from the address
                add =  re.sub(r'[^\x00-\x7f]',r'', add) 
                add = add.replace('.', ' ')
                # removing extra spaces
                add = " ".join(add.split())
                # removing country from address
                add = add.replace('/', ' ').replace('-', ' ')
                add_elems = add.split(',')
                updated_add = []
                for elem in add_elems:
                    trim = False
                    country_name = ''
                    for country in pycountry.countries:
                        if country.name in elem:
                            trim = True
                            attr['country'] = country.name
                            country_name = country.name
                            break                   

                    if trim:
                        text = elem.replace(country_name, '').strip()
                        if text != '':
                            updated_add.append(text)
                    else:
                        updated_add.append(elem)
                        
                # removing postal code from address
                if len(updated_add) > 0:
                    add_elems = updated_add[::-1]
                else:
                    add_elems = [add]
                updated_add = []
                trim = False
                #for elem in add_elems[:-1]:
                n = len(add_elems)
                for i, elem in enumerate(add_elems):
                    if not trim:
                        section = []
                        parts = elem.split(' ')
                        for j, part in enumerate(parts):
                            if part.isalpha() or part == ' ' or part == '' or "'" in part or (j == 0 and i == n-1):
                                section.append(part)
                            else:
                                trim = True
                        updated_add.append(' '.join(section))
                    else:
                        updated_add.append(elem)

                #updated_add.append(add_elems[-1])

                add = ''.join(updated_add[::-1])
                add = add.replace(',', '')

                # further format updates
                if add.find('Road') > 0:
                    add = add.replace('Road', 'Rd.') 
                elif add.find('Rd.') > 0:
                    add = add.replace('Rd.', 'Road')
                elif add.find('Rd') > 0:
                    add = add.replace('Rd', 'Road')

                if add.find('Street') > 0:
                    add = add.replace('Street', 'St') 
                elif add.find('St.') > 0:
                    add = add.replace('St.', 'Street')
                elif add.find('St') > 0:
                    add = add.replace('St', 'Street')

                break
    except:
        pass           
                
    # Tel info of the attraction
    tel = ''
    try:
        divs = wait(driver, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.zloOqf.PZPZlf")))
        for div in divs:
            if 'Phone' in div.text:
                tel = wait(div, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.LrzXr"))).text
                # changing the format to the desired one
                digits = tel.split(' ')
                for dig in digits:
                    if '+' in dig:
                        tel = ' '.join(digits[1:])
                        break

                chars = []
                for char in tel:
                    if char.isnumeric():
                        chars.append(char)
                    else:
                        if char == '-' or char == '.'or char == '('or char == ')':
                            chars.append(' ')
                        elif char == ' ':
                            chars.append('-')
                tel = ''.join(chars)
                break
    except:
        pass            
            
    # temp closed info
    closed = ''
    try:
        text = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.mWyH1d.UgLoB"))).text.strip()
        if text == 'Temporarily closed':
            closed = text
    except:
        pass

    # updating attraction name
    name = name.replace('/', '-')
    ## remove ascents
    #name = unidecode.unidecode(name)
    ## removing non ASCII characters from the address
    #name =  re.sub(r'[^\x00-\x7f]',r'', name) 
    # saving attraction info
    attr['name'] = name.strip()
    attr['rating'] = float(rating)
    attr['reviews'] = nrev
    attr['website'] = site
    attr['address'] = add
    attr['phone'] = tel
    attr['closed'] = closed

    return attr

def clear_screen():
    # for windows
    if os.name == 'nt':
        _ = os.system('cls')
    # for mac and linux
    else:
        _ = os.system('clear')

def check_password():

    global password_field, toplevel, root
    toplevel = tkinter.Toplevel(root)
    toplevel.title("Password")
    toplevel.geometry("300x100")
    toplevel.attributes("-topmost", True)
    toplevel.protocol("WM_DELETE_WINDOW", on_quit)
    ttk.Label(toplevel, text="Password:").place(relx=0.15, rely=0.3, anchor='center')

    password_field = ttk.Entry(toplevel, show="*")
    password_field.place(relx=0.62, rely=0.3, relwidth=0.68, relheight=0.35, anchor='center')

    button_login = tkinter.Button(toplevel, text="Login", command=check_password_func)
    button_login.place(relx=0.5, rely=0.6, relwidth=0.5, relheight=0.3, anchor='n')

def check_password_func():
    global password_field, toplevel, root
    if password_field.get() == "2305":
        toplevel.destroy()
        root.attributes("-topmost", True)
        root.attributes('-disabled', False)
    else:
        messagebox.showerror(title="Wrong password", message="Incorrect Password!")


def pre_start_on_thread():

    global root, running, google_sheet_enter_field, attractions_enter_field, google_drive_enter_field
    if running == False:
        running = True
        return

    # getting the arguments to be passed to the scraping function
    url = google_sheet_enter_field.get()
    limit = attractions_enter_field.get()
    folder = google_drive_enter_field.get()
    # saving the settings to the dll file 
    data = {"google_sheet": url, "attractions_limit": limit, 'password': True, "google_drive": folder}
    with open("settings.dll", "wb") as f:
        pickle.dump(data, f)

    button_start["state"] = "disabled"
    toplevel = tkinter.Toplevel(root)
    LoggerGUI(toplevel)

    # calling the scraper function
    thread = threading.Thread(target=scrape_destinations)
    thread.start()
    

def on_quit():
    global root, driver
    root.destroy()
    driver.quit()
    sys.exit()

class LoggerGUI(tkinter.Frame):
    def __init__(self, parent, *args, **kwargs):
        tkinter.Frame.__init__(self, parent, *args, **kwargs)
        self.root = parent
        self.build_gui()

    def build_gui(self):
        path = os.getcwd()
        #self.root.geometry("750x300")
        self.root.title(path)
        self.root.attributes("-topmost", True)
        self.root.option_add('*tearOff', 'FALSE')
        self.grid(column=0, row=0, sticky=tkinter.NSEW)
        self.grid_columnconfigure(0, weight=1, uniform='a')
        self.grid_columnconfigure(1, weight=1, uniform='a')
        self.grid_columnconfigure(2, weight=1, uniform='a')
        self.grid_columnconfigure(3, weight=1, uniform='a')
        st = ScrolledText.ScrolledText(self, state='disabled', width=100, height=20)

        st.configure(font='TkFixedFont')
        st.grid(column=0, row=1, sticky=tkinter.NSEW, columnspan=4)
        text_handler = TextHandler(st)

        logging.basicConfig(level=logging.INFO)

        logger = logging.getLogger()
        logger.addHandler(text_handler)

def load_settings():
    try:
        with open("settings.dll", "rb") as file:
            data = pickle.load(file)
        return data
    except FileNotFoundError:
        return {}

class TextHandler(logging.Handler):
    global start
    def __init__(self, text):
        logging.Handler.__init__(self)
        self.text = text

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text.configure(state='normal')
            self.text.insert(tkinter.END, msg + '\n')
            self.text.configure(state='disabled')
            self.text.yview(tkinter.END)

        try:
            self.text.after(0, append)
        except:
            elapsed = round((time.time() - start)/60, 2)
            hours = round(elapsed/60, 2)
            output_msg(f'The bot is manually terminated by the user. Elapsed Time: {elapsed} mins ({hours} hours)', 0)
            sys.exit()


def run_GUI():
    # GUI
    global root, running, button_start, button_stop, google_sheet_enter_field, attractions_enter_field, google_drive_enter_field, url, limit, version
    running = False
    # check settings dll file
    data = load_settings()
    # configuring the UI main window
    root = ThemedTk('breeze')
    root.resizable(True, True)
    root.protocol("WM_DELETE_WINDOW", on_quit)
    root.title(f"Attractions Scraper by Abdelrahman Hekal v{version}")
    root.geometry("825x300")
    root.attributes('-disabled', True)
    root.attributes("-topmost", False)
    # check user one time password
    #if data.get("password",-1) == -1:
    #    check_password()
    #else:
    root.attributes("-topmost", True)
    root.attributes('-disabled', False)
    
    # stying the window
    styles = ttk.Style()
    styles.configure("rights.TLabel", font="Verdana 10")
    styles.configure("TLabel", padding=1, font="Verdana 10")
    styles.theme_use('breeze')

    # google sheet user inputs
    ttk.Label(root, text="Google Sheet URL").place(relx=0.01, rely=0.1, anchor='w')
    google_sheet_enter_field = ttk.Entry(root)
    google_sheet_enter_field.place(relx=0.25, rely=0.1, relwidth=0.7, relheight=0.1, anchor='w')
    try:
        google_sheet_enter_field.insert(tkinter.END, data["google_sheet"])
    except:
        print(end="")    
        
    # google drive path to store the doc files
    ttk.Label(root, text="Google Drive Folder URL").place(relx=0.01, rely=0.3, anchor='w')
    google_drive_enter_field = ttk.Entry(root)
    google_drive_enter_field.place(relx=0.25, rely=0.3, relwidth=0.7, relheight=0.1, anchor='w')
    try:
        google_drive_enter_field.insert(tkinter.END, data["google_drive"])
    except:
        print(end="")

    ttk.Label(root, text="Attractions Limit").place(relx=0.01, rely=0.5, anchor='w')
    attractions_enter_field = ttk.Entry(root)
    attractions_enter_field.place(relx=0.3, rely=0.5, relwidth=0.1, relheight=0.1, anchor='center')
    try:
        attractions_enter_field.insert(tkinter.END, data["attractions_limit"])
    except:
        print(end="")    
        
    # status bar
    path = os.getcwd()
    ttk.Label(root, text=path, relief=tkinter.SUNKEN, anchor=tkinter.W).pack(side=tkinter.BOTTOM, fill=tkinter.X)

    button_start = ttk.Button(root, text="Start", command= pre_start_on_thread)
    button_start.place(relx=0.5, rely=0.8, relwidth=0.3, relheight=0.1, anchor='center')
    button_start.invoke()
    root.mainloop()

def scrape_destinations():
    
    
    global running, root, start, driver, version

    #url = 'https://docs.google.com/spreadsheets/d/1gj8S4eRiFuu7b2xHNwLCKL3wx7VI_ZfY4JMC07wE4tg/edit?usp=sharing'
    #limit = 100

    # getting user inputs from the UI
    url = google_sheet_enter_field.get()
    limit = attractions_enter_field.get()
    folder = google_drive_enter_field.get()
    
    # validating the google sheet url input
    if len(url) == 0 or 'docs.google.com/spreadsheets/d' not in url.lower():
        output_msg("Invalid Google sheet link, please try again!", 1)
        sys.exit()     
        
    # validating the google folder ID input
    if len(folder) == 0 or 'https://drive.google.com/drive/' not in folder.lower():
        output_msg("Invalid Google drive folder link, please try again!", 1)
        sys.exit()   

    folder = folder.split('/')[-1]

    try:
        limit = int(limit)
    except:
        output_msg("Invalid Attraction Limit Input, please try again!", 1)
        sys.exit()  
    if limit < 1:
        output_msg("Invalid Attraction Limit Input, please try again!", 1)
        sys.exit()  

    if os.path.exists(os.getcwd() + '\\session_log.log'):
             os.remove(os.getcwd() + '\\session_log.log') 

    start = time.time()
    output_msg(f'Starting the bot v{version} ...', 1)

    dests = process_sheet(url)
    driver = initialize_bot()
    clear_screen()
    for k, dest in enumerate(dests):
        # skip destinations with status mentioned
        if len(dest[1]) > 0:
            output_msg(f'Destination {dest[0].title()} has status "{dest[1]}", skipping...', 1)
            continue
        try:
            status = search_destinations(driver, dest, limit, url, k, folder)
            if len(status) != 0:
                # failure in the destination procesing
                output_msg(f'the following error occurred in exporting Google doc for destination {dest[0].title()}, skipping ...', 1)
                output_msg(status, 0)
                write_status(k+2, status, url)            

            #restarting the bot after each destination
            driver.quit()
            time.sleep(3)
            driver = initialize_bot()
        except Exception as err:
            output_msg(f'the below error occurred in destination {dest[0].title()} ...', 1)
            output_msg(str(err), 0)
            write_status(k+2, 'Error: ' + str(err).split('\n')[0], url)
            driver.quit()
            time.sleep(3)
            driver = initialize_bot()
            continue

    driver.quit()
    elapsed = round((time.time() - start)/60, 2)
    hours = round(elapsed/60, 2)
    output_msg(f'Process Completed Successfully! Elapsed Time: {elapsed} mins ({hours} hours)', 1)
       
# main program      
if __name__ == '__main__':

    global version, dummy_driver
    version = '2.8'
    dummy_driver = initialize_bot()
    # running the UI function
    run_GUI()








