import os
import sys
import openpyxl as pyxl
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from threading import Thread
import time

MAX_THREADS = 10 # max reasonable number of threads as each thread has a chomium driver
OFFSET = 2 # excel input data is offset by 2: 1 for 0 indexing and 1 for a row of titles
OFFSET_ROWS = 2 # excel input data has 2 extra rows: first is headers, last row is totaled info

SAVE_EVERY = 50 # save excel after every SAVE_EVERY number of elements scraped
MAX_NUM_TO_SCRAPE = 10 # max number of elements to scrape (set high to scrape everything)

# returns resource path for users environment given the relative path
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

# imports data from 'Equipment New List.xlsx' and returns as a dictionary
def getExcelValues():

    wb = pyxl.load_workbook('Equipment New List.xlsx')
    ws = wb.active
    n = ws.max_row - OFFSET_ROWS
    if n > MAX_NUM_TO_SCRAPE:
        n = MAX_NUM_TO_SCRAPE

    a1 = [None] * n
    a2 = [None] * n
    a3 = [None] * n
    a4 = [None] * n
    a5 = [None] * n
    a6 = [None] * n
    a7 = [None] * n
    a8 = [None] * n
    a9 = [None] * n
    a10 = [None] * n
    a11 = [None] * n
    a12 = [None] * n
    a13 = [None] * n
    
    i = 0
    while i != n:
        a1[i] = ws[f'A{i+2}'].value
        a2[i] = ws[f'B{i+2}'].value
        a3[i] = ws[f'C{i+2}'].value
        a4[i] = ws[f'D{i+2}'].value
        a5[i] = ws[f'E{i+2}'].value
        a6[i] = ws[f'F{i+2}'].value
        a7[i] = ws[f'G{i+2}'].value
        a8[i] = ws[f'H{i+2}'].value
        a9[i] = ws[f'I{i+2}'].value
        a10[i] = ws[f'J{i+2}'].value
        a11[i] = ws[f'K{i+2}'].value
        a12[i] = ws[f'L{i+2}'].value
        a13[i] = ws[f'M{i+2}'].value
        i += 1
    
    data = {
        'Emco' : a1,
        'Equipment' : a2,
        'Description' : a3,
        'VINNumber' : a4,
        'Manufacturer' : a5,
        'Model' : a6,
        'ModelYr' : a7,
        'OdoReading' : a8,
        'OdoDate' : a9,
        'HourReading' : a10,
        'HourData' : a11,
        'Location' : a12,
        'Complete' : a13
    }

    return data

# create search term for each item in data and return as array of search term strings
def get_search_terms(data):
    n = len(data['Emco'])
    search_terms = [None] * n
    for i in range(n):
        terms = []
        if data['Manufacturer'][i]:
            terms.append(data['Manufacturer'][i])
        if data['Model'][i]:
            terms.append(data['Model'][i])
        if data['ModelYr'][i]:
            terms.append(data['ModelYr'][i])
        search_term = ' '.join(str(term) for term in terms)
        if len(search_term) <= 8 and data['Description'][i]:
            search_terms[i] = f"{search_term} {data['Description'][i]}"
        else:
            search_terms[i] = search_term
    return search_terms

# scrape data from https://usedequipmentguide.com/ given a list of search terms
# saves results to 'Equipment New List.xlsx' as it searches
def scrapeAskingValues(search_terms):
    # constants and variables
    n = len(search_terms)
    avf = [None] * n # asking values found
    avl = [None] * n # asking value links
    
    # parse data: convert given dollar string to an int
    def parseDollarValue(money_str):
        str = money_str[1:] # remove dollar sign
        value = 0
        i = len(str) - 1
        multiplier = 1
        while i >= 0:
            if (len(str) - i) % 4 == 0:
                i -= 1
                continue
            else:
                value += multiplier * int(str[i])
                multiplier *= 10
                i -= 1
        return value

    # nested function for threaded scraping
    def scrape_task(index):
        driver = webdriver.Chrome(resource_path('./chromedriver_win32/chromedriver.exe')) 
        driver.get(f"https://usedequipmentguide.com/listings?query={search_terms[index]}")
        avl[index] = f"https://usedequipmentguide.com/listings?query={search_terms[index]}"
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "span.Span-hup779-0.sc-16afded-0.kzgLyd")
            )) # will only wait for first result
            time.sleep(1.5) # testing shows that an extra 1.5 seconds allows all results to finish
            elements = driver.find_elements(by=By.CSS_SELECTOR, value="span.Span-hup779-0.sc-16afded-0.kzgLyd")
            
            # pick most relavent result
            i = 0
            while i < len(elements):
                if elements[i].text != "AUCTION" and elements[i].text != "Price Unavailable" and parseDollarValue(elements[i].text) > 999:
                    avf[index] = parseDollarValue(elements[i].text)
                    break
                i += 1
        
        finally:
            driver.quit()
    
    i = 0
    while i < n:
        # run next set of threads
        threads = [None] * MAX_THREADS
        ti = 0
        while ti < MAX_THREADS and i < n:
            threads[ti] = Thread(target=scrape_task, args=(i,))
            threads[ti].start()
            i += 1
            ti += 1
        for j in range(ti):
            threads[j].join()

        # occasionally save what is found
        if i % SAVE_EVERY == 0:
            row_start = i-SAVE_EVERY
            temp_dict = {
                'Asking Value Found' : avf[i-SAVE_EVERY:i],
                'Asking Value Link' : avl[i-SAVE_EVERY:i]
            }
            tempSetExcel(temp_dict, row_start)

    # save final results
    row_start = i-SAVE_EVERY
    # corner case: SAVE_EVERY is large and row start becomes negative
    if row_start < 0:
        row_start = 0
    temp_dict = {
        'Asking Value Found' : avf[i-SAVE_EVERY:i],
        'Asking Value Link' : avl[i-SAVE_EVERY:i]
    }
    tempSetExcel(temp_dict, row_start)

    # set found prices and return
    dict = {
        'Asking Value Found' : avf,
        'Asking Value Link' : avl
    }
    return dict

# scrape data from ebay's trucks and cars site given a list of search terms
# saves results to 'Equipment New List.xlsx' as it searches
def scrapeAuctionValues(search_terms):
    # constants and variables
    n = len(search_terms)
    avf = [None] * n # auction values found
    avl = [None] * n # auction value links
    
    # parse ebay data: convert given dollar string to an int
    def parseDollarValue(money_str):
        str = money_str[1:-3] # remove dollar sign and pennies
        value = 0.01 * int(money_str[-2:]) # value in the pennies
        i = len(str) - 1
        multiplier = 1
        while i >= 0:
            if (len(str) - i) % 4 == 0:
                i -= 1
                continue
            else:
                value += multiplier * int(str[i])
                multiplier *= 10
                i -= 1
        return value

    # nested function for threaded scraping
    def scrape_task(index):
        driver = webdriver.Chrome(resource_path('./chromedriver_win32/chromedriver.exe')) 
        driver.get(f"https://www.ebay.com/sch/i.html?_from=R40&_nkw={search_terms[index]}&_sacat=6001&rt=nc&LH_Sold=1&LH_Complete=1")
        avl[index] = f"https://www.ebay.com/sch/i.html?_from=R40&_nkw={search_terms[index]}&_sacat=6001&rt=nc&LH_Sold=1&LH_Complete=1"
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "li.s-item.s-item__pl-on-bottom")
            )) # will only wait for first result
            time.sleep(1.5) # testing shows that an extra 1.5 seconds allows all results to finish
            
            elements = driver.find_elements(by=By.CSS_SELECTOR, value="span.POSITIVE")
            
            # pick only relevent trucks and cars data
            if parseDollarValue(elements[1].text) > 999:
                avf[index] = parseDollarValue(elements[1].text) # 2nd element is the 1st most relevent price
        
        finally:
            driver.quit()
    
    i = 0
    while i < n:
        # run next set of threads
        threads = [None] * MAX_THREADS
        ti = 0
        while ti < MAX_THREADS and i < n:
            threads[ti] = Thread(target=scrape_task, args=(i,))
            threads[ti].start()
            i += 1
            ti += 1
        for j in range(ti):
            threads[j].join()

        # occasionally save what is found
        if i % SAVE_EVERY == 0:
            row_start = i-SAVE_EVERY
            temp_dict = {
                'Auction Value Found' : avf[i-SAVE_EVERY:i],
                'Auction Value Link' : avl[i-SAVE_EVERY:i]
            }
            tempSetExcel(temp_dict, row_start)

    # save final results
    row_start = i-SAVE_EVERY
    # corner case: SAVE_EVERY is large and row start becomes negative
    if row_start < 0:
        row_start = 0
    temp_dict = {
        'Auction Value Found' : avf[i-SAVE_EVERY:i],
        'Auction Value Link' : avl[i-SAVE_EVERY:i]
    }
    tempSetExcel(temp_dict, row_start)

    # set found prices and return
    dict = {
        'Auction Value Found' : avf,
        'Auction Value Link' : avl
    }
    return dict

# arr_values is 2d array. Each item is an array representing data for a column 
# arr_col_strs is array of strings. arr_col_strs at index i is the col for arr_values at i
# tempSetExcel will set the values in the respective columns in 'Equipment New List.xlsx',
# starting at the row row_start
def tempSetExcel(dict, row_start):
    wb = pyxl.load_workbook('Equipment New List.xlsx')
    ws = wb.active

    # set auction values if they are given
    if 'Auction Value Found' in dict:
        row = row_start
        for val in dict['Auction Value Found']:
            ws[f'Q{row + OFFSET}'] = val
            row += 1
    
    # set auction value links if they are given
    if 'Auction Value Link' in dict:
        row = row_start
        for val in dict['Auction Value Link']:
            ws[f'R{row + OFFSET}'] = val
            row += 1

    # set market values found if they are given
    if 'Market Value Found' in dict:
        row = row_start
        for val in dict['Market Value Found']:
            ws[f'S{row + OFFSET}'] = val
            row += 1
    
    # set market value links found if they are given
    if 'Market Value Link' in dict:
        row = row_start
        for val in dict['Market Value Link']:
            ws[f'T{row + OFFSET}'] = val
            row += 1
    
    # set asking values if they are given
    if 'Asking Value Found' in dict:
        row = row_start
        for val in dict['Asking Value Found']:
            ws[f'U{row + OFFSET}'] = val
            row += 1

    # set asking value links if they are given
    if 'Asking Value Link' in dict:
        row = row_start
        for val in dict['Asking Value Link']:
            ws[f'V{row + OFFSET}'] = val
            row += 1

    # set search terms if they are given
    if 'Search Terms' in dict:
        row = row_start
        for val in dict['Search Terms']:
            ws[f'W{row + OFFSET}'] = val
            row += 1

    wb.save('Equipment New List.xlsx')

# sets given final prices in 'Equipment New List.xlsx'
def setExcel(dict):
    wb = pyxl.load_workbook('Equipment New List.xlsx')
    ws = wb.active
    n = len(dict['Search Terms'])

    row = 0 + OFFSET
    while row < n:
        ws[f'W{row}'] = dict['Search Terms'][row - OFFSET]
        if 'Auction Value Found' in dict:
            ws[f'Q{row}'] = dict['Auction Value Found'][row - OFFSET]
        if 'Auction Value Link' in dict:
            ws[f'R{row}'] = dict['Auction Value Link'][row - OFFSET]
        if 'Market Value Found' in dict:
            ws[f'S{row}'] = dict['Market Value Found'][row - OFFSET]
        if 'Market Value Link' in dict:
            ws[f'T{row}'] = dict['Market Value Link'][row - OFFSET]
        if 'Asking Value Found' in dict:
            ws[f'U{row}'] = dict['Asking Value Found'][row - OFFSET]
        if 'Asking Value Link' in dict:
            ws[f'V{row}'] = dict['Asking Value Link'][row - OFFSET]
        row += 1

    wb.save('Equipment New List.xlsx')

def main():
    # output notes from program
    file = open('temp_output.txt', 'w')
    file.write(f"Data Collection began at {datetime.now()}\n")

    # create a copy of the master
    original = r'Equipment Master List.xlsx'
    target = r'Equipment New List.xlsx'
    shutil.copyfile(original, target)

    # get data from 'Equipment New List.xlsx'
    data = getExcelValues()

    # get online prices
    search_terms = get_search_terms(data)
    dict = scrapeAskingValues(search_terms)
    dict.update(scrapeAuctionValues(search_terms))

    # writes price_data to 'Equipment New List.xlsx'
    dict['Search Terms'] = search_terms
    setExcel(dict)

    # close output file
    file.write(f"Finished at {datetime.now()}")
    file.close()

if __name__ == "__main__":
    main()
