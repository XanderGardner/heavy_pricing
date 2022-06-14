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
def scrape1(search_terms):
    # constants and variables
    n = len(search_terms)
    a1 = [None] * n
    a2 = [None] * n
    a3 = [None] * n
    
    # nested function for threaded scraping
    def scrape_task1(index):
        driver = webdriver.Chrome(resource_path('./chromedriver_win32/chromedriver.exe')) 
        driver.get(f"https://usedequipmentguide.com/listings?query={search_terms[index]}")
        a2[index] = f"https://usedequipmentguide.com/listings?query={search_terms[index]}"
        try:
            ele = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "span.Span-hup779-0.sc-16afded-0.kzgLyd")
            )) # will only wait for first result
            time.sleep(1.5) # testing shows that an extra 1.5 seconds allows all results to finish
            elements = driver.find_elements(by=By.CSS_SELECTOR, value="span.Span-hup779-0.sc-16afded-0.kzgLyd")
            
            # pick most relavent result
            i = 0
            while i < len(elements):
                if elements[i].text != "AUCTION" and elements[i].text != "Price Unavailable":
                    a1[index] = elements[i].text
                    break
                i += 1
        
        finally:
            driver.quit()
    
    n1 = 45 # number of items to scrape
    s1 = 50 # save excel after every s1 elements scraped
    i = 0
    while i < n1:
        # run next set of threads
        threads = [None] * MAX_THREADS
        ti = 0
        while ti < MAX_THREADS and i < n1:
            threads[ti] = Thread(target=scrape_task1, args=(i,))
            threads[ti].start()
            i += 1
            ti += 1
        for j in range(ti):
            threads[j].join()

        # occasionally save what is found
        if i % s1 == 0:
            row_start = i-s1+OFFSET
            arr_col_strs = ["N", "O", "P"]
            arr_values = [a1[i-s1:i], a2[i-s1:i], a3[i-s1:i]]
            tempSetExcel(arr_values, arr_col_strs, row_start)

    # set found prices and return
    prices = {
        'Auction Value' : a1,
        'Market Value' : a2,
        'Asking Value' : a3
    }
    return prices

# arr_values is 2d array. Each item is an array representing data for a column 
# arr_col_strs is array of strings. arr_col_strs at index i is the col for arr_values at i
# tempSetExcel will set the values in the respective columns in 'Equipment New List.xlsx',
# starting at the row row_start
def tempSetExcel(arr_values, arr_col_strs, row_start):
    wb = pyxl.load_workbook('Equipment New List.xlsx')
    ws = wb.active

    for col_index in range(len(arr_col_strs)):
        row_xlsx = row_start
        for row_index in range(len(arr_values[col_index])):
            ws[f'{arr_col_strs[col_index]}{row_xlsx}'] = arr_values[col_index][row_index]
            row_xlsx += 1
    
    wb.save('Equipment New List.xlsx')

# sets given final prices in 'Equipment New List.xlsx'
def setExcelPrices(prices):
    wb = pyxl.load_workbook('Equipment New List.xlsx')
    ws = wb.active

    row = 0 + OFFSET
    while row < ws.max_row:
        ws[f'N{row}'] = prices['Auction Value'][row - OFFSET]
        ws[f'O{row}'] = prices['Market Value'][row - OFFSET]
        ws[f'P{row}'] = prices['Asking Value'][row - OFFSET]
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
    prices = scrape1(search_terms) # https://usedequipmentguide.com/

    # writes price_data to 'Equipment New List.xlsx'
    setExcelPrices(prices)

    # close output file
    file.write(f"Finished at {datetime.now()}")
    file.close()

if __name__ == "__main__":
    main()
