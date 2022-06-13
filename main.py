import os
import sys
import openpyxl as pyxl
import shutil
from selenium import webdriver
from datetime import datetime


def getExcelValues():
    # first row is table headers
    # last row is totaled information

    wb = pyxl.load_workbook('Equipment New List.xlsx')
    ws = wb.active
    n = ws.max_row - 2

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

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

# https://www.zacoding.com/en/post/python-selenium-to-exe/
def scrape(data):
    # constants
    driver = webdriver.Chrome(resource_path('./chromedriver_win32/chromedriver.exe'))    
    n = len(data['Emco'])
    a1 = [None] * n
    a2 = [None] * n
    a3 = [None] * n

    # scrape data to fill Auction Values (for a1)
    driver.get("https://usedequipmentguide.com/")

    for i in range(n):
        terms = []
        if data['Manufacturer'][i]:
            terms.append(data['Manufacturer'][i])
        if data['Model'][i]:
            terms.append(data['Model'][i])
        if data['ModelYr'][i]:
            terms.append(data['ModelYr'][i])
        search_term = ' '.join(str(term) for term in terms)
        a2[i] = len(search_term)
        if len(search_term) <= 8 and data['Description'][i]:
            a1[i] = f"{search_term} {data['Description'][i]}"
        else:
            a1[i] = search_term

    prices = {
        'Auction Value' : a1,
        'Market Value' : a2,
        'Asking Value' : a3
    }
    return prices

def setExcelPrices(prices):
    # first row is table headers
    # last row is totaled information

    wb = pyxl.load_workbook('Equipment New List.xlsx')
    ws = wb.active

    print('Auction Value')
    print(prices['Auction Value'][0])

    row = 2
    while row < ws.max_row:
        ws[f'N{row}'] = prices['Auction Value'][row - 2]
        ws[f'O{row}'] = prices['Market Value'][row - 2]
        ws[f'P{row}'] = prices['Asking Value'][row - 2]
        row += 1

    wb.save('Equipment New List.xlsx')

    return

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

    # get online prices from data
    prices = scrape(data)

    # writes price_data to 'Equipment New List.xlsx'
    setExcelPrices(prices)

    # close output file
    file.write(f"Finished at {datetime.now()}")
    file.close()

if __name__ == "__main__":
    main()

