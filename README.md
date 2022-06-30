# heavy_pricing
Tool for automatically pricing heavy machinery according to current online bid prices. 

## Input:
- Excel File named "Equipment Master List.xlsx" with columns A, B, C, ... being "EMCo", "Description", "VINNumber", "Manufacturer", "Model", "ModelYr", "OdoReading", "OdoDate", "HourReading", "HourDate", "Location", and "Complete"

## Output:
- Excel File named "Equipment New List.xslx" which is a copy of "Equipment Master List.xslx" with additional columns for "Auction Value", "Market Value", and "Asking Value"
- Text File name "temp_output.txt" for outputting results of running the executable

## Creating the Executable for Windows:
https://www.zacoding.com/en/post/python-selenium-to-exe/
```
pyinstaller ./main.py --onefile --add-binary "./chromedriver_win32/chromedriver.exe;./chromedriver_win32"
```
