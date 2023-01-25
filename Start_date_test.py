import re
from datetime import datetime
import openpyxl
#Open the RPA_T1 workbook
wb2 = openpyxl.load_workbook('Sizmek_TS.xlsx')
#sheet2 = wb2.active
sh1 = wb2['1074464166']

#Matching with start date extracted with H column
def extract_date(string):
    date_string = re.search(r'\d{8}', string).group()
    date_object = datetime.strptime(date_string, "%Y%m%d")
    formatted_date = date_object.strftime("%m/%d/%Y")
    print("Start Date extracted: ",formatted_date)
    return formatted_date
def extract_start_date_matching(formatted_date, date_range,sh1):
    match_found = False
    for row in range(date_range[0][1], date_range[1][1] + 1):
        column = ord(date_range[0][0]) - ord('A') + 1
        cell = sh1.cell(row=row, column=column)
        if formatted_date != cell.value:
            continue
        else:
            if not match_found:
                match_found = True
            print(f"Matching start Date extracted in cell {cell.coordinate}: {formatted_date}")
    if not match_found:
        print(f"No matching start date found in the range "
              f"{date_range[0][0]}{date_range[0][1]} - {date_range[1][0]}{date_range[1][1]}")
#Test it with the code
extract_start_date_matching('04/18/2022',[('H',24),('H',29)],sh1)
