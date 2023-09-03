from openpyxl import load_workbook
from datetime import datetime, time, timedelta
import time as t

def create_new_sheet(wb):
    new_sheet_name = "created_sheet_" + datetime.now().strftime("%Y%m%d_%H%M%S")
    ws = wb.create_sheet(new_sheet_name)
    
    #get the previously added sheet (assuming it's the last sheet)
    previous_sheet_name = wb.sheetnames[-2] if len(wb.sheetnames) > 1 else None
    previous_sheet = wb[previous_sheet_name]

    if previous_sheet:
        #copy data from the previous sheet to the new sheet
        for i in range(1, previous_sheet.max_column +1):
            for j in range(1, previous_sheet.max_column +1):
                ws.cell(row = i, column = j).value = previous_sheet.cell(row=i , column=j).value 
    
    # Clearing the data in specified cell ranges
    for row in ws.iter_rows(min_row=2, max_row=5, min_col=3, max_col=3):
        for cell in row:
            cell.value = None

    for row in ws.iter_rows(min_row=2, max_row=5, min_col=7, max_col=7):
        for cell in row:
            cell.value = None
    #copy data from a certain column and replace it on another coulumn
    for i, row in enumerate(ws.iter_rows(min_row=2, max_row=5, min_col=1, max_col=1), start=2):
        ws.cell(row=i, column=3).value = row[0].value

    wb.save("auto.xlsx")


target_time = time(11, 11)  # Set the target time to anytime
wb = load_workbook('auto.xlsx') #loading a workbook
sheet1 = wb['sheet1']

while True:
    current_time = datetime.now().time()
    
    if current_time >= target_time:
        create_new_sheet(wb)
        print("New sheet created at", current_time)
        break
    
    t.sleep(60)  # Wait for 1 minute before checking again

print(wb)

