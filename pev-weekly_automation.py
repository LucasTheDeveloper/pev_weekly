from openpyxl import load_workbook
from datetime import datetime, time, timedelta
import time as t

def create_new_sheet(wb):
    new_sheet_name = "created_sheet_" + datetime.now().strftime("%Y%m%d_%H%M%S")
    ws = wb.create_sheet(new_sheet_name)
    
    #iteration for loop to copy data on the selected data range
    for i in range(1, sheet1.max_row + 1):
        for j in range(1, sheet1.max_column +1):
            ws.cell(row=i, column=j).value = sheet1.cell(row=i, column=j).value 
    
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


target_time = time(13, 11)  # Set the target time to anytime
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

