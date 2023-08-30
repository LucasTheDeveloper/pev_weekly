from openpyxl import load_workbook
from datetime import datetime, time, timedelta
import time as t

def create_new_sheet(wb):
    new_sheet_name = "created_sheet_" + datetime.now().strftime("%Y%m%d_%H%M%S")
    ws = wb.create_sheet(new_sheet_name)
    
    for i in range(1, sheet1.max_row + 1):
        for j in range(1, sheet1.max_column +1):
            ws.cell(row=i, column=j).value = sheet1.cell(row=i, column=j).value
            
    wb.save("auto.xlsx")


target_time = time(13, 11)  # Set the target time to 09:40
wb = load_workbook('auto.xlsx')
sheet1 = wb['sheet1']

while True:
    current_time = datetime.now().time()
    
    if current_time >= target_time:
        create_new_sheet(wb)
        print("New sheet created at", current_time)
        break
    
    t.sleep(60)  # Wait for 1 minute before checking again

print(wb)

