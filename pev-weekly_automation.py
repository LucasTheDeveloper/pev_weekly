from openpyxl import load_workbook
from datetime import datetime, time, timedelta
from openpyxl.styles import Font, PatternFill, Alignment
import time as t

def copy_chart(source_sheet, target_sheet):
    #Iterate throgh drawing objects including charts on the source sheet
    for drawing in source_sheet._images:
        if drawing.drawing_type == "Chart 2":
        #create new chart on the target sheet 
            new_drawing = drawing.copy()

        #set the new chart anchor on the target sheet
            new_drawing.anchor = new_drawing.anchor.copy(target=target_sheet)

        #add the new chart to the target sheeet 
            target_sheet.add_image(new_drawing)

def create_new_sheet(wb):
    new_sheet_name = "created_sheet_" + datetime.now().strftime("%Y%m%d_%H%M%S")
    ws = wb.create_sheet(new_sheet_name)
    
    #get the previously added sheet (assuming it's the last sheet)
    previous_sheet_name = wb.sheetnames[-2] if len(wb.sheetnames) > 1 else None
    previous_sheet = wb[previous_sheet_name]

    if previous_sheet:
    # Copy data from the previous sheet to the new sheet, including style
        for row in previous_sheet.iter_rows(min_row=1, max_row=previous_sheet.max_row, min_col=1, max_col=previous_sheet.max_column):
         for cell in row:
             new_cell = ws[cell.coordinate]
             new_cell.value = cell.value

            # Copy cell formatting from previous sheet
             new_cell.font = Font(
                name=cell.font.name,
                size=cell.font.size,
                bold=cell.font.bold,
                italic=cell.font.italic,
                color=cell.font.color,
                underline=cell.font.underline,
                strikethrough=cell.font.strikethrough,
                vertAlign=cell.font.vertAlign,
            )
                # Set text wrapping to fit text within the cell
             new_cell.alignment = Alignment(wrapText=True)

             if cell.fill is not None:
                    new_cell.fill = PatternFill(start_color=cell.fill.start_color, end_color=cell.fill.end_color, fill_type=cell.fill.fill_type)
             
             ws.row_dimensions[cell.row].height = previous_sheet.row_dimensions[cell.row].height   
             ws.column_dimensions[cell.column_letter].width = previous_sheet.column_dimensions[cell.column_letter].width
             
    
    #copy chart from previous sheet
        copy_chart(previous_sheet,ws)


         
    #hide column D, I and N
    ws.column_dimensions['D'].hidden = True
    ws.column_dimensions['I'].hidden = True
    ws.column_dimensions['N'].hidden = True
    
    # Clearing the data in column C
    for row in ws.iter_rows(min_row=5, max_row=15, min_col=3, max_col=3):
        for cell in row:
            cell.value = None
    
    # Clearing the data in column F
    for row in ws.iter_rows(min_row=5, max_row=15, min_col=6, max_col=6):
        for cell in row:
            cell.value = None

    # Clearing the data in column G
    for row in ws.iter_rows(min_row=5, max_row=15, min_col=7, max_col=7):
        for cell in row:
            cell.value = None
    
     # Clearing the data in column J
    for row in ws.iter_rows(min_row=5, max_row=15, min_col=10, max_col=10):
        for cell in row:
            cell.value = None
    
     # Clearing the data in column L
    for row in ws.iter_rows(min_row=5, max_row=15, min_col=12, max_col=12):
        for cell in row:
            cell.value = None
    
     # Clearing the data in column M
    for row in ws.iter_rows(min_row=5, max_row=15, min_col=13, max_col=13):
        for cell in row:
            cell.value = None
    
     # Clearing the data in column O
    for row in ws.iter_rows(min_row=5, max_row=15, min_col=15, max_col=15):
        for cell in row:
            cell.value = None
    
     # Clearing the data in column P
    for row in ws.iter_rows(min_row=5, max_row=15, min_col=16, max_col=16):
        for cell in row:
            cell.value = None
    
     # Clearing the data in column V
    for row in ws.iter_rows(min_row=5, max_row=15, min_col=22, max_col=22):
        for cell in row:
            cell.value = None
    
     # Clearing the data in column Y
    for row in ws.iter_rows(min_row=5, max_row=15, min_col=25, max_col=25):
        for cell in row:
            cell.value = None
    
     # Clearing the data in column AB
    for row in ws.iter_rows(min_row=5, max_row=15, min_col=28, max_col=28):
        for cell in row:
            cell.value = None


    #copy data from B5 to B15 and paste it into C5 to C15
    for i, row in enumerate(ws.iter_rows(min_row=5, max_row=15, min_col=2, max_col=2), start=5):
        ws.cell(row=i, column=3).value = row[0].value
    
     # Clearing the data in column B
    for row in ws.iter_rows(min_row=5, max_row=15, min_col=2, max_col=2):
        for cell in row:
            cell.value = None


    wb.save("PeV Weekly Summary Report 2023.xlsx")


target_time = time(8,1)  # Set the target time to anytime
wb = load_workbook('PeV Weekly Summary Report 2023.xlsx') #loading a workbook
sheet1 = wb['sheet1']

while True:
    current_time = datetime.now().time()
    
    if current_time >= target_time:
        create_new_sheet(wb)
        print("New sheet created at", current_time)
        break
    
    t.sleep(60)  # Wait for 1 minute before checking again

print(wb)

