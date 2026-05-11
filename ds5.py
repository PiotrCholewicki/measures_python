import openpyxl as xl
import random
from copy import copy
from openpyxl.worksheet.pagebreak import Break
from openpyxl.styles import Font
import os

path1 = 'C:\\Users\\Piotr\\Desktop\\mieszkania\\skrypt_pp\\szablon.xlsx' #ZMIEN LOKALIZACJE
path2 = 'C:\\Users\\Piotr\\Desktop\\mieszkania\\skrypt_pp\\ds5.xlsx' #ZMIEN LOKALIZACJE

sheet_name = "DS5"  # ZMIEN NA WŁAŚCIWĄ NAZWĘ ARKUSZA

wb1 = xl.load_workbook(filename=path1)
ws1 = wb1.worksheets[0]

sockets_index = 1
page_index = 0
HEADER_INDEX = 12
FOOTER_INDEX = 58
inner_index = HEADER_INDEX + 1 
global_index = inner_index
room_index = 1

def add_sockets(socket_name, sockets_count):
    global global_index, inner_index, sockets_index, page_index, ws2
    for _ in range(sockets_count):
        if(inner_index >= 57):
            inner_index = HEADER_INDEX + 1
            page_index += 1
            copy_rows(ws1, ws2, page_index * 60)
        global_index = page_index * 60 + inner_index
        ws2.cell(row=global_index, column=1, value=sockets_index)
        ws2.cell(row=global_index, column=2, value=socket_name)
        ws2.cell(row=global_index, column=3, value="SBinst")
        ws2.cell(row=global_index, column=4, value="B16")
        ws2.cell(row=global_index, column=5, value=230.00)
        ws2.cell(row=global_index, column=6, value="=RAND()*0.35+0.9")
        ws2.cell(row=global_index, column=7, value=80)
        ws2.cell(row=global_index, column=8, value=f"=F{global_index}*G{global_index}")  
        ws2.cell(row=global_index, column=9, value="TAK")
        inner_index += 1
        sockets_index += 1
    
   

def add_header(header_text):
    global global_index, inner_index, page_index, ws2
    if(inner_index >= 56):
        inner_index = HEADER_INDEX + 1
        page_index += 1
        copy_rows(ws1, ws2, page_index * 60)
    
    global_index = page_index * 60 + inner_index
    ws2.cell(row=global_index, column=1, value=header_text).font = Font(bold=True, size=16)
    ws2.cell(row=global_index, column=1).alignment = xl.styles.Alignment(horizontal='center')
    ws2.merge_cells(start_row=global_index, start_column=1, end_row=global_index, end_column=9)
    inner_index += 1

def copy_rows(source_sheet, target_sheet, offset):

    for row in source_sheet.iter_rows(min_row=1, max_row=source_sheet.max_row):
        for cell in row:
            target_cell = target_sheet.cell(row=cell.row + offset, column=cell.column, value=cell.value)
            target_cell.font = copy(cell.font)
            target_cell.border = copy(cell.border)
            target_cell.fill = copy(cell.fill)
            target_cell.number_format = copy(cell.number_format)
            target_cell.protection = copy(cell.protection)
            target_cell.alignment = copy(cell.alignment)
            

    ws = target_sheet
    row_number = source_sheet.max_row + offset  # the row that you want to insert page break
    page_break = Break(id=row_number)  # create Break obj
    ws.row_breaks.append(page_break)  # insert page break
    ws.col_breaks.append

def mainloop():
    global global_index, inner_index, page_index, ws2, room_index, sockets_index
    for floor in range(1, 11):
        room_index = 1
        sockets_index = 1
        floor_name = floor


        #add floor name in the beggining
        add_header(f"PIĘTRO {floor_name}")

        #add dryer in each floor
        add_header("SUSZARNIA PIĘTRO " + str(floor_name))
        add_sockets("Gniazdko 1-fazowe", 8)
        sockets_index = 1

        #add laundry in each floor
        if (floor%2 == 1):
            add_header("PRALNIA PIĘTRO " + str(floor_name))
            add_sockets("Gniazdko 1-fazowe", 8)
            sockets_index = 1


        for room in range(1, 25):

            # Check if the template isn't fullfilled
            if(inner_index >= 56):
                inner_index = HEADER_INDEX + 1
                page_index += 1
                global_index = page_index * 60 + inner_index
                copy_rows(ws1, ws2, page_index * 60)

            #first add the floor name
            global_index = page_index * 60 + inner_index
            if(int(room_index) < 10):
                room_index_str = f"0{room_index}"
            else:
                room_index_str = str(room_index)
            room_name = "POKÓJ " + str(floor_name) + room_index_str
            add_header(room_name)

            if(room_index % 2 == 0):
                sockets = 4
            else:
                sockets = 8
            

            if(room_index % 2 == 0):
                add_sockets("Gniazdko 1-fazowe", 4)
                add_sockets("Gniazdko 1-fazowe kuchnia", 6)
                add_sockets("Gniazdko 1-fazowe łazienka", 3)
                sockets_index = 1
            else:
                add_sockets("Gniazdko 1-fazowe", 8)
                sockets_index = 1
            room_index += 1
        sockets_index = 1
        inner_index = 57
                
        wb2.save(path2)
        global_index = page_index * 60 + inner_index
        sockets_index = 1
        wb2.save(path2)
        print("Zapisano piętro nr ", floor_name)

            

wb2 = xl.load_workbook(filename=path2)
ws2 = wb2.create_sheet(sheet_name)
#initial copy of a template, if it's the first run
ws2.column_dimensions['A'].width = 8.43
ws2.column_dimensions['B'].width = 40.43
ws2.column_dimensions['C'].width = 20.43
ws2.column_dimensions['D'].width = 9.57
ws2.column_dimensions['E'].width = 23.29
ws2.column_dimensions['F'].width = 16.14
ws2.column_dimensions['G'].width = 16.86
ws2.column_dimensions['H'].width = 8.8
ws2.column_dimensions['I'].width = 13.14
copy_rows(ws1, ws2, 0)
mainloop()




