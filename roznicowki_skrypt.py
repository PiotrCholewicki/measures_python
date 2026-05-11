import openpyxl as xl
import random
from copy import copy
from openpyxl.worksheet.pagebreak import Break
from openpyxl.styles import Font
import os
from openpyxl.worksheet.cell_range import CellRange
from openpyxl import Workbook
import re

path1 = 'C:\\Users\\Piotr\\Desktop\\mieszkania\\skrypt_pp\\szablon_roz.xlsx' #ZMIEN LOKALIZACJE
path2 = 'C:\\Users\\Piotr\\Desktop\\mieszkania\\skrypt_pp\\DS1_roz.xlsx' #ZMIEN LOKALIZACJE

sheet_name = "DS1"  # ZMIEN NA WŁAŚCIWĄ NAZWĘ ARKUSZA

wb1 = xl.load_workbook(filename=path1)
ws1 = wb1.worksheets[0]


page_index = 0
HEADER_INDEX = 33
FOOTER_INDEX = 51
inner_index = HEADER_INDEX + 1 
global_index = inner_index
circuits_index = 1
last_room = ""

def add_circuit(circuit_name):
    global global_index, inner_index, circuits_index, page_index, ws2
    
    if(inner_index >= FOOTER_INDEX - 1):
        inner_index = HEADER_INDEX + 1
        page_index += 1
        copy_rows(ws1, ws2, page_index * 60)

    global_index = page_index * 60 + inner_index
    ws2.cell(row=global_index, column=1, value=circuits_index)
    ws2.cell(row=global_index, column=2, value="wyłącznik różnicowo-prądowy "+circuit_name+"  25A/0,03 A")
    ws2.cell(row=global_index, column=3, value="TAK")
    ws2.cell(row=global_index, column=5, value="=RAND()*14+15.5")
    ws2.cell(row=global_index, column=6, value="=RAND()*14+15.5")
    ws2.cell(row=global_index, column=7, value="TAK")
    ws2.cell(row=global_index, column=8, value="TAK")

    inner_index += 1
    circuits_index += 1
    print("Dodano: ", circuit_name)



def addRangeOfCircuits(text):
    #print(text)
    prefix = ""
    end_suffix = ""

    left, right = text.split('-')
    

    # --- przypadek A: zakres na końcu (np. "F01-04") ---
    match = re.match(r"^(.*?)(\d+)$", left)
    if match and right.isdigit():
        prefix = match.group(1)
        start_num_str = match.group(2)
        start_num = int(start_num_str)
        end_num = int(right)
        width = len(start_num_str)

        obwody = [prefix + str(i).zfill(width) for i in range(start_num, end_num + 1)]
        for o in obwody:
            add_circuit(o)
        #print("Dodano obwody:", obwody)
        return

    # --- przypadek B: zakres na początku (np. "01-10O7") ---
    if left.isdigit() and re.match(r"^(\d+)(.*)$", right):
        start_num_str = left
        start_num = int(start_num_str)
        match = re.match(r"^(\d+)(.*)$", right)
        end_num = int(match.group(1))
        suffix = match.group(2)
        width = len(start_num_str)

        obwody = [str(i).zfill(width) + suffix for i in range(start_num, end_num + 1)]
        for o in obwody:
            add_circuit(o)
        #print("Dodano obwody:", obwody)
        return

    # --- przypadek C: zakres po kropce (np. "F1.1-6" albo "1.1-4F1") ---
    match_left = re.match(r"^(.*?\.)?(\d+)(.*)$", left)
    match_right = re.match(r"^(\d+)(.*)$", right)
    
    if match_left and match_right:
        prefix = match_left.group(1) or ""
        start_num_str = match_left.group(2)
        middle_suffix = match_left.group(3)
        start_num = int(start_num_str)

        end_num = int(match_right.group(1))
        end_suffix = match_right.group(2) or middle_suffix
        width = len(start_num_str)
        
        obwody = [prefix + str(i).zfill(width) + end_suffix for i in range(start_num, end_num + 1)]
        for o in obwody:
            add_circuit(o)
        #print("Dodano obwody:", obwody)
        return

    # --- przypadek D: przedział + dodatkowy tekst (np. "F1-10 rząd 5") ---
    match = re.match(r"^(.*?)(\d+)-(\d+)(\s.*)$", text)
    if match:
        prefix = match.group(1)
        start_num = int(match.group(2))
        end_num = int(match.group(3))
        suffix = match.group(4)  # np. " rząd 5"
        width = len(match.group(2))

        obwody = [prefix + str(i).zfill(width) + suffix for i in range(start_num, end_num + 1)]
        for o in obwody:
            add_circuit(o)
        #print("Dodano obwody:", obwody)
        return

    print(" Nie rozpoznano wzorca:", text)


def add_header(header_text):
    global global_index, inner_index, page_index, ws2, last_room, circuits_index
    if(inner_index >= FOOTER_INDEX - 2):
        inner_index = HEADER_INDEX + 1
        page_index += 1
        copy_rows(ws1, ws2, page_index * 60)
    
    global_index = page_index * 60 + inner_index

    
    ws2.cell(row=global_index, column=1, value=header_text).font = Font(bold=True, size=48)

    ws2.merge_cells(start_row=global_index, start_column=1, end_row=global_index, end_column=8)
    ws2.cell(row=global_index, column=1).alignment = xl.styles.Alignment(horizontal='center', vertical='center')
    inner_index += 1
    last_room = "ROZDZIELNICA " + header_text
    circuits_index = 1

def add_ds6():
    circuit_list = []
    for i in range(1, 9, 2):
        add_circuit(str(i))
        circuit_list.append(str(i))
    print("Dodano obwody:", circuit_list)
#the same as in previous files
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
        
    for mr in source_sheet.merged_cells.ranges:
        # Tworzymy nowy CellRange na bazie oryginalnych współrzędnych
        cr = CellRange(mr.coord)
        # Uwaga: shift() modyfikuje obiekt i ZWRACA None
        cr.shift(0, offset)
        try:
            target_sheet.merge_cells(cr.coord)
        except ValueError:
            # pomiń, jeśli zakres już jest scalony / nachodzi na inne scalenie
            pass

    ws = target_sheet
    row_number = source_sheet.max_row + offset  # the row that you want to insert page break

    # usuń wszystkie istniejące page breaki w wierszach
    
    page_break = Break(id=row_number)  # create Break obj
    ws.row_breaks.append(page_break)  # insert page break
    ws.col_breaks.append

def mainloop():
    global global_index, inner_index, page_index, ws2, last_room
    while True:
        prompt = input("Podaj nazwę obwodu/rozdzielnicy/STOP: ").upper()
        
        if(prompt == "SAVE" or prompt == "STOP"):
            wb2.save(path2)
            print("Zapisano plik")
            global_index = page_index * 60 + inner_index
            with open('C:\\Users\\Piotr\\Desktop\\mieszkania\\skrypt_pp\\last_index.txt', 'w') as file:
                file.write(str(global_index))
            with open('C:\\Users\\Piotr\\Desktop\\mieszkania\\skrypt_pp\\last_room.txt', 'w') as file:
                file.write(last_room)
            break

        isValid = re.match(r"^T[A-Z0-9]", prompt)
        
        if(prompt[0:2] == "R "):
            add_header("ROZDZIELNICA " + prompt[2:])
            print("Dodano rozdzielnicę: ", prompt[2:])
            last_room = "ROZDZIELNICA " + prompt[2:]
        elif(prompt==""):
            add_ds6()
        elif(prompt[0:2] == "P "):
            #start new floor on the new page
            if(inner_index > 35):
                inner_index = FOOTER_INDEX
            add_header("PIĘTRO " + prompt[2:])
            print("Dodano piętro: ", prompt[2:])
            last_room = "PIĘTRO " + prompt[2:]
        elif(prompt[0:2] == "O "):
            
            #for cases like o tlk1.1-4
            dashOcurred = False
            # circuit case
            for i in range(len(prompt)):
                if(prompt[i] == "-"):
                    dashOcurred = True
            if(dashOcurred):
                #print(prompt[2:])
                addRangeOfCircuits(prompt[2:])
            else:
                #print(prompt[2:])
                add_circuit(prompt[2:])
        elif(isValid):
            #for example: TK 1.4.3 without r in the beginning, using regEx above
            add_header("ROZDZIELNICA "+prompt)
            print("Dodano rozdzielnicę: ", prompt)
            last_room = "ROZDZIELNICA " + prompt

        else:
            dashOcurred = False
            # circuit case
            for i in range(len(prompt)):
                if(prompt[i] == "-"):
                    dashOcurred = True
            if(dashOcurred):
                addRangeOfCircuits(prompt)
            else:
                add_circuit(prompt)
                

        if(global_index % 5 == 0):           
            wb2.save(path2)
        global_index = page_index * 60 + inner_index
        
        with open('C:\\Users\\Piotr\\Desktop\\mieszkania\\skrypt_pp\\last_index.txt', 'w') as file:
            file.write(str(global_index))
        with open('C:\\Users\\Piotr\\Desktop\\mieszkania\\skrypt_pp\\last_room.txt', 'w') as file:
            file.write(last_room)
            
            
file_path = "C:\\Users\\Piotr\\Desktop\\mieszkania\\skrypt_pp\\DS1_roz.xlsx"


if not os.path.exists(file_path):
    wb = Workbook()
    wb.save(file_path)
    print("Utworzono plik z różnicówkami:", file_path)



if not os.path.exists('C:\\Users\\Piotr\\Desktop\\mieszkania\\skrypt_pp\\last_index.txt'):
    print("No last index file found, starting fresh.")
    
    wb2 = xl.load_workbook(filename=path2)
    ws2 = wb2.create_sheet(sheet_name)
    ws2.page_setup.scale = 15
    ws2.column_dimensions['A'].width = 27
    ws2.column_dimensions['B'].width = 140
    ws2.column_dimensions['C'].width = 38
    ws2.column_dimensions['D'].width = 28
    ws2.column_dimensions['E'].width = 52
    ws2.column_dimensions['F'].width = 45
    ws2.column_dimensions['G'].width = 51
    ws2.column_dimensions['H'].width = 73
    ws2.column_dimensions['I'].width = 1
    copy_rows(ws1, ws2, 0)
    #add_header(sheet_name)
    mainloop()

else:
    print(f"The file 'last_index.txt' already exists.")
    with open('C:\\Users\\Piotr\\Desktop\\mieszkania\\skrypt_pp\\last_index.txt', 'r') as file:
        global_index = int(file.read().strip())
        print(f"Last index loaded: {global_index}")
        
    with open('C:\\Users\\Piotr\\Desktop\\mieszkania\\skrypt_pp\\last_room.txt', 'r') as file:
        last_room = file.read().strip()
        print(f"Last room loaded: {last_room}")

    wb2 = xl.load_workbook(filename=path2)
    
    inner_index = global_index % 60
    page_index = global_index // 60
    if sheet_name in wb2.sheetnames:
        
        ws2 = wb2[sheet_name]
        for i in range(1260):
            inner_index_loop = i % 60
            if(inner_index_loop > HEADER_INDEX and inner_index_loop < FOOTER_INDEX):
                ws2.row_dimensions[i].height = 113
        mainloop()
    else:
        ws2 = wb2.create_sheet(sheet_name)




