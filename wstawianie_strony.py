import openpyxl as xl
from copy import copy
from openpyxl.worksheet.pagebreak import Break
path1 = 'C:\\Users\\Piotr\\Desktop\\mieszkania\\szablon.xlsx' #ZMIEN LOKALIZACJE
path2 = 'C:\\Users\\Piotr\\Desktop\\mieszkania\\test.xlsx' #ZMIEN LOKALIZACJE

wb1 = xl.load_workbook(filename=path1)
ws1 = wb1.worksheets[0]

wb2 = xl.load_workbook(filename=path2)
ws2 = wb2.create_sheet(ws1.title)

ws2.column_dimensions['B'].width = 41


class Strona:
    def __init__(self, gk, gl, ob_1, ob_3): #ob - obwody, gk - gniazdka w kuchni, gl - gniazdka w lazience,
        self.gk = gk
        self.gl = gl
        self.ob_1 = ob_1
        self.ob_3 = ob_3

# Funkcja do kopiowania wierszy wraz z formatowaniem i dodawaniem podziału strony
def copy_rows(source_sheet, target_sheet, start_row, end_row, offset, nr_klatki, nr_mieszkania):

    for row in source_sheet.iter_rows(min_row=start_row, max_row=end_row):
        for cell in row:
            target_cell = target_sheet.cell(row=cell.row + offset, column=cell.column, value=cell.value)
            target_cell.font = copy(cell.font)
            target_cell.border = copy(cell.border)
            target_cell.fill = copy(cell.fill)
            target_cell.number_format = copy(cell.number_format)
            target_cell.protection = copy(cell.protection)
            target_cell.alignment = copy(cell.alignment)
            
    target_sheet.cell(row=2 + offset, column=7, value=nr_klatki)
    target_sheet.cell(row=7 + offset, column=9, value=nr_klatki)
    target_sheet.cell(row=2 + offset, column=8, value=nr_mieszkania)
    target_sheet.cell(row=7 + offset, column=10, value=nr_mieszkania)

    ws = target_sheet
    row_number = end_row+offset  # the row that you want to insert page break
    page_break = Break(id=row_number)  # create Break obj
    ws.row_breaks.append(page_break)  # insert page break
    ws.col_breaks.append
#funkcja uzupelniajaca strony na podstawie danych wprowadzonych przez użytkownika
def uzupelnij_strony(strona, offset):
    target_sheet = ws2
    index = 0
    index_ob = 0
    for i in range(strona.gk):
        target_sheet.cell(row=13 + offset + index, column=1, value=index+1)
        target_sheet.cell(row=13 + offset + index, column=2, value='Gn pt 2P+Z,16A,250V kuchnia ')
        index+=1
    for i in range(strona.gl):
        target_sheet.cell(row=13 + offset + index, column=1, value=index+1)
        target_sheet.cell(row=13 + offset + index, column=2, value='Gn nt hermet 2P+Z,16A,250V łazienka ')
        index+=1

    for i in range(ob_1): #1-fazowe
        target_sheet.cell(row=49 + offset + index_ob, column=1, value=index_ob+1)
        target_sheet.cell(row=49 + offset + index_ob, column=2, value='Obwod 1-fazowy ')
        index_ob += 1
    for i in range(ob_3): #3-fazowe
        target_sheet.cell(row=49 + offset + index_ob, column=1, value=index_ob+1)
        target_sheet.cell(row=49 + offset + index_ob, column=2, value='Obwod 3-fazowy ')
        index_ob += 1
#funkcja która przyjmuje input typu "3+2", zamienia to na 3 obwody 1 fazowe, i 2 obwody 3-fazowe, 
#zwraca tablice w ktorej tab[0] to 1-fazowe, tab[1] to 3 fazowe

def podaj_obwody(wyrazenie):
    tab = [0,0]
    i = 0
    for _ in wyrazenie:
        if wyrazenie[i] == '+':
            break
        else:
            i+=1
    tab[0] = int(wyrazenie[0:i])
    tab[1] = wyrazenie[i+1:]
    if(tab[1] ==''):
        tab[1] = 0
    tab[1] = int(tab[1])
    return tab

i = 0
nr_klatki = 0
nr_mieszkania = 0
while True:
    if(nr_klatki == 0 or nr_mieszkania == 0):
        print("Podaj wartości dla pierwszej strony: ")
    else:
        print("\nOstatnio utworzono: "+nr_klatki+"/"+nr_mieszkania)
    nr_klatki = input("Podaj nr klatki lub napisz stop ")
    if(nr_klatki == "STOP" or nr_klatki == "stop" or nr_klatki == "nie" or nr_klatki == "NIE" or nr_klatki == "n" or nr_klatki == "N"):
        break
    nr_mieszkania = input("Podaj nr mieszkania ")
    copy_rows(ws1, ws2, 1, ws1.max_row, ws1.max_row * i, nr_klatki, nr_mieszkania)
    ob = input("Podaj ilosc obwodow: ")
    #korzystajac z funkcji podaj_obwody przepisywane są wartości obwodow 1 i 3 fazowych do dwoch osobnych zmiennych
    tab = podaj_obwody(ob)
    ob_1 = tab[0] 
    ob_3 = tab[1]
    gk = int(input("Podaj ilosc gniazdek w kuchni "))
    gl = int(input("Podaj ilosc gniazdek w łazience "))
    
    strona = Strona(gk, gl, ob_1, ob_3)
    uzupelnij_strony(strona, ws1.max_row*i)
    # Kopiowanie scalonych komórek - kolejne strony
    for merged_cell_range in ws1.merged_cells.ranges:
        min_row, min_col, max_row, max_col = merged_cell_range.min_row, merged_cell_range.min_col, merged_cell_range.max_row, merged_cell_range.max_col
        ws2.merge_cells(start_row=min_row + ws1.max_row * i, start_column=min_col, end_row=max_row + ws1.max_row * i, end_column=max_col)
    i += 1

wb2.save(path2)
