import os
from openpyxl import Workbook, load_workbook

#Verificar que exista el archivo
if os.path.exists("job_tracker.xlsx"):
    wb = load_workbook(filename="job_tracker.xlsx")
else:
    wb = Workbook()
#Obtener spreadsheet
part_time = wb.active
part_time.title = "Part Time Jobs"
swe = wb.create_sheet("SWE Jobs", 1)

#Nombrar celdas del part time
part_time['A1'] = "Company Title"
part_time['B1'] = "Position"
part_time['C1'] = "Date Applied"

#Nombrar celdas del SWE job
swe['A1'] = "Company Title"
swe['B1'] = "Position"
swe['C1'] = "Date Applied"

#Revisar nombres de las celdas
# for col in ws.iter_cols(min_row=1, max_col=3, max_row=1):
#     for cell in col:
#         print(cell.value) Imprime el nombre de las celdas

#Crear funcion que convierta el input del usuario en lower case
def convert_to_lower(type_of_sheet: str) -> str:
    if len(type_of_sheet) == 0:
        return "The string passed is 0"
    return type_of_sheet.lower()

#Crear funcion que detecta si la cadena esta vacia
def empty_string(type_of_sheet: str) -> bool:
    if len(type_of_sheet) == 0:
        return False
    return True

#Pedir que tipo de sheet va a rellenar
while True:
    sheet_type = input("Please enter which spreadsheet you want to populate (Part Time Jobs/ SWE Jobs: ")
    sheet_type_lower = convert_to_lower(sheet_type)

    if sheet_type_lower == "part time jobs":
        active_sheet = part_time
        break
    elif sheet_type_lower == "swe jobs":
        active_sheet = swe
        break
    else:
        print("Only 'Part Time Jobs' or 'SWE Jobs' sheets exist. Please try again ")

company_title = input("Type company name: ")
position = input("Please type the position name: ")
date_applied = input("Please type the date in which you applied: ")

#Conseguir las celdas como keys del part time
company_part = part_time['A2']
pos_part = part_time['B2']
date_part = part_time['C2']

#Conseguir las celdas como keys del swe
company_swe = swe['A2']
pos_swe = swe['B2']
date_swe = swe['C2']

#Rellenar celdas con la entrada de datos del usuario


# for col in ws.iter_cols(min_col=1, max_col=3, min_row=1, max_row=2):
#     for cell in col:
#         print(cell.value) Imprime los valores del usuario

