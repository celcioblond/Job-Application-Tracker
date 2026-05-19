import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

#Verificar que exista el archivo
if os.path.exists("job_tracker.xlsx"):
    wb = load_workbook(filename="job_tracker.xlsx")
    part_time = wb["Part Time Jobs"]
    swe= wb["SWE Jobs"]
else:
    wb = Workbook()
    part_time = wb.active
    part_time.title = "Part Time Jobs"
    swe = wb.create_sheet("SWE Jobs", 1)
    part_time.append(["Company Title", "Position", "Date Applied"])
    swe.append(["Company Title", "Position", "Date Applied"])

    for sheet in [part_time, swe]:
        for cell in sheet[1]:
            cell.font = Font(bold=True)

    for sheet in [part_time, swe]:
        for col in sheet.columns:
            max_length = max(len(str(cell.value or "")) for cell in col)
            sheet.column_dimensions[col[0].column_letter].width = max_length + 4

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
    sheet_type = input("Please enter which spreadsheet you want to populate (Part Time Jobs/ SWE Jobs): ")
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

active_sheet.append([company_title, position, date_applied])

print("Application registered succesfully")

#Guardar workbook
wb.save("job_tracker.xlsx")

