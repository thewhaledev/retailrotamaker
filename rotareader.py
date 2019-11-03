from openpyxl import load_workbook
import random
import shutil

workbook = load_workbook(filename="rota.xlsx")

sheet = workbook.active

"""
!Important! These cells dictate where the program will look when
deciding who is in for a particular day. If a person is removed from
the rota, or another position is added, the corresponding cells must
also be added to these lists.
"""

name_cells = ["A22", "A23", "A24", "A25", "A26", "A27", "A29", "A30",
"A31", "A32", "A33", "A34", "A35", "A36", "A38", "A39", "A40", "A41",
"A42", "A43", "A45", "A46", "A47", "A48", "A51", "A52", "A53", "A54",
"A55", "A56", "A57", "A58", "A59", "A60"]

mon_cells = ["B22", "B23", "B24", "B25", "B26", "B27", "B29", "B30",
"B31", "B32", "B33", "B34", "B35", "B36", "B38", "B39", "B40", "B41",
"B42", "B43", "B45", "B46", "B47", "B48", "B51", "B52", "B53", "B54",
"B55", "B56", "B57", "B58", "B59", "B60"]

tue_cells = ["F22", "F23", "F24", "F25", "F26", "F27", "F29", "F30",
"F31", "F32", "F33", "F34", "F35", "F36", "F38", "F39", "F40", "F41",
"F42", "F43", "F45", "F46", "F47", "F48", "F51", "F52", "F53", "F54",
"F55", "F56", "F57", "F58", "F59", "F60"]

wed_cells = ["J22", "J23", "J24", "J25", "J26", "J27", "J29", "J30",
"J31", "J32", "J33", "J34", "J35", "J36", "J38", "J39", "J40", "J41",
"J42", "J43", "J45", "J46", "J47", "J48", "J51", "J52", "J53", "J54",
"J55", "J56", "J57", "J58", "J59", "J60"]

thu_cells = ["N22", "N23", "N24", "N25", "N26", "N27", "N29", "N30",
"N31", "N32", "N33", "N34", "N35", "N36", "N38", "N39", "N40", "N41",
"N42", "N43", "N45", "N46", "N47", "N48", "N51", "N52", "N53", "N54",
"N55", "N56", "N57", "N58", "N59", "N60"]

fri_cells = ["R22", "R23", "R24", "R25", "R26", "R27", "R29", "R30",
"R31", "R32", "R33", "R34", "R35", "R36", "R38", "R39", "R40", "R41",
"R42", "R43", "R45", "R46", "R47", "R48", "R51", "R52", "R53", "R54",
"R55", "R56", "R57", "R58", "R59", "R60"]

sat_cells = ["V22", "V23", "V24", "V25", "V26", "V27", "V29", "V30",
"V31", "V32", "V33", "V34", "V35", "V36", "V38", "V39", "V40", "V41",
"V42", "V43", "V45", "V46", "V47", "V48", "V51", "V52", "V53", "V54",
"V55", "V56", "V57", "V58", "V59", "V60"]

sun_cells = ["Z22", "Z23", "Z24", "Z25", "Z26", "Z27", "Z29", "Z30",
"Z31", "Z32", "Z33", "Z34", "Z35", "Z36", "Z38", "Z39", "Z40", "Z41",
"Z42", "Z43", "Z45", "Z46", "Z47", "Z48", "Z51", "Z52", "Z53", "Z54",
"Z55", "Z56", "Z57", "Z58", "Z59", "Z60"]

def employee_dict_maker(day_list):
    """
    There are two values held in this employee dictionary, which makes a dictionary of employees for each day.
    "name" = the name of the employee as found in the rota spreadsheet.
    "time" = the time they come in, which will be used to seperate those on part time hours.
    """
    employees = {}
    for i in range(len(name_cells)):
        if sheet[day_list[i]].value != '-' and sheet[day_list[i]].value != None and sheet[day_list[i]].value != 'AL':
            employees[sheet[name_cells[i]].value] = sheet[day_list[i]].value
    return employees

def full_time_list(dict):
    #This function makes a seperate list of the names of those on full time hours.
    full_timers = []
    for key in dict.keys():
        if dict[key] != 1130 or dict[key] != 1140:
            full_timers.append(key)
    return full_timers

def part_time_list(dict):
    #This function makes a seperate list of the names of those on part time hours.
    part_timers = []
    for key in dict.keys():
        if dict[key] == 1130 or dict[key] == 1140:
            part_timers.append(key)
    return part_timers

#Now all the functions are called for all the days.

mon_employees = employee_dict_maker(mon_cells)
tue_employees = employee_dict_maker(tue_cells)
wed_employees = employee_dict_maker(wed_cells)
thu_employees = employee_dict_maker(thu_cells)
fri_employees = employee_dict_maker(fri_cells)
sat_employees = employee_dict_maker(sat_cells)
sun_employees = employee_dict_maker(sun_cells)

mon_full_time = full_time_list(mon_employees)
tue_full_time = full_time_list(tue_employees)
wed_full_time = full_time_list(wed_employees)
thu_full_time = full_time_list(thu_employees)
fri_full_time = full_time_list(fri_employees)
sat_full_time = full_time_list(sat_employees)
sun_full_time = full_time_list(sun_employees)

mon_part_time  = part_time_list(mon_employees)
tue_part_time = part_time_list(tue_employees)
wed_part_time = part_time_list(wed_employees)
thu_part_time = part_time_list(thu_employees)
fri_part_time = part_time_list(fri_employees)
sat_part_time = part_time_list(sat_employees)
sun_part_time = part_time_list(sun_employees)

"""
!Important! This list holds each shop floor position in rank of their
importance. If needs change, or if shops open and close, this list will
need to be edited.

This is not perfect in its current form, and will require you to change
the rota when lunch covers are on the shop floor, and when full-time
staff are doing lunch covers.
"""

positions = ["M/S, 12PM", "M/S, 1PM", "M/S 2PM", "WL-Ww", "MEZ", "COVER WL-MEZ",
 "WW-WL", "TS", "COVER WW-TS", "M/S, 12PM", "M/S, 1PM", "M/S 2PM"]

"""
!Important! This list holds the employees who have till "privaleges". This
will have to be changed when such employees change.
"""

important_employees = ["Corrina Simpson", "Ignacy Jarvis", "Zeeshan Mccullough",
"Elif Patrick", "Gracie-Mai Murphy", "Kristen Pratt", "Reggie Fletcher",
"Alexandria Lawson",]

def position_filler(full_time_list, part_time_list):
    """
    This function makes a tuple holding each position and the name of the employee who will fill that position.
    It checks agaisnt three conditions, presented in this order:
    1) That an employee who cannot be on certain positions is not placed in those positions.
    2) That those who are in the part time list are not placed in positions that only full time employees can be.
    3) That there are also employees with "till privaleges" in the main location, where such privalges (refunds etc...) are needed.
    """
    while True:
        random.shuffle(full_time_list) # The full time list is shuffled continously until all the proceeding conditions are met.
        if full_time_list[3] == "Alexandria Lawson" or full_time_list[5] == "Alexandria Lawson" or full_time_list[6] == "Alexandria Lawson" or full_time_list[8] == "Alexandria Lawson":
            continue
        if full_time_list[3] in part_time_list or full_time_list[4] in part_time_list or full_time_list[5] in part_time_list or full_time_list[6] in part_time_list:
            continue
        if full_time_list[0] in important_employees or full_time_list[1] in important_employees or full_time_list[2] in important_employees:
            break
    day_merge = tuple(zip(positions, full_time_list))
    return day_merge

#The function is called here for each day.

monday = position_filler(mon_full_time, mon_part_time)
tuesday = position_filler(tue_full_time, tue_part_time)
wednesday = position_filler(wed_full_time, wed_part_time)
thursday = position_filler(thu_full_time, thu_part_time)
friday = position_filler(fri_full_time, fri_part_time)
saturday = position_filler(sat_full_time, sat_part_time)
sunday = position_filler(sun_full_time, sun_part_time)

#Prints the tuples for the purpose of visualisation and testing.

print(monday)
print(tuesday)
print(wednesday)
print(thursday)
print(friday)
print(saturday)
print(sunday)

#Creates a copy of the rota template, so that this does not need to be done manually.
shutil.copy("daily_rota_template.xlsx", "daily_rotas.xlsx")

#Loads the template as a new workbook.
new_workbook = load_workbook(filename="daily_rotas.xlsx")

#These cells dictate where the names should be written on the template.
position_cells = ["B8", "B10", "B12", "B17", "B18", "B19", "B21",
 "B22", "B23", "B9", "B11", "B13", "B14", "B15"]

def daily_rota_maker(sheetname, position_tuple):
    """
    This function places each person from the tuple in their respective cell on the rota sheet.
    The sheetname "if" statements fill in the day of the week on each rota.
    """
    sheet = new_workbook[sheetname]
    for i in range(len(position_tuple)):
        sheet[position_cells[i]] = position_tuple[i][1]
    if sheetname == "Mon":
        sheet["G2"] = "Monday"
    if sheetname == "Tue":
        sheet["G2"] = "Tuesday"
    if sheetname == "Wed":
        sheet["G2"] = "Wednesday"
    if sheetname == "Thu":
        sheet["G2"] = "Thursday"
    if sheetname == "Fri":
        sheet["G2"] = "Friday"
    if sheetname == "Sat":
        sheet["G2"] = "Saturday"
    if sheetname == "Sun":
        sheet["G2"] = "Sunday"
    new_workbook.save("daily_rotas.xlsx")

#The function is called again for each day of the week.

daily_rota_maker("Mon", monday)
daily_rota_maker("Tue", tuesday)
daily_rota_maker("Wed", wednesday)
daily_rota_maker("Thu", thursday)
daily_rota_maker("Fri", friday)
daily_rota_maker("Sat", saturday)
daily_rota_maker("Sun", sunday)

input("Press ENTER to exit")