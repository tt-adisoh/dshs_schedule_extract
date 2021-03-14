from openpyxl import Workbook, load_workbook
import json
from collections import OrderedDict

xlsx_file = load_workbook("extracted.xlsx")
xlsx_sheet = xlsx_file['Sheet1']
Monday_alphabet = ['E','F','G','H','I','J','K']
Tuesday_alphabet = ['L','N','M','O','P','Q','R']
Wednesday_alphabet = ['S','T','U','V']
Thursday_alphabet = ['W','X','Y','Z','AA','AB','AC']
Friday_alphabet = ['AD','AE','AF','AG','AH','AI','AJ']
total_sum = []

for number_of_students in range (1,345):
    excel_data = {}
    excel_data['student_number'] = xlsx_sheet[str("A" + str(number_of_students))].value
    excel_data['id'] = xlsx_sheet[str("B" + str(number_of_students))].value
    excel_data['name'] = xlsx_sheet[str("D" + str(number_of_students))].value
    Monday = {}
    Tuesday = {}
    Wednesday = {}
    Thursday = {}
    Friday = {}

    for mon in range (len(Monday_alphabet)):
        Monday[mon+1] = xlsx_sheet[str(Monday_alphabet[mon] + str(number_of_students))].value
    for tues in range (len(Tuesday_alphabet)):
        Tuesday[tues+1] = xlsx_sheet[str(Tuesday_alphabet[tues] + str(number_of_students))].value
    for wed in range (len(Wednesday_alphabet)):
        Wednesday[wed+1] = xlsx_sheet[str(Wednesday_alphabet[wed] + str(number_of_students))].value
    for thurs in range (len(Thursday_alphabet)):
        Thursday[thurs+1] = xlsx_sheet[str(Thursday_alphabet[thurs] + str(number_of_students))].value
    for fri in range (len(Friday_alphabet)):
        Friday[fri+1] = xlsx_sheet[str(Friday_alphabet[fri] + str(number_of_students))].value
    
    excel_data['Monday'] = Monday
    excel_data['Tuesday'] = Tuesday
    excel_data['Wednesday'] = Wednesday
    excel_data['Thursday'] = Thursday
    excel_data['Friday'] = Friday

    total_sum.append(excel_data)

# 최종 json 파일 생성
with open('data.json', 'w', encoding='utf-8') as make_file:
    json.dump(total_sum, make_file, ensure_ascii=False, indent="\t")