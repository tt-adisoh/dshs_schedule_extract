from openpyxl import Workbook, load_workbook
import json
from collections import OrderedDict

# 파일 불러오기 & 초기 설정
xlsx_file = load_workbook("datafile.xlsx")
xlsx_sheet = xlsx_file['개인별 시간표']
subject_alphabet = ['B','G','H','I','J','K','L','N','M','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG']
total_sum = []

# 수강 과목 추출
for i in range (390,737):
    k = 0
    file_data = OrderedDict()
    final_data = OrderedDict()
    file_data["id"] = xlsx_sheet[str(subject_alphabet[k] + str(i))].value
    final_data["id"] = xlsx_sheet[str(subject_alphabet[k] + str(i))].value
    k = k+1
    file_data["name"] = xlsx_sheet[str(subject_alphabet[k] + str(i))].value
    final_data["name"] = xlsx_sheet[str(subject_alphabet[k] + str(i))].value
    k = k+1
    for p in range (25):
        if xlsx_sheet[str(subject_alphabet[k] + str(i))].value != None:
            file_data[str(xlsx_sheet[str(subject_alphabet[k] + str(i))].value)] = str(xlsx_sheet[str(subject_alphabet[k] + '389')].value)
        k = k+1
    file_data["0"] = "창체"

    # 구성에 따른 요일별 시간표 입력
    Monday = {}
    Monday[1] = file_data.get('E')
    Monday[2] = file_data.get('D')
    Monday[3] = file_data.get('B')
    Monday[4] = file_data.get("F")
    Monday[5] = file_data.get('C')
    Monday[6] = file_data.get("G")
    Monday[7] = file_data.get('H')

    Tuesday = {}
    Tuesday[1] = file_data.get('F')
    Tuesday[2] = file_data.get('D')
    Tuesday[3] = file_data.get('E')
    Tuesday[4] = file_data.get("B")
    Tuesday[5] = file_data.get('E')
    Tuesday[6] = file_data.get("G")
    Tuesday[7] = file_data.get('B')

    Wednesday = {}
    Wednesday[1] = file_data.get('F')
    Wednesday[2] = file_data.get('C')
    Wednesday[3] = file_data.get('H')
    Wednesday[4] = file_data.get("A")
    Wednesday[5] = file_data.get('0')
    Wednesday[6] = file_data.get("0")
    Wednesday[7] = file_data.get('0')

    Thursday = {}
    Thursday[1] = file_data.get('C')
    Thursday[2] = file_data.get('E')
    Thursday[3] = file_data.get('H')
    Thursday[4] = file_data.get("A")
    Thursday[5] = file_data.get('D')
    Thursday[6] = file_data.get("G")
    Thursday[7] = file_data.get('G')

    Friday = {}
    Friday[1] = file_data.get('H')
    Friday[2] = file_data.get('A')
    Friday[3] = file_data.get('D')
    Friday[4] = file_data.get("B")
    Friday[5] = file_data.get('F')
    Friday[6] = file_data.get("C")
    Friday[7] = file_data.get('A')

    final_data["Monday"] = Monday
    final_data["Tuesday"] = Tuesday
    final_data["Wednesday"] = Wednesday
    final_data["Thursday"] = Thursday
    final_data["Friday"] = Friday

    total_sum.append(final_data)

# 최종 json 파일 생성
with open('data.json', 'w', encoding='utf-8') as make_file:
    json.dump(total_sum, make_file, ensure_ascii=False, indent="\t")