# import win32com.client as win32
# import win32gui
import pandas as pd
import openpyxl
from datetime import datetime as dt

# 1. basepath 설정
PATH_HWP = '/Users/mac/Documents/python_hwp/자동화_과학실사용일지_200131/timetable.hwp'
PATH_EXCEL = '/Users/mac/Documents/python_hwp/자동화_과학실사용일지_200131/science.xlsx'
PATH_EXCEL_RESULT = '/Users/mac/Documents/python_hwp/자동화_과학실사용일지_200131/science_sel.xlsx'
# 2. 한/글 object 생성
# 3. 엑셀파일 불러오기
def sel_sht_name(path, year, month):
    xl = pd.ExcelFile(path)
    sht_names = xl.sheet_names
    sht_names_new = []
    for name in sht_names:
        if int(name.split('.')[0]) > year and int(name.split('.')[1]) < month:
            sht_names_new.append(name)
    return sht_names_new

start_time = dt.now()
sht_name_new = sel_sht_name(PATH_EXCEL, 18, 9)  # 19년 9월 이하 시트 이름만 수집
sht_name_new.sort()

df_res = pd.DataFrame()

for i in range(23):
    df_xl = pd.DataFrame()
    print(f'{ sht_name_new[i] }시트 로딩중...')
    df_xl = pd.read_excel(PATH_EXCEL, sheet_name=sht_name_new[i], usecols={0, 1, 2, 5}) 

    def clean_row(row):
        return str(row).replace('\n', '')

    def write_day(row):
        sht_name_new[i] = sht_name_new[i].replace('-', '~')
        day = sht_name_new[i].split('~')[0].split('.')
        if row == '월':
            return f'{int(day[1])}.{int(day[2])}'
        elif row == '화':
            return f'{int(day[1])}.{int(day[2]) + 1}'
        elif row == '수':
            return f'{int(day[1])}.{int(day[2]) + 2}'
        elif row == '목':
            return f'{int(day[1])}.{int(day[2]) + 3}'
        elif row == '금':
            return f'{int(day[1])}.{int(day[2]) + 4}'

    print('줄바꿈을 삭제합니다.')
    df_xl['요일'] = df_xl['요일'].apply(clean_row) # 요일에 '\n'을 공백 처리

    print('시트의 날짜를 기록합니다.')
    df_xl['수업일'] = df_xl['요일'].apply(write_day) # 시트이름에서 요일을 추출 하여 월, 화, 수, 목, 금에 맞게 넣기

    print(f'{sht_name_new[i]}시트 데이터프레임 합치기\n')
    df_res = pd.concat([df_res, df_xl])
print('데이터 프레임 작성 완료!')

print('엑셀파일로 저장합니다.')
df_res.to_excel(PATH_EXCEL_RESULT, sheet_name=sht_name_new[0])

end_time = dt.now()

소요시간 = end_time - start_time
print(f'작업을 완료합니다. 소요시간은 {소요시간.seconds}입니다.')

#todo 
- 학년반 나타내기
- 날짜데이터 형식을 이용한 덧셈하기



