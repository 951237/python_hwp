# import win32com.client as win32
# import win32gui
import pandas as pd

# 1. basepath 설정
PATH_HWP = '/Users/mac/Documents/python_hwp/자동화_과학실사용일지_200131/timetable.hwp'
PATH_EXCEL = '/Users/mac/Documents/python_hwp/자동화_과학실사용일지_200131/science.xlsx'
# 2. 한/글 object 생성
# 3. 엑셀파일 불러오기
xl = pd.ExcelFile(PATH_EXCEL)
sht_names = xl.sheet_names
sht_names_new = []
for name in sht_names:
    if int(name.split('.')[0]) < 19 and int(name.split('.')[1]) < 9:
        sht_names_new.append(name)

print(sht_names_new)

# 4. 

