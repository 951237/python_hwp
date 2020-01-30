# -*- coding:utf-8 -*-

import win32com.client as win32  # 윈도우에서만 설치가능
import pandas as pd
import shutil # 파일 복사용 모듈
from datetime import datetime as dt # 작업시간을 측정하기 위함.
# import win32gui #한/글 창을 백그라운드로 숨기기 위한 모듈

PATH_EXCEL = 'c:/Users/User/Desktop/python_hwp/화성오산.xls'
PATH_HWP = 'c:/Users/User/Desktop/python_hwp/award.hwp'
PATH_HWP_RESULT = 'c:/Users/User/Desktop/python_hwp/award_result.hwp'

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
# hwnd = win32gui.FindWindows(None, '빈 문서 1 - 한글') # 한/글 창의 윈도우 핸들값을 알아내서 
# win32gui.ShowWindow(hwnd, 0) # 한/글 창을 백그라운드로 숨김

# hwp.RegisterModule("FilePathCheckDll", "FilePathCheckerModule") # 보안모듈 적용(파일 열고닫을 때 팝업이 안나타남)

shutil.copyfile(PATH_HWP,       # 원본은 그대로 두고
                PATH_HWP_RESULT # 복사한 파일을 수정하려고 함.
                )

print('데이터 파일 로딩중...')
excel = pd.read_excel(PATH_EXCEL)
# excel
hwp.Open(PATH_HWP_RESULT)

start_time = dt.now()   # 작업시간을 측정하기 위해 현재 시각을 start_time 변수에 저장

field_list = [i for i in hwp.GetFieldList().split("\x02")]

hwp.Run('SelectAll') #Ctrl-A 전체선택
hwp.Run('Copy') #Ctrl-C 복사
hwp.MovePos(3) # 문서 끝으로 이동

print('페이지 복사를 시작합니다.')

for i in range(len(excel) - 1): # 엑셀파일 행 갯수-1 만큼 한/글 페이지를 복사(기존에 1페이지가 있으니까)
    hwp.Run('paste') # Ctrl -V  붙여넣기
    hwp.MovePos(3) # 문서 끝으로 이동

print(f'{len(excel)}페이지 복사를 완료하였습니다.')

for page in range(len(excel)): # 한/글 모든 페이지를 전부 순회하면서,
    for field in field_list: #모든 누름틀에 각각,
        hwp.MoveToField(f'{field}{{{{page}}}}') # 커서를 해당 누름틀로 이동(작성과정을 지켜보기 위함. 없어도 무관)
        hwp.PutFieldText(f'{field}{{{{page}}}}', # f'{{{{page}}}}'는 '{{1}}'로 입력된다. {를 출력하려면 {{를 입력.
                        excel[field].iloc[page]) #hwp.PutFieldText('index{{1}}')식으로 실행될 것.
    print(f'{page + 1} : {excel.성명[page]}') # 현재 입력이 진행되고 있는 한/글문서 페이지 번호를 콘솔창에 출력

end_time = dt.now()     # 작업종료 시각
소요시간 = end_time - start_time    # 전체작업시간을 기록

print(f'작업을 완료하였습니다. 약 {소요시간.seconds}초 소요되었습니다.')    # 작업완료된 후 출력. 끝.
