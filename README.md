# OldAutoHealthChecker

## 사용방법
studentlist.xlsx 파일 생성, 버전에 맞는 selenium을 위한 chromedriver.exe 파일  
파일 안에 끊기지 않게 이름과 생일 입력(셀 서식->텍스트)  
## 요구 사항: 
python 3(환경변수 설정할 것)  
selenium ( https://www.selenium.dev/  )  
openpyxl ( https://openpyxl.readthedocs.io/en/stable/ )  
( cmd에서 pip install selenium 과 pip install openpyxl 각각 입력해서 설치할 수 있음 )  
driver.get("https://eduro.pen.go.kr/stv_cvd_co00_002.do") 에서 링크를 자신 교육청 자가진단 페이지-> 학생정보 입력의 페이지를 입력  
send_keys('여기에 학교이름 입력') 을 자신의 학교 이름으로 수정
time.sleep(?)에 적당한 실수를 넣어서 입력간 딜레이를 조절할 수 있음
<hr/>

## 기능 :  
엑셀 파일의 끝을 자동으로 인지함  
시작할 번호를 설정가능
자가진단에 성공한 이름과 현재까지의 진행정도를 표시함
자가진단 하다 꺼졌을 경우에 끊긴 부분부터 새로 시작함(엑셀 3번째 열에 1이 기록됨)
자가진단을 처음부터 하려면 파일을 3열을 모두 지우거나 CleanExcel 사용

## 윈도우 작업 스케줄러를 이용한 정해진 시간 자동 시작
정해진 시간에 시작하려면 배치 파일을 만들어야 함  
cmd->python->import sys->sys.executable로 파이썬 설치 위치 경로 확인 후 메모장에  

    파이썬 설치경로.exe "실행하려는 .py 파일 위치"
    pause

를 bat 파일로 저장한다(ANSI 인코딩 적용)  
윈도우 작업 스케줄러에 등록한다.  
