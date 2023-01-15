# 한글 자동화에 필요한 win32com 모듈을 불러오고 win32라고 지정합니다
import win32com.client as win32
import openpyxl as os
from openpyxl import load_workbook

wb = os.load_workbook(r"C:\Users\ms\Desktop\hwp_automatation_docs\계약서 내용표 함수 지운거.xlsx") 
ws = wb.active

allWorks = []

#엑셀 파일에 있는 공사의 개수
workNum = int(input("공사의 개수를 입력하세요 : "))

#몇번째 공사인지
which_work = int(input("몇 번째 공사의 정보를 가저올까요 : (숫자로 입력) : "))


#공사 개수만큼 엑셀 파일의 데이터를 받아오는 코드
all_values = []
for row in ws.rows:
    row_value = []
    for cell in row:
        row_value.append(cell.value)
    all_values.append(row_value)

for i in range(workNum):
    if i == 0:
        allWorks.append(all_values[i+1])
    else:
        allWorks.append(all_values[i])

#한글 오브젝트를 생성하는 코드
hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")

#보안 팝업 제거 명령어
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

# 프로그램을 실행하면 한글 창이 눈에 보이게 하는 코드 False로 바꾸면 한글창이 뜨지 않은 상태로 작업 수행 
hwp.XHwpWindows.Item(0).Visible = True

#특정 한글 경로에 있는 한글 파일 읽기 (원본 양식 파일이 있는 경로)
openDirectory = r"C:\Users\ms\Desktop\hwp_automatation_docs\계약서 서식.hwp"
hwp.Open(openDirectory ,"HWP", "forceopen:true") 


field_list = [i for i in hwp.GetFieldList().split("\x02")]  # 한/글 안의 누름틀 목록 불러오기

print(field_list)


#공사 정보를 받아옴 
workName = allWorks[which_work][0] #공사명
capacity = allWorks[which_work][1]#설치용량
place = allWorks[which_work][2] #설치 장소
period = allWorks[which_work][3] #공사기간
dpayment = allWorks[which_work][4] #계약금
ipayment = allWorks[which_work][5] #중도금
balance = allWorks[which_work][6] #잔금
krcontract_amount = allWorks[which_work][7]#계약금액(한글)
contract_amount = allWorks[which_work][8]#계약금액(숫자)
krsvalue = allWorks[which_work][9]#공급가액(한글)
svalue = allWorks[which_work][10]#공급가액(숫자)
krvat = allWorks[which_work][11]#부가세(한글)
vat = allWorks[which_work][12]#부가세(숫자)
krcvat = allWorks[which_work][13]#계약금_부가세(한글)
cvat = allWorks[which_work][14]#계약금_부가세(숫자)
krivat = allWorks[which_work][15]#중도금_부가세(한글)
ivat = allWorks[which_work][16]#중도금_부가세(숫자)
krbvat = allWorks[which_work][17]#잔금_부가세(한글)
bvat = allWorks[which_work][18]#잔금_부가세(숫자)
RegistrationNum = allWorks[which_work][19]#사업자등록번호
head = allWorks[which_work][20]#대표자
phone_num = allWorks[which_work][21]#연락처
date = allWorks[which_work][22]#계약일

#공사 정보를 한글 문서에 입력함 
#(바뀌길 원하는 부분을 누름틀 설정후 누름틀 필드값, 변수명으로 입력)
hwp.PutFieldText("공사명", workName)
hwp.PutFieldText("설치용량", capacity)
hwp.PutFieldText("설치장소", place)
hwp.PutFieldText("공사기간", period)
hwp.PutFieldText("계약금", dpayment)
hwp.PutFieldText("중도금", ipayment)
hwp.PutFieldText("잔금", balance)
hwp.PutFieldText("계약금액(한글)", krcontract_amount)
hwp.PutFieldText("계약금액(숫자)", contract_amount)
hwp.PutFieldText("공급가액(한글)", krsvalue)
hwp.PutFieldText("공급가액(숫자)", svalue)
hwp.PutFieldText("부가세(한글)", krvat)
hwp.PutFieldText("부가세(숫자)", vat)
hwp.PutFieldText("계약금_부가세(한글)", krcvat)
hwp.PutFieldText("계약금_부가세(숫자)", cvat)
hwp.PutFieldText("중도금_부가세(한글)", krivat)
hwp.PutFieldText("중도금_부가세(숫자)", ivat)
hwp.PutFieldText("잔금_부가세(한글)", krbvat)
hwp.PutFieldText("잔금_부가세(숫자)", bvat)
hwp.PutFieldText("사업자등록번호", RegistrationNum)
hwp.PutFieldText("대표자", head)
hwp.PutFieldText("연락처", phone_num)
hwp.PutFieldText("계약일", date)

# 입력된 거래 정보가 무엇인지 출력해주는 코드
print("입력된 거래 정보 : " + str(allWorks[which_work]))
