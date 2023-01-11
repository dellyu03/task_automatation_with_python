# 한글 자동화에 필요한 win32com 모듈을 불러오고 win32라고 지정합니다
import win32com.client as win32

#한글 오브젝트를 생성하는 코드
hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")

# 프로그램을 실행하면 한글 창이 눈에 보이게 하는 코드 False로 바꾸면 한글창이 뜨지 않은 상태로 작업 수행 
hwp.XHwpWindows.Item(0).Visible = True

#생성할 한글 파일이 저장될 경로 지정
saveDirectory = "C:\Users\ms\Desktop\hwp_automatation_docs"
hwp.SaveAs(saveDirectory)

hwp.Save()
