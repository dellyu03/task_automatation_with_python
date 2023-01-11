# 한글 자동화에 필요한 win32com 모듈을 불러오고 win32라고 지정합니다
import win32com.client as win32

#win32
hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
hwp.XHwpWindows.Item(0).Visible = True