from pywinauto.application import Application
import time

app = Application().Start('Notepad.exe')
notepad = app[u'Notepad']
notepad.edit.TypeKeys('12234')
#notepad.MenuSelect("檔案(&F)->另存為...")
time.sleep(1)
#notepad.app['Dialog'].edit.SetText("123.txt")
time.sleep(0.5)
app[' 未命名 - 記事本 '].menu_select('檔案(&F) -> 結束(&X)')

# 在这时候不清楚“不保存”的按钮名就对app['记事本'] 使用print_control_identifiers()
#app['記事本'].print_control_identifiers()
app['記事本']['不要儲存(&N)'].click()


# notepad.SaveAs.Save.click()

# app.Kill_()cls
