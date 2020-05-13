from pywinauto.application import Application

app = Application().Start(cmd_line=u'"C:\\Program Files (x86)\\Notepad++\\notepad++.exe" ')
notepad = app[u'Notepad++']
notepad.Wait('ready')
menu_item = notepad.MenuItem(u'&?->\u95dc\u65bc Notepad++...\tF1')
menu_item.Click()

app.notepad.Edit.TypeKeys("1\thello,hello\rwahaha")
# app.Kill_()cls
