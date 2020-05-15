from pywinauto.application import Application

from pywinauto import mouse

# app = Application(backend="uia").start(cmd_line=u'"C:\\Program Files (x86)\\Microsoft\\Skype for Desktop\\Skype.exe" ')

# skp = app.window(title_re='Skype.*') # actual search is NOT happening here
#skp.wait('ready', timeout=20) # wait up to 20 sec. (actual search is done, return early if found)

app = Application(backend= "uia")
app.start(cmd_line=u'"C:\\Program Files (x86)\\Microsoft\\Skype for Desktop\\Skype.exe" ')
skp = app.window(title_re='Skype.*') # actual search is NOT happening here
actionable_skp = skp.wait('visible') 
# app.SkypeUsername.draw_outline()
# dlg =app.SkypeUsername
# contacts =dlg.Listbox  # My contacts seem to be under the header "Listbox" when I use point_control identifiers, so I assigned it to contacts for easy reference.

# search = skp["搜尋使用者、群組及訊息"]
# skp.Properties.button.click()
# skp.print_control_identifiers()
# search.click() # actual search of the sub-element is done here (inside of call to .click())
# app = Application(backend="uia").start(cmd_line=u'"C:\\Program Files (x86)\\Microsoft\\Skype for Desktop\\Skype.exe" ')
# app = Application().start(cmd_line=u'"C:\\Program Files (x86)\\Microsoft\\Skype for Desktop\\Skype.exe" ')
# skype_dig = app.window(title='Skype')
# skype_dig.Properties.print_control_identifiers()
# skype_dig.print_control_identifiers()
# skype_dig["[42.66432.4.-53]"].click()

# app.Kill_()