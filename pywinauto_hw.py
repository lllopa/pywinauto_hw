import pywinauto


pwa_app = pywinauto.application.Application()
pwa_app.start('notepad.exe')

# Type some keys into the editbox
pwa_app.UntitledNotepad.Edit.TypeKeys("test{SPACE}from{SPACE}pywinauto ^A")

# Get handle to the window
w_handle = pywinauto.findwindows.find_windows(title=u'Untitled - Notepad', class_name='Notepad')[0]
window = pwa_app.window_(handle=w_handle)

window.MenuItem(u'F&ormat->&Font...').Click()

w_format_handle = pywinauto.findwindows.find_windows(title=u'Font', class_name='#32770')[0]
w_format = pwa_app.window_(handle=w_format_handle)

#Setup font, and style
combo = w_format['ComboBox']
combo_props = combo.GetProperties()
font = combo.Select('Comic Sans MS')
combo2 = w_format['ComboBox2']
style = combo2.Select('Bold')
combo3 = w_format['ComboBox3']
combo3.TypeKeys('17')
w_format.OK.click()

#open Print dialog
window.TypeKeys('^P')
w_print = pwa_app.Print
w_print.Wait('ready')
button = w_print.GroupBox
button.SetFocus()
syslistview = w_print.ListView

#Select Print to PDF printer
listview_item = syslistview.GetItem(u'Microsoft Print to PDF')
listview_item.Click()

#Open preference
preference = w_print[u'P&references']
preference.Click()
w_preference = pwa_app.Dialog
w_preference.Wait('ready')
#Set landscape
w_preference.ComboBox.Select(u'Landscape')
#w_preference.ComboBox.Select(u'Portrait')

#Open advanced settings
w_preference[u'Ad&vanced...'].Click()
w_adv = pwa_app[u'Microsoft Print To PDF Advanced Options']
w_adv.Wait('ready')

#Select A3 for seize and save
w_adv.ComboBox.Select(u'A3')
w_adv.OK.Click()


#Save preference
w_preference.OK.Click()

#Send to print
print = w_print[u'&Print']
print.Click()



#Set file with path
w_save = pwa_app[u'Save Print Output As']
w_save.Wait('ready')
w_save.TypeKeys(r'C:\temp\test_print')
w_save.TypeKeys(u'{ENTER}')

#Try case if file exists
try:

    w_save = pwa_app[u'Confirm Save As']
    w_save.Wait('ready')
    w_save[u'&Yes'].Click()

except:
    pass

edit = window['Edit']

#Clear the text to close the Notepad
edit.TypeKeys(u'{DELETE}')
window.Close()


