import pyautogui
import pygetwindow as gw
import time
import pandas as pd
import pyperclip

excel = r'C:\Temp\Invoices.xlsx'
invoices_img = r'C:\Temp\Images\Invoices.PNG'
new_record_img = r'C:\Temp\Images\New.PNG'
save_record_img = r'C:\Temp\Images\Save.PNG'

# Luetaan Excelin data dataframeen
data = pd.read_excel(excel)

# Avataan ohjelma
pyautogui.hotkey('winleft', 'r')
pyautogui.write(r'C:\Program Files (x86)\Contoso, Inc\Contoso Invoicing\LegacyInvoicingApp.exe')
pyautogui.press('enter')

program_title = 'Contoso Invoicing'

while True:
    windows = gw.getWindowsWithTitle(program_title)
    if windows:
        print(f'Found: {windows[0]}')
        break
    else:
        print('Not found yet')
        time.sleep(1)

'''
left = windows[0].left
top = windows[0].top
width = windows[0].width
height = windows[0].height
'''

def click_tree_item(image_path):
    try:
        item_location = pyautogui.locateOnScreen(image_path)
        if item_location:
            print(item_location)
            pyautogui.click(item_location)
            print('Invoices clicked.')
        else:
            print('Cannot find Accounting from navigation.')
    except Exception as e:
        print(f'Error: {e}')
        print(item_location)

def click_new_record(image_path):
    try:
        item_location = pyautogui.locateOnScreen(image_path)
        if item_location:
            print(item_location)
            pyautogui.click(item_location)
        else:
            print(f'Cannot find New Record button.')
    except Exception as e:
        print(f'Error: {e}')

def click_save_record(image_path):
    try:
        item_location = pyautogui.locateOnScreen(image_path)
        if item_location:
            print(item_location)
            pyautogui.click(item_location)
        else:
            print(f'Cannot find New Record button.')
    except Exception as e:
        print(f'Error: {e}')

#Annetaan ohjelmalle aikaa käynnistyä - time.sleep ei kuitenkaan ole paras käytäntö
time.sleep(2)
click_tree_item(invoices_img)

for i in range(len(data)):
    click_new_record(new_record_img)
    pyautogui.write(data.loc[i, 'Date'])
    pyautogui.hotkey('tab')
    
    pyautogui.write(data.loc[i, 'Account'])
    pyautogui.hotkey('tab')
    
    pyperclip.copy(data.loc[i, 'Contact'])
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')
    
    pyautogui.write(str(data.loc[i, 'Amount']))
    
    #time.sleep(1)
    click_save_record(save_record_img)
    time.sleep(1)


# Suljetaan ohjelma 4 s kuluttua
time.sleep(4)
pyautogui.hotkey('Alt', 'F4')
    
