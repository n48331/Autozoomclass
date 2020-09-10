import gspread as gs #pip install gspread
from pprint import *
import time
import pyautogui as pi #pip install pyautogui
import schedule #pip install schedule
import subprocess

gc = gs.service_account(filename='creds.json')# your credensials
sh = gc.open_by_key('14jIEZN8llvlYCDEIoMal6Cj9nilFQMwU4WtryVIqnqY') #your spread sheet key.seen between spread sheet link
sheet = sh.sheet1
res = sheet.get_all_records()
pprint(res)

def zoom1():
    z=subprocess.call(['C:/Users/M Nabeel P/AppData/Roaming/Zoom/bin/Zoom.exe'])#zoom exe file location
    time.sleep(1)
    r=pi.locateCenterOnScreen('2.png')#screenshots(PNGs).if it does not work make your own with snipping tool
    print(r)
    pi.click(r)
    time.sleep(1)
    pi.write(str(sheet.cell(2,1).value),interval=0.05)
    time.sleep(1)
    j=pi.locateCenterOnScreen('join.png')
    pi.click(j)
    time.sleep(2)
    pi.write(str(sheet.cell(2,2).value),interval=0.05)
    time.sleep(1)
    jm=pi.locateCenterOnScreen('join2.png')
    pi.click(jm)
    time.sleep(3)

#2ndclass
def zoom2():
    z=subprocess.call(['C:/Users/M Nabeel P/AppData/Roaming/Zoom/bin/Zoom.exe'])
    time.sleep(1)
    r=pi.locateCenterOnScreen('2.png')
    print(r)
    pi.click(r)
    time.sleep(1)
    pi.write(str(sheet.cell(3,1).value),interval=0.05)
    time.sleep(1)
    j=pi.locateCenterOnScreen('join.png')
    pi.click(j)
    time.sleep(2)
    pi.write(str(sheet.cell(3,2).value),interval=0.05)
    time.sleep(1)
    jm=pi.locateCenterOnScreen('join2.png')
    pi.click(jm)
    time.sleep(3)

#3rdclass
def zoom3():
    z=subprocess.call(['C:/Users/M Nabeel P/AppData/Roaming/Zoom/bin/Zoom.exe'])
    time.sleep(1)
    r=pi.locateCenterOnScreen('2.png')
    print(r)
    pi.click(r)
    time.sleep(1)
    pi.write(str(sheet.cell(4,1).value),interval=0.05)
    time.sleep(1)
    j=pi.locateCenterOnScreen('join.png')
    pi.click(j)
    time.sleep(2)
    pi.write(str(sheet.cell(4,2).value),interval=0.05)
    time.sleep(1)
    jm=pi.locateCenterOnScreen('join2.png')
    pi.click(jm)
    time.sleep(3)

def zoom4():
    z=subprocess.call(['C:/Users/M Nabeel P/AppData/Roaming/Zoom/bin/Zoom.exe'])
    time.sleep(1)
    r=pi.locateCenterOnScreen('2.png')
    print(r)
    pi.click(r)
    time.sleep(1)
    pi.write(str(sheet.cell(5,1).value),interval=0.05)
    time.sleep(1)
    j=pi.locateCenterOnScreen('join.png')
    pi.click(j)
    time.sleep(2)
    pi.write(str(sheet.cell(5,2).value),interval=0.05)
    time.sleep(1)
    jm=pi.locateCenterOnScreen('join2.png')
    pi.click(jm)
    time.sleep(3)

def zoom5():
    z=subprocess.call(['C:/Users/M Nabeel P/AppData/Roaming/Zoom/bin/Zoom.exe'])
    time.sleep(1)
    r=pi.locateCenterOnScreen('2.png')
    print(r)
    pi.click(r)
    time.sleep(1)
    pi.write(str(sheet.cell(6,1).value),interval=0.05)
    time.sleep(1)
    j=pi.locateCenterOnScreen('join.png')
    pi.click(j)
    time.sleep(2)
    pi.write(str(sheet.cell(6,2).value),interval=0.05)
    time.sleep(1)
    jm=pi.locateCenterOnScreen('join2.png')
    pi.click(jm)
    time.sleep(3)
    
print("Started Schedules")
schedule.every().day.at(str(sheet.cell(2,3).value)).do(zoom1)
schedule.every().day.at(str(sheet.cell(3,3).value)).do(zoom2)
schedule.every().day.at(str(sheet.cell(4,3).value)).do(zoom3)
schedule.every().day.at(str(sheet.cell(5,3).value)).do(zoom4)
schedule.every().day.at(str(sheet.cell(6,3).value)).do(zoom5)

while True: 
	schedule.run_pending()
	time.sleep(1) 
schedule.CancelJob()