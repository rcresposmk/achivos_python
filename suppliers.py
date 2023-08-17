import pyautogui as robot
import time
import subprocess
import psutil
import pandas as pd

robot.hotkey('winleft', 'd')
# Open GIGO from network
subprocess.run([r"\\m-sys003\GIGOBIN\GIGO.exe"])
time.sleep(10)

robot.hotkey('winleft', 'up')
for i in range(3):
    robot.press('down')
robot.press('enter')

time.sleep(3)

robot.write('m03', interval=0.25)
robot.write('mis4', interval=0.25)
robot.press('enter') 

time.sleep(3)
robot.press('tab')
time.sleep(1)
robot.write('8', interval=0.25)
robot.press('enter')

time.sleep(2)
robot.press('tab')
time.sleep(3)
robot.press('enter')
time.sleep(3)

# Select supplier 030
for i in range(3):
    robot.press('down')
    time.sleep(1)

robot.press('enter')

# press space to change Y fields
robot.press('right')
robot.press('space')
robot.press('down')
robot.press('space')
robot.press('down')
robot.press('space')
robot.press('right')
robot.press('right')
robot.press('space')
robot.press('up')
robot.press('space')
robot.press('up')
robot.press('space')
robot.press('left')
robot.press('left')
robot.press('left')
robot.press('enter')
time.sleep(3)
robot.press('enter')
time.sleep(10)
robot.hotkey('winleft', 'up')

robot.press('altleft')
robot.press('down')
robot.press('enter')

robot.press('tab')
robot.press('tab')

robot.write(r'\\10.1.2.224\d$\CSV\M\DN\PM\SUPPLIERS')
robot.press('enter')
robot.press('enter')

time.sleep(15)

# kill process excel and GIGO
for proc in psutil.process_iter():
    if proc.name() == "EXCEL.EXE":
        proc.kill()
    
    if proc.name() == "ESTET.EXE":
        proc.kill()

time.sleep(5)

# convert .xls to .cvs
read_file = pd.read_excel(r'\\10.1.2.224\d$\CSV\M\DN\PM\SUPPLIERS.xls', 'ADHOC')
read_file.to_csv (r'\\10.1.2.224\d$\CSV\M\DN\PM\SUPPLIERS.csv', index=False, sep=';')