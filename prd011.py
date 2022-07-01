import pyautogui as robot
import time
import subprocess
import psutil
import pandas as pd

robot.hotkey('winleft', 'd')
# Open GIGO from network
subprocess.run([r"\\m-sys002\GIGOBIN\GIGO.exe"])
time.sleep(1)

robot.hotkey('winleft', 'up')
for i in range(4):
    robot.press('down')
robot.press('enter')

robot.write('xxx', interval=0.25)
robot.write('356z', interval=0.25)
robot.press('enter')

robot.write('prt', interval=0.25)
robot.press('enter')

robot.press('tab')
robot.press('down')
robot.press('down')
robot.press('enter')

robot.write('prd011', interval=0.25)
for i in range(7):
    robot.press('left')
robot.press('enter')

time.sleep(20)

robot.press('altleft')
robot.press('down')
robot.press('enter')
robot.press('tab')
robot.press('tab')

robot.write(r'\\10.1.2.47\d$\CSV\M\DN\PM\PRD011H')
robot.press('enter')
robot.press('enter')

time.sleep(25)

# kill process excel and GIGO
for proc in psutil.process_iter():
    if proc.name() == "EXCEL.EXE":
        proc.kill()
    
    if proc.name() == "ESTET.EXE":
        proc.kill()

time.sleep(5)

# convert .xls to .cvs
read_file = pd.read_excel (r'\\10.1.2.47\d$\CSV\M\DN\PM\PRD011H.xls', 'PRD011')
read_file.to_csv (r'\\10.1.2.47\d$\CSV\M\DN\PM\PRD011H.csv', index=False, sep=';')