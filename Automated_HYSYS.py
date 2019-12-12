import subprocess
import time
import pyautogui as pg
import win32gui, win32con
import re, traceback
import custom_functions_HYSYS #customized function developed specifically for this project
from tkinter import *
from tkinter import ttk
from tkinter import messagebox

path_HYSYS = r'C:\Users\szjt\Desktop\UBFPSO working\Automated_HYSYS\case 6a NGA2 FPSO_rev D01_Revise GL tapping point_RAW.hsc'
path_excel = r'C:\Users\szjt\Desktop\UBFPSO working\Automated_HYSYS\Simulation_Cases.xlsm'

#subprocess.Popen(path_excel, shell = True)
#time.sleep(1)
#subprocess.Popen(path_HYSYS, shell = True)
#time.sleep(1)

#print("Open HYSYS and Excel Successfully")
run_case()

GUI = Tk()
GUI.geometry('600x300+50+50')
GUI.title('Checking whether HYSYS is running OK')

mainmenu = Menu(GUI)
GUI.config(menu = mainmenu)

filemenu = Menu(mainmenu, tearoff = 0)
filemenu.add_command(label='Copy HYSYS Value to Excel', command = copy_hysys)
filemenu.add_command(label='Exit', command = GUI.quit)
mainmenu.add_cascade(label='File', menu=filemenu)

text_show = StringVar()
text_show.set('Check HYSYS and ensure that there is no error')

L1 = ttk.Label(GUI, textvariable = text_show , font=('Arial', 18))
L1.place(x=20, y=50)

FB1 = Frame(GUI)
FB1.place(x=100, y = 100)

B1 = ttk.Button(FB1, text ='Checked and found OK. Proceed to next case', width = 50, command = copy_hysys)
B1.grid(row=0, column =1, ipadx = 20, ipady = 10)

FB2 = Frame(GUI)
FB2.place(x=100, y = 160)

B2 = ttk.Button(FB2, text ='This is the final case. Close program', width = 50, command = is_OK)
B2.grid(row=0, column =1, ipadx = 20, ipady = 10)

GUI.mainloop()
GUI.BringToTop()



