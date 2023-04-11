# File:     TkinterGUI_2023-04-05
# Version:  0.0.01
# Author:   Susan Haynes
# Comments/Notes: 
#   (0,0) coordinates are the top left corner of the screen for 1920x1080
#   (0,0) coordinates are the bottom right corner of the screen for 1919x1079
# Online References: 
#   https://pypi.org/project/PyAutoGUI/
#   https://pyautoGUI.readthedocs.io/en/latest/mouse.html
# Revision History: N/A 
# To check tkinter is installed, use this in command promt.
# python -m tkinter 
#################################################      CLASSES OF LIBRARIES TO USE      ################################################  
from functools import partial
from mimetypes import init                           # for allowing 015 040 buttons to equal specific values when clicked.
from openpyxl import *                                  # Write to excel
import xlsxwriter                                       # Excel Writer library 
import tkinter as tk                                    # Tkinter's Tk class
import tkinter.ttk as ttk                               # Tkinter's Tkk class
import datetime as dt                                   # Date library
import pandas as pd
import subprocess                                       # Needed to open an executable
import time                                             # Needed to call time to count/pause
import psutil,os                                        # Needed for closing an executable
from PIL import ImageTk, Image                          # Displaying LAL background photo
from tkinter import messagebox                          # Exit standard message box

##########################################################################################################################################
#################################################         1st    GUI SCREEN              #################################################
##########################################################################################################################################
#################################################      INITIALIZING STANDARD DISPLAY     ################################################# 
GUI = tk.Tk()
GUI.title("LAL Measurement")
GUI.geometry("1500x1000")                               # Set the geometry of Tkinter frame
GUI.configure(bg = 'white')                             # Set background color
GUI.option_add("*Font", "Helvetica 12 bold")            # set the font and size for entire GUI
GUI.option_add("*fg", "black")                          # set the text color, hex works too "#FFFFFF"
GUI.option_add("*bg", "white")                          # set the background to white

#################################################            BUTTON PRESS STYLE           ################################################ 
style = ttk.Style(); 
style.theme_use('default');     # alt, default, clam and classic
style.map('T.TButton',background=[('active', 'pressed', 'white'),('!active','white'), ('active','!pressed','grey')]) # active, not active, not pressed
style.map('T.TButton',relief    =[('pressed','sunken'),('!pressed','raised')]) # pressed, not pressed
style.configure("T.Button", font= ('Helvetica', '12', 'bold'))
style.map('P.TButton',background=[('active', 'pressed', '#FF69B4'),('!active','white'), ('active','!pressed','grey')]) # Press me Button always hot pink when pressed
style.map('P.TButton',relief    =[('pressed','sunken'),('!pressed','raised')]) # pressed, not pressed
style.configure("P.Button", font= ('Helvetica', '12', 'bold'))

# Python is serial, so each widget will output in the order listed below;
#################################################           LAL BACKGROUND IMAGE          ################################################  
def resize_image(event):
    new_width = event.width
    new_height = event.height
    background_image = copy_of_image.resize((new_width, new_height))
    bkgnd_img = ImageTk.PhotoImage(background_image)
    lbl_photo.config(image = bkgnd_img)
    lbl_photo.background_image = bkgnd_img #avoid garbage collection

background_image = Image.open(r"\\RXS-FS-02\userdocs\shaynes\My Documents\R&D - Software\Python\TkinterGUI_2023-03-24\LAL.png")
copy_of_image = background_image.copy()
bkgnd_img = ImageTk.PhotoImage(background_image)

lbl_photo = ttk.Label(GUI, image = bkgnd_img)
lbl_photo.bind('<Configure>', resize_image)
lbl_photo.pack(fill=tk.BOTH, expand = True)

date = dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

#################################################               EXCEL FILE               ################################################  
try:
    workbook = load_workbook(f"{filename},{date}.xlsx", index=False)                     # does this even do anything?
    sheet = workbook.active
except:
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Credentials"         # Column 1
    sheet["B1"] = "subWO"               # Column 2
    sheet["C1"] = "Sample"              # Column 3
    sheet["D1"] = "Measurement"         # Column 4
    sheet["E1"] = "Date&Time"           # Column 5
    sheet["F1"] = "OCT Eq."             # Column 6
    sheet["G1"] = "Prod/RD"             # Column 7

new_line = sheet.max_row + 1

################################################                 MAIN BODY                ################################################   
# Display the command label before the entry box to indicate what information the Opterator is to type
lbl_cmd_date = tk.Label(GUI, text="Today's Date is:",            bg= "white", width= 12).place(x=50,y=25)  
lbl_cmd_fold = tk.Label(GUI, text="Folder Name:",                bg= "white", width= 10).place(x=50,y=75) 
lbl_cmd_file = tk.Label(GUI, text="File Name:",                  bg= "white", width=  8).place(x=50,y=125) 
lbl_cmd_cred = tk.Label(GUI, text="Enter Operator Credentials:", bg= "white", width= 20).place(x=50,y=175) 
lbl_cmd_WO   = tk.Label(GUI, text="Enter Work Order Number:",    bg= "white", width= 20).place(x=50,y=225)   
lbl_cmd_samp = tk.Label(GUI, text="Enter Sample Size:",          bg= "white", width= 14).place(x=50,y=275)  
lbl_cmd_meas = tk.Label(GUI, text="Select Measurement Size:",    bg= "white", width= 19).place(x=50,y=325) 
lbl_cmd_oct  = tk.Label(GUI, text="Select OCT Equipment:",       bg= "white", width= 17).place(x=50,y=375)  
lbl_cmd_prd  = tk.Label(GUI, text="Production:",                 bg= "white", width=  9).place(x=50,y=425)  
lbl_cmd_rnd  = tk.Label(GUI, text="R&D:",                        bg= "white", width=  4).place(x=350,y=425)  

# Entry boxes to take information from operator
entry_cred = tk.Entry(GUI, bg= "white", width= 10) 
entry_cred.focus_set()                              # Places cursor in the first entry box.
entry_cred.place(x=300,y=175) 
entry_WO   = tk.Entry(GUI, bg= "white", width= 10) 
entry_WO.place(x=300,y=225) 
entry_samp = tk.Entry(GUI, bg= "white", width= 10) 
entry_samp.place(x=300,y=275)  

# Display the label of what user input will be displayed.
lbl_disp_cred = tk.Label(GUI, text="Credentials:",       bg= "white", width= 9) .place(x=50, y=600) 
lbl_disp_WO   = tk.Label(GUI, text="Work Order Number:", bg= "white", width= 16).place(x=50, y=640) 
lbl_disp_samp = tk.Label(GUI, text="Sample Size:",       bg= "white", width= 10).place(x=50, y=680) 
lbl_disp_meas = tk.Label(GUI, text="Measurement Size:",  bg= "white", width= 14).place(x=50, y=720) 
lbl_disp_oct  = tk.Label(GUI, text="OCT Equipment:",     bg= "white", width= 12).place(x=50, y=760) 
lbl_disp_pr   = tk.Label(GUI, text="Production/R&D:",    bg= "white", width= 13).place(x=50, y=800) 

# Display the user inputs as outputs 
lbl_out_date = tk.Label(GUI, text=f"{dt.datetime.now():%b %d, %Y}", bg= "white", width= 9).place(x=300, y=25) 
lbl_out_cred = tk.Label(GUI, text= "", bg= "white", width= 3)
lbl_out_cred.place(x=300, y=600) 
lbl_out_WO   = tk.Label(GUI, text= "", bg= "white", width= 6)
lbl_out_WO.place(x=300, y=640) 
lbl_out_samp = tk.Label(GUI, text= "", bg= "white", width= 3)
lbl_out_samp.place(x=300, y=680) 
 
# Display user inputs as outputs
def fun_cred():
    global entry
    cred = entry_cred.get()[:3]                          # entry_cred is the variable we are passing. Limit 3 characters
    lbl_out_cred.configure(text = cred)                  # Display cred entry from user on GUI
    sheet.cell(column=1, row=new_line, value = entry_cred.get()[:3])
    print(entry_cred.get()[:3])                          # Print can be removed after developed.

def fun_WO():
    global entry
    WO = entry_WO.get()[:6]                              # entry_WO is the variable we are passing. Limit 10 characters
    lbl_out_WO.configure(text = WO)                      # Display WO entry from user on GUI
    sheet.cell(column=2, row=new_line, value = entry_WO.get()[:6])
    print(entry_WO.get()[:6]) 

def fun_samp():
    global entry
    samp = entry_samp.get()[:3]                          # entry_samp is the variable we are passing. Limit 3 characters
    lbl_out_samp.configure(text = samp)                  # Display sample entry from user on GUI
    sheet.cell(column=3, row=new_line, value = entry_samp.get()[:3])
    print(entry_samp.get()[:3]) 

def fun_meas(entry_meas, excel_meas):
    if entry_meas== '-B':
        btn_015 = tk.Entry(GUI, width= 10)
        btn_015.insert(0,entry_meas) 
        btn_015.place(x=300, y=720)
        sheet.cell(column=4, row=new_line).value = excel_meas
    elif entry_meas== '-A':
        btn_040 = tk.Entry(GUI, width= 10)
        btn_040.insert(0,entry_meas) 
        btn_040.place(x=300, y=720)
        sheet.cell(column=4, row=new_line).value = excel_meas
    elif entry_meas== "":
        btn_100 = tk.Entry(GUI, width= 10)
        btn_100.insert(0,entry_meas) 
        btn_100.place(x=300, y=720)
        sheet.cell(column=4, row=new_line).value = excel_meas       
    print("entry_meas is: ", entry_meas)
    print("excel_meas is: ", excel_meas)
    fun_cred()
    fun_WO()
    fun_samp()

def fun_oct(entry_oct):
    if entry_oct== 'OCT 1':
        btn_oct1 = tk.Entry(GUI, width= 10)
        btn_oct1.insert(0,entry_oct) 
        btn_oct1.place(x=300, y=760)
    elif entry_oct== 'OCT 2':
        btn_oct2 = tk.Entry(GUI, width= 10)
        btn_oct2.insert(0,entry_oct) 
        btn_oct2.place(x=300, y=760)
    sheet.cell(column=6, row=new_line).value = entry_oct
    print("entry_oct is: ", entry_oct)

def fun_prd(entry_pr, excel_pr):
    if entry_pr== '02':    # Haptics
        btn_pr02 = tk.Entry(GUI, width= 10)
        btn_pr02.insert(0,entry_pr) 
        btn_pr02.place(x=300, y=800)
    elif entry_pr== '06':   # R&D
        btn_rd06 = tk.Entry(GUI, width= 10)
        btn_rd06.insert(0,entry_pr) 
        btn_rd06.place(x=300, y=800)
    elif entry_pr== '07':   # Standard Production
        btn_pr07 = tk.Entry(GUI, width= 10)
        btn_pr07.insert(0,entry_pr) 
        btn_pr07.place(x=300, y=800)
    elif entry_pr== '08':   # Next Gen LAL+ R&D
        btn_rd08 = tk.Entry(GUI, width= 10)
        btn_rd08.insert(0,entry_pr) 
        btn_rd08.place(x=300, y=800)
    sheet.cell(column=7, row=new_line).value = excel_pr
    print("entry_pr is: ", entry_pr)
    print("excel_pr is: ", excel_pr)

def fun_save(): #entry_meas, entry_pr
    sheet.cell(column=5, row=new_line).value = date
    workbook.save(filename= 'L'+entry_WO.get()[:6]+'.xlsx')
    print('L'+entry_WO.get()[:6]+'.xlsx')   #     print('L'+entry_pr+entry_WO.get()[:6]+entry_meas+'.xlsx')
    lbl_out_nam_fil = tk.Label(GUI, text= 'L'+entry_WO.get()[:6]+'.xlsx', bg= "white")
    lbl_out_nam_fil.place(x=300, y=125)

# workbook.save(filename= xlsxwriter.Workbook(dt.datetime.now().strftime('%Y-%m-%d')+'.xlsx'))

def open_lum():
    # Open the calculator, and pause for 2 seconds before executing, this gives the calculator time to open.
    subprocess.Popen('C:\\Windows\\System32\\calc.exe')                     # Open windows calculator
    time.sleep(5)                                                           # wait 5 seconds
    os.system("TaskKill /F /IM CalculatorApp.exe")                          # Close windows calculator

def exit_app(): 
    msg_box = tk.messagebox.askquestion('Exit', 'Are you sure you want to exit the application?', icon='warning') 
    if msg_box == 'yes': 
        GUI.destroy() 
    else: 
        tk.messagebox.showinfo('Exit', 'Thanks for staying, please continue.') 
 
btn_pres_cnt = 1                                                         # setting count to 0 to be able to call it a global variable within the function
def pink(event):                     
    global btn_pres_cnt                                                  # initializing btn_pres_cnt as a global varaible so that it adds through every iteration
    if(btn_pres_cnt==5 or btn_pres_cnt==10 or btn_pres_cnt==15 or btn_pres_cnt==20 or btn_pres_cnt==25): # button turns pink when btn_pres_cnt=100, and =200 and = 300.
        style.map('T.TButton',background=[('active', 'pressed', '#FF69B4'),('!active','white'), ('active','!pressed','grey')])    # only the button being pressed turns hot pink
        style.configure("T.Button", font= ('Helvetica', '12', 'bold'))
    else:   # else is the normal style
        style.map('T.TButton',background=[('active', 'pressed', 'white'),('!active','white'), ('active','!pressed','grey')])
        style.configure("T.Button", font= ('Helvetica', '12', 'bold'))
    print("btn_pres_cnt = ", btn_pres_cnt)                               # This is always executed at the end of the if else
    btn_pres_cnt +=1                                                     # This is always executed at the end of the if else

#################################################        BUTTONS TO BE CLICKED         ################################################   
btn_015  = ttk.Button(GUI, text="Posterior", style= 'T.TButton', command=partial(fun_meas, "-B", "Posterior"))       # Post - 015 is the variable we are passing to excel,-B for excel file name
btn_015.bind("<Button-1>", pink)
btn_015.place(x=300,y=325)  

btn_040  = ttk.Button(GUI, text="Anterior",  style= 'T.TButton', command=partial(fun_meas, "-A", "Anterior"))       # Ant - 040 is the variable we are passing to excel, -A for excel file name
btn_040.bind("<Button-1>", pink)
btn_040.place(x=400,y=325) 

btn_100  = ttk.Button(GUI, text="Full Lens", style= 'T.TButton', command=partial(fun_meas, "", "Full Lens"))         # Full - 100 is the variable we are passing to excel, blank for the excel file name
btn_100.bind("<Button-1>", pink)
btn_100.place(x=500,y=325) 

btn_oct1  = ttk.Button(GUI, text= "OCT 1",   style= 'T.TButton', command=partial(fun_oct, "OCT 1"))      # OCT 1 is the variable we are passing to excel
btn_oct1.bind("<Button-1>", pink)
btn_oct1.place(x=300,y=375) 

btn_oct2  = ttk.Button(GUI, text= "OCT 2",   style= 'T.TButton', command=partial(fun_oct, "OCT 2"))      # OCT 2 is the variable we are passing to excel
btn_oct2.bind("<Button-1>", pink)
btn_oct2.place(x=400,y=375) 

btn_pr02  = ttk.Checkbutton(GUI, text= "Haptics 02", onvalue= 1, offvalue= 0, style= 'T.TButton', command=partial(fun_prd, "02", "Haptics Production")) # 02 is the variable we are passing to excel
btn_pr02.bind("<Button-1>", pink)
btn_pr02.place(x=175,y=425) 

btn_pr07  = ttk.Checkbutton(GUI, text= "Standard 07", onvalue= 1, offvalue= 0, style= 'T.TButton', command=partial(fun_prd, "07", "Standard Production")) # 07 is the variable we are passing to excel
btn_pr07.bind("<Button-1>", pink)
btn_pr07.place(x=250,y=425) 

btn_rd06  = ttk.Checkbutton(GUI, text= "LAL 06", onvalue= 1, offvalue= 0, style= 'T.TButton', command=partial(fun_prd, "06", "R&D LAL")) # 06 is the variable we are passing to excel
btn_rd06.bind("<Button-1>", pink)
btn_rd06.place(x=400,y=425) 

btn_rd08  = ttk.Checkbutton(GUI, text= "LAL+ 08", onvalue= 1, offvalue= 0, style= 'T.TButton', command=partial(fun_prd, "08", "R&D LAL+")) # 08 is the variable we are passing to excel
btn_rd08.bind("<Button-1>", pink)
btn_rd08.place(x=475,y=425) 

btn_sav   = ttk.Button(GUI, text= "Save",    style= 'T.TButton', command=partial(fun_save))
btn_sav.bind("<Button-1>", pink)
btn_sav.place(x=1150, y=930)

btn_lum = ttk.Button(GUI,text="Open Lumedica",style='T.TButton', command=open_lum)                              # Currently opens calculator, eventually will open lumedica.exe
btn_lum.bind("<Button-1>", pink)
btn_lum.place(x=1260,y=930)

btn_exit = ttk.Button(GUI, text= "Exit",    style= 'T.TButton', command=exit_app)
btn_exit.bind("<Button-1>", pink)
btn_exit.place(x=1400,y=930) 

# Must be at the end of the program in order for the application to run b/c windows is constantly updating
GUI.mainloop()


