import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from PIL import ImageTk, Image
import openpyxl
from openpyxl import *
import xlrd
import xlwt
from tkinter.ttk import *


def leave():
    global leavewindow
    leavewindow = tkinter.Tk()

    canvas = tkinter.Canvas(leavewindow, height=700, width=900) 

    frame = tkinter.Frame(leavewindow, bg='#5F9EA0')
    frame.place(relwidth=1, relheight=1)

    boxlabel = tkinter.Label(frame, text='Main Profile', bg='#5F9EA0', fg='#FFFFFF')
    boxlabel.place(relx=0.05, rely=0.025)

    mainbox = tkinter.Canvas(frame, height=100, width=800, bd=2, bg='#5F9EA0')# the box of main profile
    mainbox.place(relx=0.05, rely=0.05)

    global identry,empnamentry,gradentry,designationentry,typentry,monthentry,startentry,endentry,periodentry,adressentry,reasonentry
    
    idlabel = tkinter.Label(mainbox, text="Employee ID", bg='#5F9EA0', fg='#FFFFFF')
    idlabel.place(relx=0.03, rely=0.2, relwidth=0.1)

    identry = tkinter.Entry(mainbox)
    identry.place(relx=0.15, rely=0.2, relwidth=0.2)

    empnamelabel = tkinter.Label(mainbox, text="Employee Name", bg='#5F9EA0', fg='#FFFFFF')
    empnamelabel.place(relx=0.03, rely=0.6, relwidth=0.12)

    empnamentry = tkinter.Entry(mainbox)
    empnamentry.place(relx=0.15, rely=0.6, relwidth=0.2)

    designationlabel = tkinter.Label(mainbox, text="Position", bg='#5F9EA0', fg='#FFFFFF')
    designationlabel.place(relx=0.64, rely=0.2, relwidth=0.1)

    designationentry = tkinter.Entry(mainbox)
    designationentry.place(relx=0.765, rely=0.2, relwidth=0.2)

    leaveboxlabel = tkinter.Label(frame, text='Leave application form', bg='#5F9EA0', fg='#FFFFFF')
    leaveboxlabel.place(relx=0.05, rely=0.225)
       
    leavebox = tkinter.Canvas(frame,height=450, width=800, bd=2, bg='#5F9EA0')# the box for leave application form
    leavebox.place(relx=0.05, rely=0.25)

    monthlabel = tkinter.Label(leavebox, text="Month", bg='#5F9EA0', fg='#FFFFFF')
    monthlabel.place(relx=0.03, rely=0.1, relwidth=0.1)

    monthentry = tkinter.Entry(leavebox)
    monthentry.place(relx=0.15, rely=0.1, relwidth=0.2)

    startlabel = tkinter.Label(leavebox, text="Start date", bg='#5F9EA0', fg='#FFFFFF')
    startlabel.place(relx=0.03, rely=0.2, relwidth=0.1)

    startentry = tkinter.Entry(leavebox)
    startentry.place(relx=0.15, rely=0.2, relwidth=0.2)

    endlabel = tkinter.Label(leavebox, text="End date", bg='#5F9EA0', fg='#FFFFFF')
    endlabel.place(relx=0.64, rely=0.1, relwidth=0.1)

    endentry = tkinter.Entry(leavebox)
    endentry.place(relx=0.765, rely=0.1, relwidth=0.2)

    periodlabel = tkinter.Label(leavebox, text="Leave period", bg='#5F9EA0', fg='#FFFFFF')
    periodlabel.place(relx=0.64, rely=0.2, relwidth=0.1)

    periodentry = tkinter.Entry(leavebox)
    periodentry.place(relx=0.765, rely=0.2, relwidth=0.2)

    reasonlabel = tkinter.Label(leavebox, text="Reason for leave", bg='#5F9EA0', fg='#FFFFFF')
    reasonlabel.place(relx=0.02, rely=0.3, relwidth=0.13)

    global reasoncombo
    
    list1 = ["Maternity", "Paternity", "Sick", "Vacation", "Other"]
    reasoncombo = ttk.Combobox(leavebox, values=list1)
    reasoncombo.place(relx=0.15, rely=0.3, relwidth=0.3)

    adresslabel = tkinter.Label(leavebox, text="Adress during \n leave period", bg='#5F9EA0', fg='#FFFFFF')
    adresslabel.place(relx=0.564, rely=0.3, relwidth=0.1)

    adressentry = tkinter.Entry(leavebox)
    adressentry.place(relx=0.665, rely=0.3, relwidth=0.3, relheight=0.6)

    applybutton = tkinter.Button(frame, text='Apply',command = getinput2)
    applybutton.place(relx=0.2, rely=0.9125, relwidth=0.1)

    applybutton = tkinter.Button(frame, text='Exit',command= leavewindow.destroy)
    applybutton.place(relx=0.325, rely=0.9125, relwidth=0.1)

    canvas.pack()

    leavewindow.mainloop()
    return leavewindow

def getinput2():
    global ID,empname,designation,month,start,reason,end,period,address
    ID,empname,designation,month = identry.get(),empnamentry.get(),designationentry.get(),monthentry.get()
    start,reason,end,period,address = startentry.get(),reasoncombo.get(),endentry.get(),periodentry.get(),adressentry.get()

    if ID == '' or empname == '' or designation == '' or month == '' or start == '' or reason == '' or end == '' or period == '' or \
    address == '':
        messagebox.showinfo("Incomplete","Please fill out all entries!")

    else:
        excel2()
        messagebox.showinfo("Saved","Succesfully Saved!")
        leavewindow.destroy()
               
def excel2():
    wb = load_workbook('HRM_Excel.xlsx')
    ws = wb["Leaves"]
    data = [[ID,designation,empname,month,start,reason,end,period,address]]
    for i in data:
        ws.append(i)

    wb.save("HRM_Excel.xlsx")

def viewMessage():

    excel=('HRM_Excel.xlsx')
    wb=xlrd.open_workbook(excel)
    sheet=wb.sheet_by_name("Announcement")

    list1=[]                #list of id
    for i in range(0,sheet.nrows):
        list1.append(sheet.cell_value(i,0))

    global ancwindow

    ancwindow = tkinter.Tk()
    ancwindow.title("View Announcement")
    canvas = tkinter.Canvas(ancwindow, height=700, width=900)

    frame = tkinter.Frame(ancwindow, bg='#5F9EA0')
    frame.place(relwidth=1, relheight=1)

    boxlabel = tkinter.Label(frame, text='Announcement', bg='#5F9EA0', fg='#FFFFFF')
    boxlabel.place(relx=0.05, rely=0.025)

    mainbox = tkinter.Canvas(frame, height=600, width=800, bd=2, bg='#5F9EA0')# the box of main profile
    mainbox.place(relx=0.05, rely=0.05)

    global  combo, subjectentry, messagentry
    empnamelabel = tkinter.Label(mainbox, text="Employee Name", bg='#5F9EA0', fg='#FFFFFF')
    empnamelabel.place(relx=0.03, rely=0.1, relwidth=0.14)

    empnamentry = tkinter.Entry(mainbox)
    empnamentry.place(relx=0.19, rely=0.1, relwidth=0.46)

    list2=["-","-","-","-"]

    def displayinfo():
        list2.clear()
        for i in range(0,sheet.nrows):      
            if empnamentry.get()==sheet.cell_value(i,0):

                list2.append(sheet.cell_value(i,1))
                subjectlabel = tkinter.Label(mainbox, text=list2[0],bg='#5F9EA0', fg='#FFFFFF' )
                subjectlabel.place(relx=0.15, rely=0.2, relwidth=0.5)

                list2.append(sheet.cell_value(i,2))
                messaglabel = tkinter.Label(mainbox, text=list2[1],bg='#5F9EA0', fg='#FFFFFF')
                messaglabel.place(relx=0.15, rely=0.3, relwidth=0.5, relheight=0.5)

                list2.append(sheet.cell_value(i,3))

                
    subjectlabel = tkinter.Label(mainbox, text="Subject", bg='#5F9EA0', fg='#FFFFFF')
    subjectlabel.place(relx=0.03, rely=0.2, relwidth=0.1)

    subjectlabel = tkinter.Label(mainbox)
    subjectlabel.place(relx=0.15, rely=0.2, relwidth=0.5)

    messagelabel = tkinter.Label(mainbox, text="Message", bg='#5F9EA0', fg='#FFFFFF')
    messagelabel.place(relx=0.03, rely=0.3, relwidth=0.1)

    messaglabel = tkinter.Label(mainbox)
    messaglabel.place(relx=0.15, rely=0.3, relwidth=0.5, relheight=0.5)

    savebutton = tkinter.Button(frame, text='View', command = displayinfo)
    savebutton.place(relx=0.2, rely=0.85, relwidth=0.1)
    
    viewbutton = tkinter.Button(ancwindow,text="Exit",command= ancwindow.destroy)
    viewbutton.place(relx=0.4, rely=0.85, relwidth=0.1)

    canvas.pack()

    ancwindow.mainloop()

    return ancwindow

def profileView():
    excel=('HRM_Excel.xlsx')
    wb=xlrd.open_workbook(excel)
    sheet=wb.sheet_by_name("Merged")

    list1=[]                #list of id
    for i in range(0,sheet.nrows):
        list1.append(sheet.cell_value(i,0))

    global window,gradentry,designationentry
    window = tkinter.Tk()
    window.title("Update employee's Personal Data")
    canvas = tkinter.Canvas(window, height=700, width=900) 

    frame = tkinter.Frame(window, bg='#5F9EA0')
    frame.place(relwidth=1, relheight=1)

    boxlabel = tkinter.Label(frame, text='Main Profile', bg='#5F9EA0', fg='#FFFFFF')
    boxlabel.place(relx=0.05, rely=0.025)

    mainbox = tkinter.Canvas(frame, height=100, width=800, bd=2, bg='#5F9EA0')# the box of main profile
    mainbox.place(relx=0.05, rely=0.05)

    global identry,namentry,typentry
    idlabel = tkinter.Label(mainbox, text="Employee ID", bg='#5F9EA0', fg='#FFFFFF')
    idlabel.place(relx=0.03, rely=0.2, relwidth=0.14)


    identry=tkinter.Entry(mainbox)
    identry.place(relx=0.19, rely=0.2, relwidth=0.2)

    list2=["-","-","-","-","-","-","-","-","-","-","-","-","-","-","-","-","-","-","-","-"]

    def displayinfo():
        list2.clear()
        for i in range(1,sheet.nrows):          
            if identry.get()==sheet.cell_value(i,0):
                
                list2.append(sheet.cell_value(i,1))
                designationview = tkinter.Label(mainbox, text=list2[0], bg='#5F9EA0', fg='#FFFFFF')
                designationview.place(relx=0.15, rely=0.6, relwidth=0.2)

                list2.append(sheet.cell_value(i,2))
                empnameview = tkinter.Label(mainbox, text=list2[1], bg='#5F9EA0', fg='#FFFFFF')
                empnameview.place(relx=0.75, rely=0.2, relwidth=0.2)
                nameview = tkinter.Label(personalbox, text=list2[1], bg='#5F9EA0', fg='#FFFFFF')
                nameview.place(relx=0.15, rely=0.1, relwidth=0.2)

                list2.append(sheet.cell_value(i,3))
                sexview = tkinter.Label(personalbox, text=list2[2], bg='#5F9EA0', fg='#FFFFFF')
                sexview.place(relx=0.15, rely=0.2, relwidth=0.2)

                list2.append(sheet.cell_value(i,4))
                bloodview = tkinter.Label(personalbox, text=list2[3], bg='#5F9EA0', fg='#FFFFFF')
                bloodview.place(relx=0.15, rely=0.3, relwidth=0.2)

                list2.append(sheet.cell_value(i,5))
                dateview = tkinter.Label(personalbox, text=list2[4], bg='#5F9EA0', fg='#FFFFFF')
                dateview.place(relx=0.15, rely=0.4, relwidth=0.2)

                list2.append(sheet.cell_value(i,6))
                placeview = tkinter.Label(personalbox, text=list2[5], bg='#5F9EA0', fg='#FFFFFF')
                placeview.place(relx=0.15, rely=0.5, relwidth=0.2)

                list2.append(sheet.cell_value(i,7))
                nationview = tkinter.Label(personalbox, text=list2[6], bg='#5F9EA0', fg='#FFFFFF')
                nationview.place(relx=0.15, rely=0.6, relwidth=0.2)

                list2.append(sheet.cell_value(i,8))
                townview = tkinter.Label(personalbox, text=list2[7], bg='#5F9EA0', fg='#FFFFFF')
                townview.place(relx=0.15, rely=0.7, relwidth=0.2)

                list2.append(sheet.cell_value(i,9))
                houseview = tkinter.Label(personalbox, text=list2[8], bg='#5F9EA0', fg='#FFFFFF')
                houseview.place(relx=0.15, rely=0.8, relwidth=0.2)

                list2.append(sheet.cell_value(i,10))
                mobileview = tkinter.Label(personalbox, text=list2[9], bg='#5F9EA0', fg='#FFFFFF')
                mobileview.place(relx=0.15, rely=0.9, relwidth=0.2)

                list2.append(sheet.cell_value(i,11))
                teleview = tkinter.Label(personalbox, text=list2[10], bg='#5F9EA0', fg='#FFFFFF')
                teleview.place(relx=0.765, rely=0.8, relwidth=0.2)

                list2.append(sheet.cell_value(i,12))
                emailview = tkinter.Label(personalbox, text=list2[11], bg='#5F9EA0', fg='#FFFFFF')
                emailview.place(relx=0.765, rely=0.9, relwidth=0.2)

                list2.append(sheet.cell_value(i,13))
                departmentview = tkinter.Label(jobox, text=list2[12], bg='#5F9EA0', fg='#FFFFFF')
                departmentview.place(relx=0.15, rely=0.125, relwidth=0.2)

                list2.append(sheet.cell_value(i,14))
                salaryview = tkinter.Label(jobox, text=list2[13], bg='#5F9EA0', fg='#FFFFFF')
                salaryview.place(relx=0.15, rely=0.375, relwidth=0.2)

                list2.append(sheet.cell_value(i,15))
                perview = tkinter.Label(jobox, text=list2[14], bg='#5F9EA0', fg='#FFFFFF')
                perview.place(relx=0.15, rely=0.625, relwidth=0.2)


                list2.append(sheet.cell_value(i,16))
                joinview = tkinter.Label(jobox, text=list2[15], bg='#5F9EA0', fg='#FFFFFF')
                joinview.place(relx=0.765, rely=0.125, relwidth=0.2)

                
                list2.append(sheet.cell_value(i,17))
                confirmview = tkinter.Label(jobox, text=list2[16], bg='#5F9EA0', fg='#FFFFFF')
                confirmview.place(relx=0.765, rely=0.375, relwidth=0.2)

               
                list2.append(sheet.cell_value(i,18))
                lastview = tkinter.Label(jobox, text=list2[17], bg='#5F9EA0', fg='#FFFFFF')
                lastview.place(relx=0.765, rely=0.625, relwidth=0.2)

                list2.append(sheet.cell_value(i,19))


    designationlabel = tkinter.Label(mainbox, text="Position", bg='#5F9EA0', fg='#FFFFFF')
    designationlabel.place(relx=0.03, rely=0.6, relwidth=0.1)

    designationview = tkinter.Label(mainbox,text=list2[0], bg='#5F9EA0')
    designationview.place(relx=0.15, rely=0.6, relwidth=0.2)

    empnamelabel = tkinter.Label(mainbox, text="employee name", bg='#5F9EA0', fg='#FFFFFF')
    empnamelabel.place(relx=0.6, rely=0.2, relwidth=0.14)

    empnameview = tkinter.Label(mainbox,text=list2[1], bg='#5F9EA0')
    empnameview.place(relx=0.75, rely=0.2, relwidth=0.2)

    personalboxlabel = tkinter.Label(frame, text='Personal Profile', bg='#5F9EA0', fg='#FFFFFF')
    personalboxlabel.place(relx=0.05, rely=0.225)

    personalbox = tkinter.Canvas(frame,height=300, width=800, bd=2, bg='#5F9EA0')# the box for personal profile
    personalbox.place(relx=0.05, rely=0.25)

    namelabel = tkinter.Label(personalbox, text="Name", bg='#5F9EA0', fg='#FFFFFF')
    namelabel.place(relx=0.03, rely=0.1, relwidth=0.1)

    nameview = tkinter.Label(personalbox,text=list2[1], bg='#5F9EA0')
    nameview.place(relx=0.15, rely=0.1, relwidth=0.2)

    sexlabel = tkinter.Label(personalbox, text="Sex", bg='#5F9EA0', fg='#FFFFFF')
    sexlabel.place(relx=0.03, rely=0.2, relwidth=0.1)

    sexview = tkinter.Label(personalbox,text=list2[2], bg='#5F9EA0')
    sexview.place(relx=0.15, rely=0.3, relwidth=0.2)

    #global bloodentry,datentry,placentry
    bloodlabel = tkinter.Label(personalbox, text="Blood Group", bg='#5F9EA0', fg='#FFFFFF')
    bloodlabel.place(relx=0.03, rely=0.3, relwidth=0.1)

    bloodview = tkinter.Label(personalbox,text=list2[3], bg='#5F9EA0')
    bloodview.place(relx=0.15, rely=0.3, relwidth=0.2)

    datelabel = tkinter.Label(personalbox, text="Date of birth", bg='#5F9EA0', fg='#FFFFFF')
    datelabel.place(relx=0.03, rely=0.4, relwidth=0.1)

    dateview = tkinter.Label(personalbox,text=list2[4], bg='#5F9EA0')
    dateview.place(relx=0.15, rely=0.4, relwidth=0.2)

    placelabel = tkinter.Label(personalbox, text="Place of birth", bg='#5F9EA0', fg='#FFFFFF')
    placelabel.place(relx=0.03, rely=0.5, relwidth=0.1)

    placeview = tkinter.Label(personalbox,text=list2[5], bg='#5F9EA0')
    placeview.place(relx=0.15, rely=0.5, relwidth=0.2)

    nationlabel = tkinter.Label(personalbox, text="Nationality", bg='#5F9EA0', fg='#FFFFFF')
    nationlabel.place(relx=0.03, rely=0.6, relwidth=0.1)

    #global nationentry,townentry,housentry
    nationview = tkinter.Label(personalbox,text=list2[6], bg='#5F9EA0')
    nationview.place(relx=0.15, rely=0.6, relwidth=0.3)

    townlabel = tkinter.Label(personalbox, text="Town", bg='#5F9EA0', fg='#FFFFFF')
    townlabel.place(relx=0.03, rely=0.7, relwidth=0.1)

    townview = tkinter.Label(personalbox,text=list2[7], bg='#5F9EA0')
    townview.place(relx=0.15, rely=0.7, relwidth=0.2)

    houselabel = tkinter.Label(personalbox, text="House Number", bg='#5F9EA0', fg='#FFFFFF')
    houselabel.place(relx=0.03, rely=0.8, relwidth=0.1)

    housview = tkinter.Label(personalbox,text=list2[8], bg='#5F9EA0')
    housview.place(relx=0.15, rely=0.8, relwidth=0.2)

    #global mobilentry,telentry,emailentry,departmententry
    mobilelabel = tkinter.Label(personalbox, text="Mobile Number", bg='#5F9EA0', fg='#FFFFFF')
    mobilelabel.place(relx=0.028, rely=0.9, relwidth=0.12)

    mobileview = tkinter.Label(personalbox,text=list2[9], bg='#5F9EA0')
    mobileview.place(relx=0.15, rely=0.9, relwidth=0.2)

    telelabel = tkinter.Label(personalbox, text="Telephone Number", bg='#5F9EA0', fg='#FFFFFF')
    telelabel.place(relx=0.61, rely=0.8, relwidth=0.13)

    teleview = tkinter.Label(personalbox,text=list2[10], bg='#5F9EA0')
    teleview.place(relx=0.765, rely=0.8, relwidth=0.2)

    emailabel = tkinter.Label(personalbox, text="E-mail", bg='#5F9EA0', fg='#FFFFFF')
    emailabel.place(relx=0.64, rely=0.9, relwidth=0.1)

    emailview = tkinter.Label(personalbox,text=list2[11], bg='#5F9EA0')
    emailview.place(relx=0.765, rely=0.9, relwidth=0.2)

    joboxlabel = tkinter.Label(frame, text='Job detail', bg='#5F9EA0', fg='#FFFFFF')
    joboxlabel.place(relx=0.05, rely=0.7)

    jobox = tkinter.Canvas(frame,height=150, width=800, bd=2, bg='#5F9EA0')# box for job details
    jobox.place(relx=0.05, rely=0.725)

    departmentlabel = tkinter.Label(jobox, text="Department", bg='#5F9EA0', fg='#FFFFFF')
    departmentlabel.place(relx=0.05, rely=0.125, relwidth=0.1)

    departmentview = tkinter.Label(jobox, text=list2[14], bg='#5F9EA0', fg='#FFFFFF')
    departmentview.place(relx=0.15, rely=0.125, relwidth=0.2)

    #global joinentry, confirmentry, lastentry, salaryentry, perentry
    joinlabel = tkinter.Label(jobox, text="Date of Joining", bg='#5F9EA0', fg='#FFFFFF')
    joinlabel.place(relx=0.64, rely=0.125, relwidth=0.1)

    joinview = tkinter.Label(jobox, text=list2[15], bg='#5F9EA0', fg='#FFFFFF')
    joinview.place(relx=0.765, rely=0.125, relwidth=0.2)

    confirmlabel = tkinter.Label(jobox, text="Date of confirmation", bg='#5F9EA0', fg='#FFFFFF')
    confirmlabel.place(relx=0.6, rely=0.375, relwidth=0.14)

    confirmview = tkinter.Label(jobox, text=list2[16], bg='#5F9EA0', fg='#FFFFFF')
    confirmview.place(relx=0.765, rely=0.375, relwidth=0.2)

    lastlabel = tkinter.Label(jobox, text="Date of last Increment", bg='#5F9EA0', fg='#FFFFFF')
    lastlabel.place(relx=0.6, rely=0.625, relwidth=0.14)

    lastview = tkinter.Label(jobox, text=list2[17], bg='#5F9EA0', fg='#FFFFFF')
    lastview.place(relx=0.765, rely=0.625, relwidth=0.2)

    salarylabel = tkinter.Label(jobox, text="Salary", bg='#5F9EA0', fg='#FFFFFF')
    salarylabel.place(relx=0.03, rely=0.375, relwidth=0.14)

    salaryview = tkinter.Label(jobox, text=list2[12], bg='#5F9EA0', fg='#FFFFFF')
    salaryview.place(relx=0.15, rely=0.375, relwidth=0.2)

    perlabel = tkinter.Label(jobox, text="Salary per day", bg='#5F9EA0', fg='#FFFFFF')
    perlabel.place(relx=0.03, rely=0.625, relwidth=0.14)

    perview = tkinter.Label(jobox, text=list2[13], bg='#5F9EA0', fg='#FFFFFF' )
    perview.place(relx=0.15, rely=0.625, relwidth=0.2)

    viewbutton = tkinter.Button(frame, text='View', command= displayinfo)
    viewbutton.place(relx=0.1, rely=0.955, relwidth=0.1)


    cancelbutton = tkinter.Button(frame, text='Exit', command = window.destroy)
    cancelbutton.place(relx=0.225, rely=0.955, relwidth=0.1)

    canvas.pack()

    window.mainloop()
    return window

def Logout():
    def LG():
        Home_window.destroy()
        win.destroy()
    global win
    win=tkinter.Tk()
    win.geometry('150x150')
    win.title('Logout')
    tkinter.Label(win, text="Are you sure?").pack()
    tkinter.Button(win, text="Ok", command= LG).place(x = '50', y = '30')
    tkinter.Button(win, text="Cancel", command=win.destroy).place(x = '50', y = '70')
   
    win.mainloop()
    return win

def homescreen():
    global Home_window
    Home_window = tkinter.Toplevel()
    #Home_window = Tk()

    Home_window.title(" Employee Home page")
    Home_window.geometry("1300x700")
    Home_window.config(bg = 'gray')
    Home_window.resizable(False, False)
    menubar = tkinter.Menu(Home_window)
    Home_window.config(menu=menubar)
    Employee = tkinter.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Employee", menu=Employee)
    Employee.add_command(label="View Profile", command=profileView)

    leavemenu = tkinter.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Leaves", menu=leavemenu)
    leavemenu.add_command(label="Leave Application", command=leave)

    Announcementmenu = tkinter.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Announcement", menu=Announcementmenu)
    Announcementmenu.add_command(label="Open Announcement", command=viewMessage)

    Logoutmenu = tkinter.Menu(menubar,tearoff=0)
    menubar.add_command(label="Logout", command=Logout)

    path =( "aait2.jpg")
    img = ImageTk.PhotoImage(Image.open(path))
    #tkinter.Frame(Home_window, width="1000", height="400").pack()
    panel = tkinter.Label(Home_window, image = img)
    panel.place(x="10", y="60")

    Home_window.mainloop()


#homescreen()
