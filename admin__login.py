import tkinter
from tkinter import *
from Human_Resource_Admin import mainscreen

def adminscreen():
    global adminroot
    adminroot = Tk()
    adminroot.title('Admin Login')
    adminroot.geometry('700x700')
    adminroot.config(bg = "gainsboro")

    labeladmin = Label(adminroot, text = "Welcome Admin!",fg = 'brown4', bg = "gainsboro", font = ("Times","25","bold")).\
            place(x="230", y="50")
    labelsign = Label(adminroot, text = "Sign In",fg = 'brown4', bg = "gainsboro", font = ("Times","20")).\
            place(x="300", y="130")

    labeluser = Label(adminroot, text = "Username *", bg = "gainsboro", font = ("Times","14")).\
            place(x="200", y="200")

    global userentry, passentry
    
    userentry = Entry(adminroot, font =("Times","13"))
    userentry.place(x = "200", y = "240", width = "300", height = "30")

    passlabel = Label(adminroot, text = "Password *", bg = "gainsboro", font = ("Times","14")).\
            place(x="200", y="300")

    passentry = Entry(adminroot, font =("Times","13"), show = '*')
    passentry.place(x = "200", y = "340", width = "300", height = "30")

    global signbutton
    
    signbutton = Button(adminroot, text='Sign In', bg = 'brown4', fg = 'white', font = ('15'), command = login)
    signbutton.place(x = "280", y = "400", width = '100', height = '30')
    
    adminroot.mainloop()

    login()
    

def login():
    username = userentry.get()
    password = passentry.get()
    if username == "admin" and password == '12345678':
        signbutton.config(command = mainscreen)
        return adminroot.destroy(),mainscreen()

    elif username == '' and password == '':
        messagebox.showinfo("Incomplete","Please fill all entries")
        adminroot.mainloop()
        return messagebox.showinfo("Incomplete","Please fill all entries")
        
    else:
        labelwrong = Label(adminroot, text = "Wrong username or password!",fg = 'red', bg = "gainsboro", font = \
                    ("Times","18")).place(x="200", y="440")
        adminroot.mainloop()
        
        

