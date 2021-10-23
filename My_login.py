import tkinter
from tkinter import *
import PIL
from PIL import ImageTk, Image
from admin__login import adminscreen
from emp_login import first_screen

global rootchoose

rootchoose= Tk()
rootchoose.geometry("1300x700")
rootchoose.resizable(False, False)
rootchoose.title("Login")
path = ("aait4.jpg")


img = ImageTk.PhotoImage(Image.open(path))
panel = Label(rootchoose, image = img)
panel.place(x="10", y="10")

label = Label(rootchoose, text = "Welcome To AAit Human Resource Managment", bg = "gainsboro", fg = 'brown4',font = \
              ("Times","18","bold")).place(x="20", y="50")


label2 = Label(rootchoose, text = "Login as:", bg = "gainsboro", font =("Times","25","bold")).place(x="150", y="200")

adminbutton = Button(rootchoose, text = 'Admin', font =("Times","18","bold"),command = adminscreen)\
              .place(x = "170", y = "270")


Empbutton = Button(rootchoose, text = 'Employee', font =("Times","18","bold"),command = first_screen)\
            .place(x = "170", y = "340")

rootchoose.mainloop()


