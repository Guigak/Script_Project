from tkinter import *
from tkinter import font

class HAMBOGOGA :
    def __init__(self) :
        self.window_main = Tk()
        self.window_main.geometry("600x900")

        # Init
        self.InitAppTitle()
        self.InitLocationListBox()

        self.window_main.mainloop()

    def InitAppTitle(self) :
        self.button_AppTitle = Button(self.window_main, text= "HAMBOGOGA", width= 70, height= 2)    # height 1 : 25?
        self.button_AppTitle.place(x= 15, y= 15)

    def InitLocationListBox(self) :
        self.scrollbar_Location = Scrollbar(self.window_main)
        self.scrollbar_Location.place(x=290, y=80)

        self.listbox_Location = Listbox(self.window_main, width= 33, height= 10, yscrollcommand= self.scrollbar_Location.set)
        self.listbox_Location.place(x= 15, y= 80)

HAMBOGOGA()