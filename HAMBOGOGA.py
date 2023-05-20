from tkinter import *
from tkinter import font

class HAMBOGOGA :
    def __init__(self) :
        self.window_main = Tk()
        self.window_main.geometry("600x900")

        # Init
        self.InitAppTitle()

        self.window_main.mainloop()

    def InitAppTitle(self) :
        self.button_AppTitle = Button(self.window_main, text= "HAMBOGOGA", width= 70, height= 2)
        self.button_AppTitle.place(x= 15, y= 15)

HAMBOGOGA()