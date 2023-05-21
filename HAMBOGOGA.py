from tkinter import *
from tkinter import font

from openpyxl import load_workbook

class HAMBOGOGA :
    def __init__(self) :
        self.window_main = Tk()
        self.window_main.geometry("600x900")

        # data
        self.ReadExelFile()

        # Init
        self.InitAppTitle()
        self.InitLocationListBox()
        self.InitSearchLayout()
        self.InitInfoButton()
        self.InitInfoCanvas()

        self.window_main.mainloop()

    # read file

    def ReadExelFile(self) :
        self.LocalData_wb = load_workbook("LocalData.xlsx", data_only= True)
        self.LocalData_ws = self.LocalData_wb['Sheet1']

        # A : 시도
        # B : 시군구
        # C : 읍면동
        # D : x
        # E : y
        # F : 경도
        # G : 위도

        # # test
        # for row in self.LocalData_ws.rows :
        #     print(row[7].value)

    # init

    def InitAppTitle(self) :
        self.button_AppTitle = Button(self.window_main, text= "HAMBOGOGA", width= 70, height= 2)    # height 1 : 25?
        self.button_AppTitle.place(x= 15, y= 15)

    def InitLocationListBox(self) :
        self.frame_Location = Frame(self.window_main, width= 33, height= 10)
        self.frame_Location.place(x= 15, y= 80)

        self.listbox_Location = Listbox(self.frame_Location, width= 33, height= 10)
        self.listbox_Location.pack(side= 'left', fill= 'y')

        self.scrollbar_Location = Scrollbar(self.frame_Location, command= self.listbox_Location.yview)
        self.scrollbar_Location.pack(side= 'right', fill= 'y')

        self.listbox_Location.config(yscrollcommand= self.scrollbar_Location.set)
        self.listbox_Location.event_generate("<<ListboxSelect>>")
        self.listbox_Location.bind("<<ListboxSelect>>", self.Selected_ListBox)
        
        count = 1
        for row in self.LocalData_ws.rows :
            self.listbox_Location.insert(count, row[7].value)
            count += 1

    def InitSearchLayout(self) :
        self.entry_Search = Entry(self.window_main, width= 25)
        self.entry_Search.insert(0, "정왕동")
        self.entry_Search.place(x= 315, y= 84)

        self.button_Search = Button(self.window_main, text= "검색", width=5)
        self.button_Search.place(x= 535, y= 80)

    def InitInfoButton(self) :
        self.button_PM = Button(self.window_main, text= "미세먼지 정보", width= 33, height= 1)
        self.button_PM.place(x= 311, y= 150)
        
        self.butto_Weather = Button(self.window_main, text= "날씨 정보", width= 33, height= 1)
        self.butto_Weather.place(x= 311, y= 200)
        
        self.butto_Stock = Button(self.window_main, text= "주식 정보", width= 33, height= 1)
        self.butto_Stock.place(x= 311, y= 250)

    def InitInfoCanvas(self) :
        self.canvas_Info = Canvas(self.window_main, width= 567, height= 300, bg= 'white')
        self.canvas_Info.place(x= 15, y= 300)

        self.canvas_Info2 = Canvas(self.window_main, width= 567, height= 260, bg= 'white')
        self.canvas_Info2.place(x= 15, y= 620)

    # about listbox

    def Selected_ListBox(self, event) :
        self.entry_Search.delete(0, 'end')
        self.entry_Search.insert(0, self.LocalData_ws.cell(event.widget.curselection()[0] + 1, 3).value)

HAMBOGOGA()