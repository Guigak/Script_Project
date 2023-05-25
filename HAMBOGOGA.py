# tkinter
from tkinter import *
from tkinter import font

# webbrowser
import webbrowser

# excel
from openpyxl import load_workbook

# xml
import urllib
from urllib.request import urlopen
from urllib.parse import urlencode, unquote, quote_plus
import xml.etree.ElementTree as ET

# re
import re

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
        self.InitAPIInfo()

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
        self.button_AppTitle = Button(self.window_main, text= "HAMBOGOGA", width= 70, height= 2, command= self.Clicked_Title)    # height 1 : 25?
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

        self.button_Search = Button(self.window_main, text= "검색", width=5, command= self.Clicked_Search)
        self.button_Search.place(x= 535, y= 80)

    def InitInfoButton(self) :
        self.button_PM = Button(self.window_main, text= "미세먼지 정보", width= 33, height= 1)
        self.button_PM.place(x= 311, y= 150)
        
        self.butto_Weather = Button(self.window_main, text= "날씨 정보", width= 33, height= 1)
        self.butto_Weather.place(x= 311, y= 200)
        
        self.butto_Stock = Button(self.window_main, text= "주식 정보", width= 33, height= 1)
        self.butto_Stock.place(x= 311, y= 250)

    def InitInfoCanvas(self) :
        # 1
        self.canvas_Info = Canvas(self.window_main, width= 567, height= 300, bg= 'white')
        self.canvas_Info.place(x= 15, y= 300)

        # 2
        self.canvas_Info2 = Canvas(self.window_main, width= 567, height= 260, bg= 'white')
        self.canvas_Info2.place(x= 15, y= 620)

        # test
        test1 = self.canvas_Info2.create_text(142, 10, text= "test1")
        rect1 = self.canvas_Info2.create_rectangle(self.canvas_Info2.bbox(test1), fill= 'red')
        x1, y1, x2, y2 = self.canvas_Info2.coords(rect1)
        self.canvas_Info2.coords(rect1, 0, y1, 284, y2)
        self.canvas_Info2.tag_lower(rect1, test1)

        test2 = self.canvas_Info2.create_text(426, 10, text= "test2")
        rect2 = self.canvas_Info2.create_rectangle(self.canvas_Info2.bbox(test2), fill= 'yellow')
        x1, y1, x2, y2 = self.canvas_Info2.coords(rect2)
        self.canvas_Info2.coords(rect2, 284, y1, 567, y2)
        self.canvas_Info2.tag_lower(rect2, test2)

        test3 = self.canvas_Info2.create_text(142, 30, text= "test3")
        rect3 = self.canvas_Info2.create_rectangle(self.canvas_Info2.bbox(test3), fill= 'green')
        x1, y1, x2, y2 = self.canvas_Info2.coords(rect3)
        self.canvas_Info2.coords(rect3, 0, y1, 284, y2)
        self.canvas_Info2.tag_lower(rect3, test3)

        test4 = self.canvas_Info2.create_text(426, 30, text= "test4")
        rect4 = self.canvas_Info2.create_rectangle(self.canvas_Info2.bbox(test4), fill= 'blue')
        x1, y1, x2, y2 = self.canvas_Info2.coords(rect4)
        self.canvas_Info2.coords(rect4, 284, y1, 567, y2)
        self.canvas_Info2.tag_lower(rect4, test4)

    def InitAPIInfo(self) :
        self.serviceKey = "QHpOtm0e0OwX2cl8WWXuWGQoaOkbXfXYjF60tquzusBWCg3488dfLbTLACxkPJr1EyPxSYd27VCOUh6ZS+RhPQ=="

    # about title
    def Clicked_Title(self) :
        webbrowser.open("https://www.tukorea.ac.kr")

    # about listbox
    def Selected_ListBox(self, event) :
        self.entry_Search.delete(0, 'end')
        self.entry_Search.insert(0, self.LocalData_ws.cell(event.widget.curselection()[0] + 1, 3).value)

    # about search
    def Clicked_Search(self) :
        callbackURL = "http://apis.data.go.kr/B552584/ArpltnInforInqireSvc/getMsrstnAcctoRltmMesureDnsty"
        params = '?' + urlencode({
            quote_plus("serviceKey"): self.serviceKey,
            quote_plus("returnType"): "xml",
            quote_plus("numOfRows"): "10",
            quote_plus("pageNo"): "1",
            quote_plus("stationName"): re.sub('\d+', '', self.entry_Search.get()),
            quote_plus("dataTerm"): "DAILY",
            quote_plus("ver"): "1.4"
        })

        url = callbackURL + params
        response_body = urlopen(url).read()
        root = ET.fromstring(response_body.decode('utf-8'))
        items = root.findall(".//item")

        allinfo_PM = []
        for item in items :
            info_PM = {
                "dataTime": item.findtext("dataTime"),
                "stationName": item.findtext("stationName"),
                "pm10": item.findtext("pm10Value"),
                "pm10Grade": item.findtext("pm10Grade"),
                "pm25": item.findtext("pm25Value"),
                "pm25Grade": item.findtext("pm25Grade"),
                "o3": item.findtext("o3Value"),
                "o3Grade": item.findtext("o3Grade"),
                "no2": item.findtext("no2Value"),
                "no2Grade": item.findtext("no2Grade"),
                "co": item.findtext("coValue"),
                "coGrade": item.findtext("coGrade"),
                "so2": item.findtext("so2Value"),
                "so2Grade": item.findtext("so2Grade")
            }

            allinfo_PM.append(info_PM)

        self.Show_PMInfo2(allinfo_PM)

    def Show_PMInfo2(self, info_pm) :
        self.canvas_Info2.create_text(100, 10, text= "측정 시간 : " + info_pm[0]["dataTime"])
        #self.canvas_Info2.create_text(234, 0, text= "측정소 : " + info_pm[0]["stationName"])



HAMBOGOGA()