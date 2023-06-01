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

# date
import datetime

# cef
import threading
import sys
from cefpython3 import cefpython as cef

class HAMBOGOGA :
    def __init__(self) :
        self.window_main = Tk()
        self.window_main.geometry("1200x900")

        # data
        self.ReadExelFile()

        # Init
        self.InitAppTitle()
        self.InitLocationListBox()
        self.InitSearchLayout()
        self.InitInfoButton()
        self.InitInfoCanvas()
        self.InitAPIInfo()
        self.InitMap()

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
        self.entry_Search.insert(0, "정왕1동")
        self.entry_Search.place(x= 315, y= 84)

        self.button_Search = Button(self.window_main, text= "검색", width=5, command= self.Clicked_Search)
        self.button_Search.place(x= 535, y= 80)

    def InitInfoButton(self) :
        self.button_PM = Button(self.window_main, text= "미세먼지 정보", width= 33, height= 1, command= self.Show_PM)
        self.button_PM.place(x= 311, y= 150)
        
        self.butto_Weather = Button(self.window_main, text= "날씨 정보", width= 33, height= 1, command= self.Show_Weather)
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

    def InitAPIInfo(self) :
        self.serviceKey = "QHpOtm0e0OwX2cl8WWXuWGQoaOkbXfXYjF60tquzusBWCg3488dfLbTLACxkPJr1EyPxSYd27VCOUh6ZS+RhPQ=="

    def InitMap(self) :
        self.frame_Map = Frame(self.window_main, width= 570, height= 870)
        self.frame_Map.place(x= 615, y= 15)
        thread = threading.Thread(target= self.Show_Map, args= (self.frame_Map,))
        thread.daemon = True
        thread.start()

    # about map
    def Show_Map(self, frame) :
        sys.excepthook = cef.ExceptHook
        window_info = cef.WindowInfo(frame.winfo_id())
        window_info.SetAsChild(frame.winfo_id(), [0, 0, 570, 870])
        cef.Initialize()
        browser = cef.CreateBrowserSync(window_info, url= 'https://www.misemise.co.kr/')
        cef.MessageLoop()

    # about title
    def Clicked_Title(self) :
        webbrowser.open("https://www.tukorea.ac.kr")

    # about listbox
    def Selected_ListBox(self, event) :
        self.entry_Search.delete(0, 'end')
        self.entry_Search.insert(0, self.LocalData_ws.cell(event.widget.curselection()[0] + 1, 3).value)

    # about search
    def Clicked_Search(self) :
        # PM
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

        if not items :
            location = ""

            for row in self.LocalData_ws.rows :
                if row[2].value == self.entry_Search.get() :
                    location = row[1].value
                    break

            callbackURL = "http://apis.data.go.kr/B552584/ArpltnInforInqireSvc/getMsrstnAcctoRltmMesureDnsty"
            params = '?' + urlencode({
                quote_plus("serviceKey"): self.serviceKey,
                quote_plus("returnType"): "xml",
                quote_plus("numOfRows"): "10",
                quote_plus("pageNo"): "1",
                quote_plus("stationName"): location,
                quote_plus("dataTerm"): "DAILY",
                quote_plus("ver"): "1.4"
            })

            url = callbackURL + params
            response_body = urlopen(url).read()
            root = ET.fromstring(response_body.decode('utf-8'))
            items = root.findall(".//item")
            

        self.allinfo_PM = []
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

            self.allinfo_PM.append(info_PM)

        self.Show_PM()

        # Weather fcst
        today = re.sub(r'[^0-9]', '', str(datetime.date.today()))   # date

        # x, y
        nx, ny = -1, -1

        for row in self.LocalData_ws.rows :
            if row[2].value == self.entry_Search.get() :
                nx, ny = row[3].value, row[4].value
                break

        if nx != -1 and ny != -1 :
            callbackURL = "http://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getVilageFcst"
            params = '?' + urlencode({
                quote_plus("serviceKey"): self.serviceKey,
                quote_plus("pageNo"): "1",
                quote_plus("numOfRows"): "500",
                quote_plus("dataType"): "XML",
                quote_plus("base_date"): today,
                quote_plus("base_time"): "0200",
                quote_plus("nx"): nx,
                quote_plus("ny"): ny
            })

            url = callbackURL + params
            response_body = urlopen(url).read()
            root = ET.fromstring(response_body.decode('utf-8'))
            items = root.findall(".//item")

            self.allinfo_Weather = dict()
            for item in items :
                if item.findtext("fcstDate") == item.findtext("baseDate") :
                    fcsttime = item.findtext("fcstTime")

                    if fcsttime not in self.allinfo_Weather :
                        self.allinfo_Weather[fcsttime] = dict()
                    self.allinfo_Weather[fcsttime][item.findtext("category")] = item.findtext("fcstValue")

        # Weather now
        nowtime = datetime.datetime.now().strftime("%H")

        if nowtime[0] == '0' :
            nowtime = nowtime[0] + str(eval(nowtime[1]) - 1) + "00"
        else :
            nowtime = str(eval(nowtime) - 1) + "00"

        # x, y
        nx, ny = -1, -1

        for row in self.LocalData_ws.rows :
            if row[2].value == self.entry_Search.get() :
                nx, ny = row[3].value, row[4].value
                break

        if nx != -1 and ny != -1 :
            callbackURL = "http://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getUltraSrtNcst"
            params = '?' + urlencode({
                quote_plus("serviceKey"): self.serviceKey,
                quote_plus("pageNo"): "1",
                quote_plus("numOfRows"): "500",
                quote_plus("dataType"): "XML",
                quote_plus("base_date"): today,
                quote_plus("base_time"): nowtime,
                quote_plus("nx"): nx,
                quote_plus("ny"): ny
            })

            url = callbackURL + params
            response_body = urlopen(url).read()
            root = ET.fromstring(response_body.decode('utf-8'))
            items = root.findall(".//item")

            self.nowinfo_Weather = dict()

            self.nowinfo_Weather["baseDate"] = today
            self.nowinfo_Weather["baseTime"] = nowtime

            for item in items :
                self.nowinfo_Weather[item.findtext("category")] = item.findtext("obsrValue")

    # PMInfo
    def Show_PM(self) :
        self.Show_PMInfo2(self.allinfo_PM)

    def Show_PMInfo2(self, info_pm) :
        self.canvas_Info2.delete('all')

        self.Create_Rectangle_In_Canvas("측정 시간 : " + info_pm[0]["dataTime"], 0, 0, 0)
        self.Create_Rectangle_In_Canvas("측정소 : " + info_pm[0]["stationName"], 0, 1, 0)
        self.Create_Rectangle_In_Canvas("미세먼지 : " + info_pm[0]["pm10"] + " ㎍/㎥", 1, 0, eval(info_pm[0]["pm10Grade"]))
        self.Create_Rectangle_In_Canvas("초미세먼지 : " + info_pm[0]["pm25"] + " ㎍/㎥", 1, 1, eval(info_pm[0]["pm25Grade"]))
        self.Create_Rectangle_In_Canvas("오존 : " + info_pm[0]["o3"] + " ppm", 2, 0, eval(info_pm[0]["o3Grade"]))
        self.Create_Rectangle_In_Canvas("이산화질소 : " + info_pm[0]["no2"] + " ppm", 2, 1, eval(info_pm[0]["no2Grade"]))
        self.Create_Rectangle_In_Canvas("일산화탄소 : " + info_pm[0]["co"] + " ppm", 3, 0, eval(info_pm[0]["coGrade"]))
        self.Create_Rectangle_In_Canvas("아황산가스 : " + info_pm[0]["so2"] + " ppm", 3, 1, eval(info_pm[0]["so2Grade"]))

        # grade
        self.Create_Rectangle_In_Canvas("좋음", 11, 0, 1)
        self.Create_Rectangle_In_Canvas("보통", 11, 1, 2)
        self.Create_Rectangle_In_Canvas("나쁨", 12, 0, 3)
        self.Create_Rectangle_In_Canvas("매우 나쁨", 12, 1, 4)

    def Create_Rectangle_In_Canvas(self, intext, row, col, grade) :
        color = ["white", "SkyBlue1", "pale green", "yellow2", "IndianRed1"]

        test = self.canvas_Info2.create_text(142 + 284 * col, 10 + 20 * row, text= intext)
        rect = self.canvas_Info2.create_rectangle(self.canvas_Info2.bbox(test), fill= color[grade])
        x1, y1, x2, y2 = self.canvas_Info2.coords(rect)
        self.canvas_Info2.coords(rect, 284 * col, y1, 284 * col + 284, y2)
        self.canvas_Info2.tag_lower(rect, test)

    # Weather Info
    def Show_Weather(self) :
        self.Show_WeatherInfo2(self.nowinfo_Weather)

    def Show_WeatherInfo2(self, info_pm) :
        self.canvas_Info2.delete('all')

        self.Create_Rectangle_In_Canvas("발표 일자 : " + info_pm["baseDate"], 0, 0, 0)
        self.Create_Rectangle_In_Canvas("발표 시각 : " + info_pm["baseTime"], 0, 1, 0)
        self.Create_Rectangle_In_Canvas("현재 기온 : " + info_pm["T1H"] + " ℃", 1, 0, 0)
        self.Create_Rectangle_In_Canvas("현재 습도 : " + info_pm["REH"] + " %", 1, 1, 0)
        self.Create_Rectangle_In_Canvas("1시간 강수량 : " + info_pm["RN1"] + " mm", 2, 0, 0)
        self.Create_Rectangle_In_Canvas("강수 형태 : " + info_pm["PTY"], 2, 1, 0)
        self.Create_Rectangle_In_Canvas("동서방향 풍속 : " + info_pm["UUU"] + " m/s", 3, 0, 0)
        self.Create_Rectangle_In_Canvas("남북방향 풍속 : " + info_pm["VVV"] + " m/s", 3, 1, 0)
        self.Create_Rectangle_In_Canvas("현재 풍향 : " + info_pm["VEC"] + " °", 4, 0, 0)
        self.Create_Rectangle_In_Canvas("현재 풍속 : " + info_pm["WSD"] + " m/s", 4, 1, 0)


HAMBOGOGA()