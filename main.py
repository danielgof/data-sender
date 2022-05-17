# from cgitb import text
# from posixpath import sep
import tkinter as tk
# import tkinter.ttk as ttk
import tkinter.filedialog as fd
import requests
# import json
import xlrd
from collections import OrderedDict
from tkinter.ttk import Style
import time
# from datetime import datetime
from tkinter import *



class App(tk.Tk):

    """"старт приложения"""
    def __init__(self):
        super().__init__()

        self.style = Style()
        self.style.theme_use("alt")

        """инициализация кнопок"""
        self.quitButton = tk.Button(self, text="Выход", font=("Arial Bold", 16), fg="RED", command=self.quit)
        self.quitButton.place(x=700, y=500)

        self.btn_file = tk.Button(self, text="Выбрать файл", font=("Arial Bold", 16),
                            command=self.choose_file)
        self.btn_zakl = tk.Button(self, text="Парсинг и отправка заключений на сервер", font=("Arial Bold", 16),
                            command=self.select_zakl)
        self.btn_serf = tk.Button(self, text="Парсинг и отправка свидетельств на сервер", font=("Arial Bold", 16),
                            command=self.select_serf)
        self.btn_open = tk.Button(self, text="Результат отправки на сервер", font=("Arial Bold", 16),
                            command=self.open_window)


        """размещение кнопок"""                    
        self.btn_file.pack(padx=60, pady=20)
        self.btn_zakl.pack(padx=60, pady=10)
        self.btn_serf.pack(padx=60, pady=10)
        self.btn_open.pack(padx=60, pady=20)


        # self.backGroundImage = PhotoImage(file=".png")
        # self.backGroundImageLable = Label(self, image=self.backGroundImage)
        # self.backGroundImageLable.place(x=0, y=0)


    """функция выбора файла"""
    def choose_file(self):

        filetypes = (("Excel файл", "*.xls*"),
                     ("Любой", "*"))

        self.filename = fd.askopenfilename(title="Открыть файл", initialdir="/",
                                      filetypes=filetypes)
        # with open(self.filename, "r", encoding="utf8") as f:
        #     self.data = json.load(f)
        #     print(json.dumps(self.data,
        #         sort_keys=False,
        #         indent=4,
        #         ensure_ascii=False,
        #         separators=(',', ': ')))

        if self.filename:
            print(self.filename)



    """открытие дополнительного окна"""
    def open_window(self):
        sub = tk.Toplevel(app)
        sub.transient(app)
        sub.title("НТЦ РОССЕТИ ФСК ЕС")
        sub.geometry("580x200+460+420")
        mylabel = Label(sub, text ='Scrollbars', font = "30")  
        mylabel.pack() 
        myscroll = Scrollbar(sub) 
        myscroll.pack(side = RIGHT, fill = Y) 
        mylist = tk.Label(sub, text=self.str).pack(padx=30, pady=30)

        mylist.pack(side = LEFT, fill = BOTH )    
        myscroll.config(command = mylist.yview) 
            



    """"парсинг и отправка заключений на сервер"""
    def select_zakl(self):
        
        self.results = []
        wb = xlrd.open_workbook(self.filename)
        sh = wb.sheet_by_index(0)
        cnt = 0

        for rownum in range(1, sh.nrows):
            data = OrderedDict()
            row_values = sh.row_values(rownum)
            data['COK_ID'] = str(row_values[0])
            data['PERSON_REQUISITES'] = row_values[1]
            data['URL'] = row_values[2]
            data['EX_DATE_START'] = row_values[3]
            data['EX_DATE_END'] = row_values[4]
            data['RECOMMENDATIONS'] = str(row_values[5])
            data['QUAL_ID'] = row_values[6]
            data['DATE'] = row_values[7]
            data['THEORY_PLATFORM_ID'] = row_values[8]
            data['PRACTICE_PLATFORM_ID'] = row_values[9]
            data['PROTOCOL_URL'] = row_values[10]

            dct = dict(data)
            dct2 = OrderedDict([("type", "2"), ("login", "011"), ("fields", dct), ("solution", "1")])
            dct3 = dict(dct2)
            dct4 = OrderedDict([("method", "/cert/reg"), ("token", "your_token"), ("data", dct3)])
            dct5 = dict(dct4)
            df = OrderedDict([("apiToken", "your_apiToken"), ("dstService", "your_dstService"), ("sync", True), ("data", dct5)])
            self.df = dict(df)
            # print(self.df)
            # print(json.dumps(self.df,
            #     sort_keys=False,
            #     indent=4,
            #     ensure_ascii=False,
            #     separators=(',', ': ')))
            response = requests.post("your_API", json=self.df)
            print(dict(response.json()))
            if dict(response.json()).get("success") == True:
                cnt += 1
            self.results.append(response.json())
        # print(f"result: {self.results}")
        self.str1 = ''.join(str(e)+'\n' for e in self.results)
        self.str2 = ''.join(f"Чило успешно отправленных файлов: {cnt}")
        self.str = self.str1 + self.str2
        # with open(f"{datetime.datetime.now():%Y-%m-%d_%H:%M:%S}.txt", "w") as f:
        #     f.write(self.str)
        date = time.strftime("%Y-%m-%d")
        with open(f"date{date}.txt", "w") as f:
            f.write(self.str)
        # print(f"Чило успешно отправленных файлов: {cnt}")
        # print(self.str)
        

    """парсинг и отправка сертификатов на сервер"""
    def select_serf(self):
        
        self.results = []
        wb = xlrd.open_workbook(self.filename)
        sh = wb.sheet_by_index(0)
        cnt = 0

        for rownum in range(1, sh.nrows):
            data = OrderedDict()
            row_values = sh.row_values(rownum)
            data['COK_ID'] = str(row_values[0])
            data['PERSON_REQUISITES'] = row_values[1]
            data['URL'] = row_values[2]
            data['QUAL_ID'] = row_values[3]
            data['DATE'] = row_values[4]
            data['THEORY_PLATFORM_ID'] = row_values[5]
            data['PRACTICE_PLATFORM_ID'] = row_values[6]
            data['PROTOCOL_URL'] = row_values[7]

            dct = dict(data)
            dct2 = OrderedDict([("type", "2"), ("login", "011"), ("fields", dct), ("solution", "1")])
            dct3 = dict(dct2)
            dct4 = OrderedDict([("method", "/cert/reg"), ("token", "your_token"), ("data", dct3)])
            dct5 = dict(dct4)
            df = OrderedDict([("apiToken", "your_apiToken"), ("dstService", "your_dstService"), ("sync", True), ("data", dct5)])
            self.df = dict(df)
            # print(self.df)
            # print(json.dumps(self.df,
            #     sort_keys=False,
            #     indent=4,
            #     ensure_ascii=False,
            #     separators=(',', ': ')))

            response = requests.post("your_API", json=self.df)
            print(dict(response.json()))
            if dict(response.json()).get("success") == True:
                cnt += 1
            self.results.append(str(response.json()))
        self.str1 = ''.join(str(e)+'\n' for e in self.results)
        self.str2 = ''.join(f"Чило успешно отправленных файлов: {cnt}")
        self.str = self.str1 + self.str2
        date = time.strftime("%Y-%m-%d")
        with open(f"date{date}.txt", "w") as f:
            f.write(self.str)


    # """вывод времени"""
    # def update_time(self):
    #     a = Label(app, font=("helvetica", 15))
    #     a.config(text=f"{datetime.now():%H:%M:%S}")
       
    #     a.pack()


if __name__ == "__main__":
    app = App()
    app.geometry("800x600+350+100")
    app.title("НТЦ РОССЕТИ ФСК ЕС")
    app["bg"] = "Cadetblue1"
    # app.Label()
    # app.after(1, app.update_time())
    app.mainloop()