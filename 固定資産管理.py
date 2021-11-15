import sqlite3
from tkinter import ttk
from tkinter import *
import datetime
# import re
from tkinter import messagebox
import os
import shutil
from tkinter import filedialog
from tkinter.font import BOLD
# import configparser
# from typing import DefaultDict
from PIL import Image
import csv
import re
if not os.path.exists("files"):
    os.mkdir("files")
if not os.path.exists("files\image"):
    os.mkdir("files\image")
if not os.path.exists("files\data"):
    os.mkdir("files\data")
DBName = "files\data\Fixed_asset_MD.db"
con = sqlite3.connect(DBName)
cur = con.cursor()
def GetUsers():
    try:
        con = sqlite3.connect(DBName)
        cur = con.cursor()
        sql = f"Select GID from Users"
        cur.execute(sql)
        a = cur.fetchall()
        con.commit()
        con.close()
        return a
    except Exception as e:
        return e

try:
    # すべての記録を登録しておくテーブル
    cur.execute("""CREATE TABLE Data (ID integer, Image text, Number text Unique, X text, Y text, Detail text)""")

except:
    pass

try:
    # ログ情報を記録。移動とか。
    cur.execute("""CREATE TABLE Log (Time text, User text, ID text, Method text, Detail text, PC_USERNAME text)""")
except:
    pass

try:
    # 座標に応じた場所を登録しておくテーブル
    cur.execute("""CREATE TABLE NamePlaces (Image text, X1 integer, Y1 integer, X2 integer, Y2 integer, Name text Unique)""")
except:
    pass

try:
    # 削除ができる人のデータを保存しておくテーブル
    cur.execute("""CREATE TABLE Users (GID text, Date text, Register text)""")
    cur.execute(f"INSERT INTO Users VALUES (0070200367)")
except:
    pass


con.commit()
con.close()
a = None
subwindow = None
InfoWindow = None
RCStartX, RCStartY, RCEndX, RCEndY = 0, 0, 0, 0
ImagePlaceRegistWindow = None
UserName = None
PCUserName = None
OKUserID = [i[0] for i in GetUsers()]
print(OKUserID)


# 優先!
# 見るところ、場所に対しての一覧があってもいいかも


# 入れ替える(精度管理とか)機能
# ヘルプウインドウのチェックボックス

def SaveLogs(Method, Detail = "", ID = ""):
    # Time , User , ID , Method , Detail , PC_USERNAME
    # Methodは登録、変更、削除、開始。
    global UserName, PCUserName
    try:
        Data = (str(datetime.datetime.now()), UserName, ID, Method, Detail, PCUserName)
        print(Data)
        con = sqlite3.connect(DBName)
        cur = con.cursor()
        sql = f"INSERT INTO Log values {Data}"
        cur.execute(sql)
        con.commit()
        con.close()
        return True
    except Exception as e:
        con.close()
        messagebox.showerror("エラー。", e)
        return False

class main():
    def __init__(self) -> None:
        self.Canvas = None
        global root
        root = Tk()
        root.title("固定資産場所管理")
        root.minsize(300, 100)
        mainstyle = ttk.Style()
        mainstyle.theme_use("clam")
        mainstyle.configure("office.TButton", font=20)
        EntryPlaceButton = ttk.Button(root, text="所在地入力", command=lambda:[self.EntryPlace()], style="office.TButton")
        ShowPlaceButton = ttk.Button(root, text="所在地確認・変更", command=lambda:[self.ShowPlace()], style="office.TButton")
        ToCSVButton = ttk.Button(root, text="データ書き出し(CSV)", command=lambda:[self.ToCSV()], style="office.TButton")
        AddImageButton = ttk.Button(root, text="地図情報(追加、登録、変更)", command=lambda:[self.AddImageWindow()], style="office.TButton")
        EntryPlaceButton.pack(padx=10, pady=10)
        ShowPlaceButton.pack(padx=10, pady=10)
        ToCSVButton.pack(padx=10, pady=10)
        AddImageButton.pack(padx=10, pady=10)
        root.mainloop()
    def ToCSV(self):
        global subwindow
        if subwindow == None or not subwindow.winfo_exists():
            subwindow = Toplevel(root)
        else:
            return
        subwindow.title("データ書き出し(CSV)")
        self.InformationLabel = ttk.Labelframe(subwindow, text="所在情報書き出し情報")
        self.OutputPathEntry = ttk.Entry(self.InformationLabel)
        self.OutputPathEntry.bind("<Double-1>" , self.SelectPath)
        self.ReferenceButton = ttk.Button(self.InformationLabel, text="参照(パスは共通)", command=lambda:[self.SelectPath()])
        self.SubmitButton = ttk.Button(self.InformationLabel, text="所在情報書き出し", command=lambda:[self.ToCSVReal()])
        self.InformationLabel.pack(padx=5, pady=5)
        self.OutputPathEntry.pack(padx=5, pady=5, anchor=W)
        self.ReferenceButton.pack(padx=5, pady=5, anchor=W)
        self.SubmitButton.pack(padx=5, pady=5, anchor=W)


        self.MoveLogOut = ttk.Labelframe(subwindow, text="移動ログ書き出し情報")
        self.NumberEntry = ttk.Combobox(self.MoveLogOut)
        LogRaw = self.GetMoveLogID()
        LogList = [i[0] for i in LogRaw if i[0] != "ForAdminLog"]

        self.NumberEntry["values"] = LogList
        NumberEntryLabel = ttk.Label(self.MoveLogOut, text="登録番号")
        OutputButton = ttk.Button(self.MoveLogOut, text="移動ログ書き出し", command=lambda:[self.MoveLogOutput()])

        self.MoveLogOut.pack(padx=5, pady=5)
        NumberEntryLabel.pack(padx=5, pady=5)
        self.NumberEntry.pack(padx=5, pady=5)
        OutputButton.pack(padx=5, pady=5)
    def MoveLogOutput(self):
        a = self.GetMoveLog()
        a = [i[:5] for i in a]
        
        if not os.path.exists(self.OutputPathEntry.get()):
            messagebox.showinfo("パスなし", "出力先を選択してください。")
            return
        with open(f"{self.OutputPathEntry.get()}/{str(datetime.date.today())}.csv", "w", newline="", encoding="shift-jis") as f:
            writer = csv.writer(f)
            writer.writerow(["Time", "User", "ID", "Method", "Details"])
            writer.writerows(a)
        messagebox.showinfo("Success", f"Success\nPath:{self.OutputPathEntry.get()}/{str(datetime.date.today())}.csv")
        SaveLogs("ログ保存", f"{a[0][3]}", "ForAdminLog")
        subwindow.lift()
    def GetMoveLog(self):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"Select * from Log WHERE ID = '{self.NumberEntry.get()}'"
            cur.execute(sql)
            a = cur.fetchall()
            con.commit()
            con.close()
            return a
        except Exception as e:
            return e
    def GetMoveLogID(self):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"Select DISTINCT ID from Log"
            cur.execute(sql)
            a = cur.fetchall()
            con.commit()
            con.close()
            return a
        except Exception as e:
            return e
    def AddImageWindow(self):
        global subwindow
        self.LogData = []
        self.Canvas = None
        self.ImageFrame = None
        if subwindow == None or not subwindow.winfo_exists():
            subwindow = Toplevel(root)
        else:
            return
        subwindow.title("画像追加")
        mainframe = ttk.Frame(subwindow)
        AddFrame = ttk.Labelframe(mainframe, text="追加フォーム")
        self.AddPathEntry = ttk.Entry(AddFrame)
        AddReferanceButton = ttk.Button(AddFrame, text="参照", command=lambda:[self.AddImageSelectPath()])
        AddButton = ttk.Button(AddFrame, text="追加", command=lambda:[self.ImageCopy()])

        # 追加と変更
        RCFrame = ttk.Labelframe(mainframe, text="登録、変更フォーム")
        self.RCPathEntry = ttk.Combobox(RCFrame, width=18)
        path = "files\image"
        files = os.listdir(path)
        files_file = [f[0:-4] for f in files if os.path.isfile(os.path.join(path, f)) and f[-4:].lower() == ".png"]
        self.RCPathEntry["values"] = files_file
        RCShowButton = ttk.Button(RCFrame, text="表示", command=lambda:[self.RCShowImage()])

        # データの確認
        DataViewFrame = ttk.Labelframe(mainframe, text="登録されたデータ")
        self.DVTree = ttk.Treeview(DataViewFrame)
        self.DVTree["columns"] = (1)
        self.DVTree["show"] = "headings"
        self.DVTree.heading(1, text="登録名称")
        self.DVTree.column(1, width=220)
        self.DVTree.pack(padx=5, pady=5)
        # 画像上に登録されたデータを表示
        self.DVTree.bind("<<TreeviewSelect>>", self.DVSelected)
        self.DVTree.bind("<Double-1>", self.DVEdit)
        self.DVTree.bind("<Control-Delete>", self.DVDelete)

        # 変更確認ウインドウ+変更可能に
        # self.DVTree.bind("<Double-1>", )

        # 画像のフレーム
        self.ImageFrame = ttk.Labelframe(subwindow, text="地図")

        # 配置
        mainframe.grid(row=0, column=0, sticky=N)
        mainframe.pack_propagate(True)
        AddFrame.pack(padx=5, pady=5, anchor=W)
        AddFrame.grid_propagate(True)
        RCFrame.pack(padx=5, pady=5, anchor=W)
        RCFrame.grid_propagate(True)
        DataViewFrame.pack(padx=5, pady=5)
        self.AddPathEntry.grid(row=0, column=0, padx=5, pady=5)
        AddReferanceButton.grid(row=0, column=1, padx=5, pady=5)
        AddButton.grid(row=1, column=1, pady=5, padx=5)
        self.RCPathEntry.grid(row=0, column=0, padx=5, pady=5)
        RCShowButton.grid(row=1, column=1, padx=5, pady=5)
    def ImageCopy(self):
        try:
            dirname = self.AddPathEntry.get()
            distdir = f"files/image/{os.path.basename(dirname)}"
            if not os.path.isfile(dirname):
                return
            shutil.copyfile(dirname, distdir)
            messagebox.showinfo("完了", "追加完了しました。")
            self.AddPathEntry.delete(0, END)
        except Exception as e:
            messagebox.showerror("エラー", e)
    
    def StartCreateSquare(self, event = None):
        try:
            self.Canvas.delete("Square")
        except:
            pass
        global RCStartX, RCStartY, RCEndX, RCEndY
        RCStartX, RCStartY, RCEndX, RCEndY = 0, 0, 0, 0
        RCStartX, RCEndX = event.x, event.x
        RCStartY, RCEndY = event.y, event.y
        self.Canvas.create_rectangle(RCStartX, RCStartY, RCEndX, RCEndY, outline="BLUE", tag="Square", width=2)

    def CreateSquare(self, event = None):
        global RCStartX, RCStartY, RCEndX, RCEndY
        if RCStartX > event.x:
            RCStartX = event.x
        if RCStartY > event.y:
            RCStartY = event.y
        self.Canvas.coords("Square", RCStartX, RCStartY, event.x, event.y)
        RCEndX = event.x
        RCEndY = event.y
    def RCPlaceValidation(self, S):
        # 画像の範囲内であること、大小関係が正しいことの確認
        if re.match(re.compile('[0-9]+'), S) and (int(self.PlaceX1_Entry.get()) < int(self.PlaceX2_Entry.get())) and (int(self.PlaceY1_Entry.get()) < int(self.PlaceY2_Entry.get())) and int(self.PlaceX1_Entry.get()) >= 0 and int(self.PlaceX2_Entry.get()) < self.RCWidth and int(self.PlaceY1_Entry.get()) >= 0 and int(self.PlaceY2_Entry.get()) < self.RCHeight:
            # if (int(self.PlaceX1_Entry.get()) < int(self.PlaceX2_Entry.get())) and (int(self.PlaceY1_Entry.get()) < int(self.PlaceY2_Entry.get())):
            #     if int(self.PlaceX1_Entry.get()) >= 0 and int(self.PlaceX2_Entry.get()) < self.RCWidth and int(self.PlaceY1_Entry.get()) >= 0 and int(self.PlaceY2_Entry.get()) < self.RCHeight:
            self.PlaceEditted()
            return True
        else:
            self.PlaceX1_Entry.delete(0, END)
            self.PlaceX1_Entry.insert(0, RCStartX)
            self.PlaceY1_Entry.delete(0, END)
            self.PlaceY1_Entry.insert(0, RCStartY)
            self.PlaceX2_Entry.delete(0, END)
            self.PlaceX2_Entry.insert(0, RCEndX)
            self.PlaceY2_Entry.delete(0, END)
            self.PlaceY2_Entry.insert(0, RCEndY)
            return False
    def CheckInside(self, Num, high, low):
        # 範囲内か確認する関数
        if low <= Num <= high:
            return True
        else:
            return False


    def EndCreateSquare(self, event = None):
        global RCStartX, RCStartY, RCEndX, RCEndY
        self.Canvas.itemconfig(tagOrId="Square", outline="Red")
        global ImagePlaceRegistWindow
        if ImagePlaceRegistWindow == None or not ImagePlaceRegistWindow.winfo_exists():
            ImagePlaceRegistWindow = Toplevel(subwindow)
            ImagePlaceRegistWindow.grab_set()
        else:
            return
        RegistInfoFrame = ttk.Labelframe(ImagePlaceRegistWindow, text="登録情報入力")
        self.NameLabel = ttk.Label(RegistInfoFrame, text=f"登録名称 {self.RCPathEntry.get()}-")
        self.PlaceNameEntry = ttk.Entry(RegistInfoFrame)
        self.RegistButton = ttk.Button(RegistInfoFrame, text="登録", command=lambda:[self.RCRegist()])
        CancelButton = ttk.Button(RegistInfoFrame, text="取消", command=lambda:[ImagePlaceRegistWindow.destroy()])

        PlaceInfo = ttk.Labelframe(RegistInfoFrame, text="場所データ")
        PlaceX1_Label = ttk.Label(PlaceInfo, text="左上X座標")
        PlaceY1_Label = ttk.Label(PlaceInfo, text="左上Y座標")
        PlaceX2_Label = ttk.Label(PlaceInfo, text="右下X座標")
        PlaceY2_Label = ttk.Label(PlaceInfo, text="右下Y座標")
        placevali = root.register(self.RCPlaceValidation)
        self.PlaceX1_Entry = ttk.Entry(PlaceInfo, width=6, validate='focusout', validatecommand=(placevali, "%P"))
        self.PlaceY1_Entry = ttk.Entry(PlaceInfo, width=6, validate='focusout', validatecommand=(placevali, "%P"))
        self.PlaceX2_Entry = ttk.Entry(PlaceInfo, width=6, validate='focusout', validatecommand=(placevali, "%P"))
        self.PlaceY2_Entry = ttk.Entry(PlaceInfo, width=6, validate='focusout', validatecommand=(placevali, "%P"))

        RegistInfoFrame.pack(padx=5, pady=5, anchor=W)
        RegistInfoFrame.grid_propagate(True)
        self.NameLabel.grid(row=0, column=0, pady=5, sticky=E)
        self.PlaceNameEntry.grid(row=0, column=1, pady=5, sticky=W)
        self.RegistButton.grid(row=2, column=1, padx=5, pady=5)
        CancelButton.grid(row=2, column=0, padx=5, pady=5)

        PlaceInfo.grid(row=1, column=0, columnspan=2, padx=5, pady=5)
        PlaceX1_Label.grid(row=0, column=0, padx=5, pady=5)
        self.PlaceX1_Entry.grid(row=0, column=1, padx=5, pady=5)
        PlaceY1_Label.grid(row=0, column=2, padx=5, pady=5)
        self.PlaceY1_Entry.grid(row=0, column=3, padx=5, pady=5)
        PlaceX2_Label.grid(row=1, column=0, padx=5, pady=5)
        self.PlaceX2_Entry.grid(row=1, column=1, padx=5, pady=5)
        PlaceY2_Label.grid(row=1, column=2, padx=5, pady=5)
        self.PlaceY2_Entry.grid(row=1, column=3, padx=5, pady=5)

        self.PlaceX1_Entry.insert(0, RCStartX)
        self.PlaceY1_Entry.insert(0, RCStartY)
        self.PlaceX2_Entry.insert(0, RCEndX)
        self.PlaceY2_Entry.insert(0, RCEndY)


        # 登録名称のところ、OCR機能ワンチャン
    def RCShowImage(self):
        imagepath = "files/image/" + self.RCPathEntry.get() + ".png"
        if not os.path.isfile(imagepath):
            return
        if self.ImageFrame != None:
            self.ImageFrame.pack_forget()
        self.ImageFrame.grid(row=0, column=1)
        self.RCImageName = self.RCPathEntry.get()
        img = Image.open(imagepath)
        width, height = img.size
        self.RCWidth, self.RCHeight = img.size
        if self.Canvas != None:
            self.Canvas.pack_forget()
            self.Canvas.delete()
        self.Canvas = Canvas(self.ImageFrame, height=height, width=width)
        self.Canvas.pack(anchor=N)
        self.photo_image = PhotoImage(file = imagepath)
        self.Canvas.create_image(width/2, height/2, image=self.photo_image, tag="image")
        
        self.Canvas.tag_bind("image", "<1>", self.StartCreateSquare)
        self.Canvas.tag_bind("image","<Button1-Motion>", self.CreateSquare)
        self.Canvas.tag_bind("image", "<ButtonRelease-1>", self.EndCreateSquare)
        DVData = self.DVGetData()
        for i in self.DVTree.get_children():
            self.DVTree.delete(i)

        for i in DVData:
            self.DVTree.insert("", "end", values=(i))

    def PlaceEditted(self, event = None):
        global RCStartX, RCStartY, RCEndX, RCEndY
        RCStartX = self.PlaceX1_Entry.get()
        RCStartY = self.PlaceY1_Entry.get()
        RCEndX = self.PlaceX2_Entry.get()
        RCEndY = self.PlaceY2_Entry.get()
        self.Canvas.coords("Square", RCStartX, RCStartY, RCEndX, RCEndY)


    def DVSelected(self, event = None):
        # 選択中のデータを使って保存されているデータを取得、表示。

        Data = self.DVDetailsGet(self.DVTree.set(self.DVTree.selection())["1"])[0]
        try:
            self.Canvas.delete("Square")
        except:
            pass
        self.Canvas.create_rectangle(Data[1], Data[2], Data[3], Data[4],outline="YELLOW", tag="Square", width=2)
    def DVDelete(self, event = None):
        Data = self.DVTree.set(self.DVTree.selection())["1"]
        if not PCUserName in OKUserID:
            messagebox.showerror("エラー", "削除の権限がありません。")
            return
        if messagebox.askyesno("変更確認", f"選択中のデータを削除しますか？{Data}"):
            if self.DVDeleteData(Data):
                messagebox.showinfo("完了", "削除しました。")
                self.RCShowImage()
            pass
    def DVDeleteData(self, Name):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"DELETE FROM NamePlaces WHERE Name = '{Name}'"
            cur.execute(sql)
            con.commit()
            con.close()
            return True
        except Exception as e:
            return e

    def DVEdit(self, event = None):
        global RCStartX, RCStartY, RCEndX, RCEndY
        Data = self.DVTree.set(self.DVTree.selection())['1']
        if messagebox.askyesno("変更確認", f"選択中のデータを変更しますか？\n{Data}"):
            # 変更するような画面を表示
            RCStartX, RCStartY, RCEndX, RCEndY = self.Canvas.bbox("Square")
            self.EndCreateSquare()
            # 登録ボタンの文字とメソッド変更
            self.RegistButton["text"] = "変更"
            self.RegistButton["command"] = lambda:[self.DVChange(Data)]

            self.PlaceNameEntry.insert(0, Data[len(self.RCImageName)+1:])
            pass
        else:
            return





    def DVChange(self, Name):
        # 下の関数をコピーして、データベースへの書き込みを変更にする。
        global ImagePlaceRegistWindow
        Data = []
        # 場所の名前(画像の名前)
        if len(self.PlaceNameEntry.get()) == 0 or re.match(r"\s+", self.PlaceNameEntry.get()):
            messagebox.showerror("エラー", "登録名称を入力してください")
            return
        Data.append(self.RCImageName)

        # 座標たち
        Data.append(self.PlaceX1_Entry.get())
        Data.append(self.PlaceY1_Entry.get())
        Data.append(self.PlaceX2_Entry.get())
        Data.append(self.PlaceY2_Entry.get())

        # 名前
        Data.append(self.RCImageName + "-" + self.PlaceNameEntry.get())

        # ここから下を書き換える
        if self.RCChangeDB(Name, Data):
            messagebox.showinfo("完了", f"変更完了しました。\n{self.RCImageName}-{self.PlaceNameEntry.get()}")
            self.RCShowImage()
            self.Canvas.delete("Square")
            ImagePlaceRegistWindow.destroy()
        else:
            pass
    
    def RCChangeDB(self, Name, Data):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"UPDATE NamePlaces SET X1 = '{Data[1]}', Y1 = '{Data[2]}', X2 = '{Data[3]}', Y2 = '{Data[4]}', Name = '{Data[5]}' WHERE Name = '{Name}'"
            cur.execute(sql)
            con.commit()
            con.close()
            return True
        except Exception as e:
            return e

    
    def CheckPlaceName(self, Value, SData):
        # Valueに選択されたエリアを。SDataに規準となるDBデータを。
        # X1 integer, Y1 integer, X2 integer, Y2 integer
        if SData[1] <= Value[0] <= SData[3] and SData[2] <= Value[1] <= SData[4]:
            return True
        else:
            return False

    def RCRegist(self):
        global ImagePlaceRegistWindow
        Data = []
        # 場所の名前(画像の名前)
        if len(self.PlaceNameEntry.get()) == 0 or re.match(r"\s+", self.PlaceNameEntry.get()):
            messagebox.showerror("エラー", "登録名称を入力してください")
            return
        Data.append(self.RCImageName)

        # 座標たち
        Data.append(self.PlaceX1_Entry.get())
        Data.append(self.PlaceY1_Entry.get())
        Data.append(self.PlaceX2_Entry.get())
        Data.append(self.PlaceY2_Entry.get())

        # 名前
        Data.append(self.RCImageName + "-" + self.PlaceNameEntry.get())
        if self.RCRegistToDB(Data):
            messagebox.showinfo("完了", f"登録完了しました。\n{self.RCImageName}-{self.PlaceNameEntry.get()}")
            self.RCShowImage()
            self.Canvas.delete("Square")
            ImagePlaceRegistWindow.destroy()
        # cur.execute("""CREATE TABLE NamePlaces (Image text, X1 integer, Y1 integer, X2 integer, Y2 integer, Name text)""")
    def DVGetData(self):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"Select Name from NamePlaces WHERE Image = '{self.RCImageName}'"
            cur.execute(sql)
            a = cur.fetchall()
            con.commit()
            con.close()
            return a
        except Exception as e:
            return e
    def DVGetAllData(self, Name):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"Select * from NamePlaces WHERE Image = '{Name}'"
            cur.execute(sql)
            a = cur.fetchall()
            con.commit()
            con.close()
            return a
        except Exception as e:
            return e
    def DVDetailsGet(self, Data):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"Select * from NamePlaces WHERE Name = '{Data}'"
            cur.execute(sql)
            a = cur.fetchall()
            con.commit()
            con.close()
            return a
        except Exception as e:
            return e
        
    def RCRegistToDB(self, Data):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"INSERT INTO NamePlaces values {tuple(Data)}"
            cur.execute(sql)
            con.commit()
            con.close()
            return True
        except Exception as e:
            con.close()
            messagebox.showerror("エラー", "名前が重複しています。\n違う名前を選択してください")
            return False

        
    def AddImageSelectPath(self, event = None):
        self.AddPathEntry.delete(0, "end")
        self.AddPathEntry.insert(0, filedialog.askopenfilename(filetypes = [("PNG", "*.png")]))
        subwindow.lift()
        
    def SelectPath(self, event = None):
        self.OutputPathEntry.delete(0, "end")
        self.OutputPathEntry.insert(0, filedialog.askdirectory())
        subwindow.lift()

    def ToCSVReal(self):
        try:
            if not messagebox.askyesno("確認", "出力しますか？"):
                return
            if not os.path.exists(self.OutputPathEntry.get()):
                messagebox.showerror("エラー", "パスが存在しません。")
                return 
            a = self.GetAllData()
            for index, i in enumerate(a):
                self.placedata = self.DVGetAllData(i[1][:-4])
                if not len(self.placedata) == 0:
                    for l in self.placedata:
                        if self.CheckPlaceName([int(i[3]), int(i[4])], l):
                            a[index] = list(a[index])
                            a[index][5] += " 登録された位置名:" + l[5]
                            a[index] = tuple(a[index])
                            break
            print(a)

            with open(f"{self.OutputPathEntry.get()}/{str(datetime.date.today())}.csv", "w", newline="", encoding="shift-jis") as f:
                writer = csv.writer(f)
                writer.writerow(["内部管理ID", "場所", "登録番号", "X座標", "Y座標", "詳細場所"])
                writer.writerows(a)
            messagebox.showinfo("Success", f"Success\nPath:{self.OutputPathEntry.get()}/{str(datetime.date.today())}.csv")
            SaveLogs("ログ保存", f"{a[0][2]}", "ForAdminLog")
            subwindow.lift()
            return 
        except Exception as e:
            messagebox.showerror("エラー", f"失敗しました。\nエラー内容→{e}")
            subwindow.lift()
            return 
    def GetAllData(self):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"Select * from Data"
            cur.execute(sql)
            a = cur.fetchall()
            con.commit()
            con.close()
            return a
        except Exception as e:
            return e
    def GetNumberData(self):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"Select Number from Data"
            cur.execute(sql)
            a = cur.fetchall()
            con.commit()
            con.close()
            return a
        except Exception as e:
            return e
    def ShowPlace(self):
        global subwindow 
        self.LogData = []
        self.Canvas = None
        self.ImageFrame = None
        self.TF1 = BooleanVar(value=True)
        self.TF2 = BooleanVar(value=True)
        if subwindow == None or not subwindow.winfo_exists():
            subwindow = Toplevel(root)
        else:
            return
        subwindow.title("所在地確認・変更")
        subwindow.bind("<Control-d>", self.Delete)
        self.DataEntryLabel = ttk.Labelframe(subwindow, text="登録番号入力")
        self.DataEntryLabel.pack(side="left", anchor=N)
        self.DataEntryBox = ttk.Combobox(self.DataEntryLabel)
        self.DataEntryBox["values"] = self.GetNumberData()
        self.DataEntryBox.pack()
        self.DataEntryBox.bind("<Return>", self.ShowPlaceInfo)
        self.SubmitButton = ttk.Button(self.DataEntryLabel, text="表示", command=lambda:[self.ShowPlaceInfo()])
        self.SubmitButton.pack()

        self.ImageCombo = ttk.Combobox(self.DataEntryLabel)
        path = "files\image"
        files = os.listdir(path)
        files_file = [f[0:-4] for f in files if os.path.isfile(os.path.join(path, f)) and f[-4:].lower() == ".png"]
        self.ImageCombo["values"] = files_file

        AskConfirm = ttk.Checkbutton(self.DataEntryLabel, text="確認メッセージ表示", variable=self.TF1)
        AskConfirm.pack(anchor=W, pady=5)
        AskError = ttk.Checkbutton(self.DataEntryLabel, text="エラーメッセージ表示", variable=self.TF2)
        AskError.pack(anchor=W, pady=5)
        LogFrame = ttk.Labelframe(self.DataEntryLabel, text="Log")
        LogFrame.pack(anchor=W, pady=5)
        self.Logs = Listbox(LogFrame, width=40)
        self.Logs.pack(anchor=W, pady=5)
        self.deletebutton = ttk.Button(self.DataEntryLabel, text="[D]削除", command=lambda:[self.Delete()])
        
        self.ChangeImageButton = ttk.Button(self.DataEntryLabel, text="画像変更", command=lambda:[self.ShowedImageChange()])

        self.PlaceDetailLabel = ttk.Labelframe(self.DataEntryLabel, text="詳細情報入力")
        self.PlaceDetailEntry = ttk.Entry(self.PlaceDetailLabel)
        self.PlaceDetailEntry.pack(anchor=W, pady=5)
    def ShowPlaceInfo(self, event = None):
        if self.DataEntryBox.get() == "":
            return
        a = self.GetData(self.DataEntryBox.get())
        try:
            self.deletebutton.pack_forget()
        except:
            pass
        try:
            self.PlaceDetailLabel.pack_forget()
        except:
            pass
        try:
            self.ImageCombo.pack_forget()
        except:
            pass
        try:
            self.ChangeImageButton.pack_forget()
        except:
            pass
        if a[0] == "NG":
            return
        if a[1] == []:
            if self.TF2.get():
                messagebox.showerror("登録なし", f"このデータは登録されていません。({self.DataEntryBox.get()})")
                subwindow.lift()
            self.showlogs("[Missed]"+self.DataEntryBox.get()+ "is not defined")
            self.DataEntryBox.delete(0, "end")
            return
        if self.ImageFrame != None:
            self.ImageFrame.pack_forget()
        self.ImageFrame = ttk.Labelframe(subwindow, text="Map")
        self.ImageFrame.pack(side="left")
        imagepath = "files/image/" + a[1][0][1]
        if not os.path.exists(imagepath):
            messagebox.showerror("エラー", f"登録時に使用した画像が存在していません。\n登録時の画像名 : {imagepath}")
            if messagebox.askyesno("確認", "データを削除しますか？"):
                self.ImageFrame["text"] = "Map-" + a[1][0][2]
                self.ChangeBeforeImagename = a[1][0][1]
                self.ChangeBeforePlace = f"{a[1][0][3]}, {a[1][0][4]}"
                f"最終場所:{self.ChangeBeforeImagename}({self.ChangeBeforePlace})"
                self.Delete()
            subwindow.lift()
            return
        img = Image.open(imagepath)
        width, height = img.size
        if self.Canvas != None:
            self.Canvas.pack_forget()
            self.Canvas.delete()
        
        self.placedata = self.DVGetAllData(a[1][0][1][:-4])
        placename = f"x = {a[1][0][3]}, y = {a[1][0][4]}"
        if not len(self.placedata) == 0:
            for i in self.placedata:
                if self.CheckPlaceName([int(a[1][0][3]), int(a[1][0][4])], i):
                    placename = i[5]
                    break
        self.imagename = a[1][0][1]
        self.ChangeBeforeImagename = self.imagename
        self.ChangeBeforePlace = placename
        self.Canvas = Canvas(self.ImageFrame, height=height, width=width)
        self.Canvas.pack(anchor=N)
        self.photo_image = PhotoImage(file = imagepath)
        self.Canvas.create_image(width/2, height/2, image=self.photo_image, tag="image")
        self.Canvas.create_text(a[1][0][3], a[1][0][4], text=f"← ({placename})", font=("HG丸ｺﾞｼｯｸM-PRO",12), anchor="w", tags="Point")
        self.Canvas.tag_bind("image","<1>", self.Change)
        self.ImageFrame["text"] = "Map-" + a[1][0][2]
        self.showlogs("[Show]"+self.DataEntryBox.get())
        self.DataEntryBox.delete(0, "end")
        self.ImageCombo.pack(padx=5, pady=5)
        self.ChangeImageButton.pack(padx=5, pady=5)
        self.deletebutton.pack(padx=5, pady=5)
        self.PlaceDetailLabel.pack()
        self.PlaceDetailEntry.delete(0, "end")
        self.PlaceDetailEntry.insert(0, a[1][0][5])
    def ShowedImageChange(self):
        imagepath = "files/image/" + self.ImageCombo.get() + ".png"
        if not os.path.isfile(imagepath):
            return
        img = Image.open(imagepath)
        width, height = img.size
        self.photo_image = PhotoImage(file = imagepath)
        self.Canvas.delete("image")
        self.Canvas.create_image(width/2, height/2, image=self.photo_image, tag="image")
        self.Canvas.tag_bind("image","<1>", self.Change)
        self.imagename = self.ImageCombo.get() + ".png"
        self.placedata = self.DVGetAllData(self.ImageCombo.get())

    def Change(self, event):
        if self.TF1.get():
            if not messagebox.askyesno("確認", "クリックした位置に変更します。よろしいですか？\n※詳細情報も変更されます。"):
                return
        placename = f"x = {event.x}, y = {event.y}"
        if not len(self.placedata) == 0:
            for i in self.placedata:
                if self.CheckPlaceName([event.x, event.y], i):
                    placename = i[5]
                    break
        # self.Canvas.moveto("Point", event.x, event.y)
        self.Canvas.delete("Point")
        self.Canvas.create_text(event.x, event.y, text=f"← ({placename})", font=("HG丸ｺﾞｼｯｸM-PRO",12), anchor="w", tags="Point")
        
        self.ypoint = event.y
        self.xpoint = event.x
        if self.TF1.get():
            messagebox.showinfo("完了", "完了しました。")
            subwindow.lift()
        self.ChangeData(self.ypoint, self.xpoint, self.imagename)
        self.showlogs("[Changed]Point was Changed")
        SaveLogs("移動", f"{self.ChangeBeforeImagename[:-4]}({self.ChangeBeforePlace})>>>{self.imagename[:-4]}({placename})", self.ImageFrame["text"][4:])
        self.ChangeBeforePlace = placename
        return
    def Delete(self, event=None):
        global OKUserID, PCUserName
        if not self.deletebutton.winfo_exists():
            return
        if not PCUserName in OKUserID:
            messagebox.showerror("エラー", "削除の権限がありません。")
            return
                
        # ユーザー認証が必要。

        if messagebox.askyesno("確認", "現在表示中のデータを削除します。よろしいですか？"):
            self.DeleteData()
            self.Canvas.pack_forget()
            self.deletebutton.pack_forget()
            self.PlaceDetailLabel.pack_forget()
            self.DataEntryBox["values"] = self.GetNumberData()
            messagebox.showinfo("完了", "完了しました。")
            SaveLogs("削除", f"最終場所:{self.ChangeBeforeImagename}({self.ChangeBeforePlace})", self.ImageFrame["text"][4:])
            subwindow.lift()
            return

    def DeleteData(self):
        try:
            Number = self.ImageFrame["text"][4:]
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"DELETE FROM Data WHERE Number = '{Number}'"
            cur.execute(sql)
            con.commit()
            con.close()
            self.showlogs(f"[deleted]{Number}")
            return "OK"
        except Exception as e:
            self.showlogs(f"[missed]{Number}, {e}")
            return e
        
    def ChangeData(self, datax, datay, imagename):
        try:
            Number = self.ImageFrame["text"][4:]
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"UPDATE DATA SET X = '{datay}', Y = '{datax}', Detail = '{self.PlaceDetailEntry.get()}', Image = '{imagename}' WHERE Number = '{Number}'"
            cur.execute(sql)
            con.commit()
            con.close()
            return "OK"
        except Exception as e:
            return e
    def GetData(self, Number):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"Select * from Data where Number = '{Number}' "
            cur.execute(sql)
            a = cur.fetchall()
            con.commit()
            con.close()
            return ["OK", a]
        except Exception as e:
            return ["NG", e]

    def EntryPlace(self):
        global a
        a = None
        global subwindow
        global root
        if subwindow == None or not subwindow.winfo_exists():
            subwindow = Toplevel(root)
        else:
            return
        subwindow.title("場所入力")
        self.TF1 = BooleanVar(value=True)
        self.TF2 = BooleanVar(value=True)

        self.xpoint = None
        self.ypoint = None
        self.Canvas = None
        self.LogData = []
        self.imgframe = ttk.Labelframe(subwindow, text="地図")
        self.imgframe.pack(side="right", anchor=N)
        infoframe = ttk.Labelframe(subwindow, text="Info", width=100)
        infoframe.pack(side="right", anchor=N)
        ImageSelectLabel = ttk.Labelframe(infoframe, text="地図選択")
        ImageSelectLabel.pack(anchor=W, pady=5)
        self.ImageCombo = ttk.Combobox(ImageSelectLabel)
        path = "files\image"
        files = os.listdir(path)
        files_file = [f[0:-4] for f in files if os.path.isfile(os.path.join(path, f)) and f[-4:].lower() == ".png"]
        self.ImageCombo["values"] = files_file
        self.ImageCombo.pack(anchor=W, pady=5)
        self.ImageCombo.bind("<Return>", self.ShowImage)
        ImageSelectButton = ttk.Button(infoframe, text="選択", command=lambda:[self.ShowImage(0)])
        ImageSelectButton.pack(anchor=W, pady=5)
        NumberEntryLabel = ttk.Labelframe(infoframe, text="登録番号入力")
        NumberEntryLabel.pack(anchor=W, pady=5)
        self.NumberEntryBox = ttk.Entry(NumberEntryLabel)
        self.NumberEntryBox.bind("<Return>", self.RegistPlace)
        self.NumberEntryBox.pack(anchor=W, pady=5)
        PlaceDetailInfoLabel = ttk.Labelframe(infoframe, text="詳細(なしでも可)")
        PlaceDetailInfoLabel.pack(anchor=W, pady=5)
        self.PlaceDetailEntry = ttk.Entry(PlaceDetailInfoLabel)
        self.PlaceDetailEntry.pack(anchor=W, pady=5)
        RegistButton = ttk.Button(infoframe, text="登録", command=lambda:[self.RegistPlace()])
        RegistButton.pack(anchor=W, pady=5)
        AskConfirm = ttk.Checkbutton(infoframe, text="確認メッセージ表示", variable=self.TF1)
        AskConfirm.pack(anchor=W, pady=5)
        AskError = ttk.Checkbutton(infoframe, text="エラーメッセージ表示", variable=self.TF2)
        AskError.pack(anchor=W, pady=5)
        LogFrame = ttk.Labelframe(infoframe, text="Log")
        LogFrame.pack(anchor=W, pady=5)
        self.Logs = Listbox(LogFrame, width=40)
        self.Logs.pack(anchor=W, pady=5)
        self.placeinfo = ttk.Label(infoframe, text="")
        self.placeinfo.pack(anchor=W, pady=5)
        # self.ShowImage("k-s.png")
    def ShowImage(self, event, path = None):
        if path == None:
            path = self.ImageCombo.get() + ".png"
        if path == "":
            return
        if not os.path.exists("files/image/"+path):
            return
        global a
        global subwindow
        img = Image.open("files/image/"+path)
        width, height = img.size
        if not self.Canvas == None:
            self.Canvas.pack_forget()
            self.Canvas.delete()
        self.Canvas = Canvas(self.imgframe, height=height, width=width)
        self.Canvas.pack()



        self.photo_image = PhotoImage(file = "files/image/" + path)
        self.imagename = self.ImageCombo.get()

        self.Canvas.create_image(width/2, height/2, image=self.photo_image, tag="image")
        self.Canvas.tag_bind("image", "<1>", self.ImagePoint)
        a = None
        self.xpoint = None
        self.ypoint = None
        self.showlogs(f"[Showed] image/{path}")


    def ImagePoint(self, event):
        global a
        placedata = self.DVGetAllData(self.imagename)
        placename = f"x = {event.x}, y = {event.y}"
        if not len(placedata) == 0:
            for i in placedata:
                if self.CheckPlaceName([event.x, event.y], i):
                    placename = i[5]
                    break
        self.RegistPlaceData = placename
        
        if not self.TF1.get():
            if a == None:
                a = self.Canvas.create_text(event.x, event.y, text=f"← ({placename})", font=("HG丸ｺﾞｼｯｸM-PRO",12), anchor="w", tags="Point")
            else:
                self.Canvas.delete("Point")
                self.Canvas.create_text(event.x, event.y, text=f"← ({placename})", font=("HG丸ｺﾞｼｯｸM-PRO",12), anchor="w", tags="Point")
            self.ypoint = event.y
            self.xpoint = event.x
            self.placeinfo["text"] = f"x = {self.xpoint}, y = {self.ypoint}"
            self.showlogs("[Changed]Point was Changed")
            return
        elif a != None and messagebox.askyesno("確認", "クリックした位置に変更します。よろしいですか？"):
            self.Canvas.delete("Point")
            self.Canvas.create_text(event.x, event.y, text=f"← ({placename})", font=("HG丸ｺﾞｼｯｸM-PRO",12), anchor="w", tags="Point")
            self.ypoint = event.y
            self.xpoint = event.x
            self.placeinfo["text"] = f"x = {self.xpoint}, y = {self.ypoint}"
            messagebox.showinfo("完了", "完了しました。")
            subwindow.lift()
            self.showlogs("[Changed]Point was Changed")
            return
        elif a == None and messagebox.askyesno("確認", "登録位置をクリックした位置で確定します。よろしいですか？"):
            a = self.Canvas.create_text(event.x, event.y, text=f"← ({placename})", font=("HG丸ｺﾞｼｯｸM-PRO",12), anchor="w", tags="Point")
            self.ypoint = event.y
            self.xpoint = event.x
            self.placeinfo["text"] = f"x = {self.xpoint}, y = {self.ypoint}"
            messagebox.showinfo("完了", "完了しました。")
            subwindow.lift()
            self.showlogs("[Changed]Point was Changed")
            return
    def RegistPlace(self, event = None):
        if self.NumberEntryBox.get() == "":
            if self.TF1.get():
                messagebox.showinfo("エラー", "番号が入力されていません。")
            self.showlogs("[empty] Number form is empty")
            return
        if self.xpoint == None or self.ypoint == None:
            if self.TF2.get():
                messagebox.showerror("エラー", "地図上の登録位置が設定されていません。")
            self.showlogs("[error]" + self.NumberEntryBox.get() + " Coodinate is not selected")
            self.NumberEntryBox.delete(0, "end")
            self.PlaceDetailEntry.delete(0, "end")
            return
        time = str(datetime.datetime.now().year)+str(datetime.datetime.now().month)+str(datetime.datetime.now().day)+str(datetime.datetime.now().hour)+str(datetime.datetime.now().minute)+str(datetime.datetime.now().second)+str(datetime.datetime.now().microsecond)
        timehash = hash(time)
        PreparedData = []
        PreparedData.append(timehash)
        PreparedData.append(self.imagename+".png")
        PreparedData.append(self.NumberEntryBox.get().replace("[", "(").replace("]", ")"))
        PreparedData.append(self.xpoint)
        PreparedData.append(self.ypoint)
        PreparedData.append(self.PlaceDetailEntry.get().replace("[", "(").replace("]", ")"))
        if self.TF1.get():
            if not messagebox.askyesno("確認", "登録しますか？"):
                self.showlogs(f"[Canceled]{self.NumberEntryBox.get()} is canceled to entry")
                return
        result = self.RegistToDB(PreparedData)
        if result[0] == "NG":
            if self.TF2.get():
                messagebox.showerror("エラー", "データが重複しています。\n(登録済みデータ)")
                subwindow.lift()
            self.showlogs("[Duplicate]" + str(result[1]))
        else:
            self.showlogs("[Registed]" + self.NumberEntryBox.get())
            SaveLogs("登録", f"{self.imagename}({self.RegistPlaceData})", self.NumberEntryBox.get())
        self.NumberEntryBox.delete(0, "end")
        self.PlaceDetailEntry.delete(0, "end")
    def showlogs(self, data):
        if len(self.LogData) == 10:
            New = []
            New.append(data)
            for i in (0, 1, 2, 3, 4, 5, 6, 7, 8):
                New.append(self.LogData[i])
            self.LogData = New
        else:
            self.LogData.insert(0, data)
        self.Logs.delete(0, 9)
        for index, i in enumerate(self.LogData):
            self.Logs.insert(index, i)
    def RegistToDB(self, Data):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"INSERT INTO Data values {tuple(Data)}"
            cur.execute(sql)
            con.commit()
            con.close()
            return ["OK", None]
        except Exception as e:
            con.close()
            print(e)
            return ["NG", e]




def startmain(event = None):
    global UserName, PCUserName
    UserName = UsernameEntry.get()
    if not re.match(r"^\s*$", UserName):
        PCUserName = os.getlogin()
        userinfo.destroy()
        SaveLogs("開始", "", "ForAdminLog")
        main()
if __name__ == "__main__":
    userinfo = Tk()
    userinfo.title("ユーザー情報入力")
    UsernameLabel = ttk.Label(userinfo, text="ユーザー名入力")
    UsernameEntry = ttk.Entry(userinfo)
    UsernameEntry.bind("<Return>", startmain)
    LoginButton = ttk.Button(userinfo, text="開始", command=lambda:[startmain()])
    UsernameLabel.grid(row=0, column=0, padx=15, pady=5)
    UsernameEntry.grid(row=0, column=1, padx=15, pady=5)
    LoginButton.grid(row=1, column=0, padx=5, pady=5, columnspan=2)
    userinfo.mainloop()



    pass