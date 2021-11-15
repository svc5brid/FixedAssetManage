import sqlite3
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import datetime
import os
from tkinter import messagebox



DBName = None

class main():
    def __init__(self) -> None:
        global DBName
        root = Tk()
        root.title("削除可能ユーザー追加")

        UserLabel = ttk.Label(root, text="登録するIDを入力")
        self.UserEntry = ttk.Entry(root)

        RegistButton = ttk.Button(root, text="登録", command=lambda:[self.Registdata()])

        self.Tree = ttk.Treeview(root)
        self.Tree["columns"] = (1)
        self.Tree["show"] = "headings"
        self.Tree.heading(1, text="登録済み番号")
        self.Tree.column(1, width=220)
        UserLabel.pack(padx=5, pady=5)
        self.UserEntry.pack(padx=5, pady=5)
        RegistButton.pack(padx=5, pady=5)
        self.Tree.pack(padx=5, pady=5)
        DBName = filedialog.askopenfilename(filetypes=[("DB", "*.db")])
        root.focus_force()
        a = self.GetUserData()
        for i in a:
            print(i)
            self.Tree.insert("", "end", values=i[0])
        self.Tree.bind("<Delete>", self.Deletedata)
        root.mainloop()
    def Registdata(self):
        if len(self.UserEntry.get()) == 0:
            return
        Data = (self.UserEntry.get(), str(datetime.datetime.now()), os.getlogin())
        if self.RegistUser(Data):
            messagebox.showinfo("完了", "登録完了しました。")
            for i in self.Tree.get_children():
                self.Tree.delete(i)
            a = self.GetUserData()
            for i in a:
                self.Tree.insert("", "end", values = i[0])
            self.UserEntry.delete(0, END)
    def Deletedata(self, event):
        if messagebox.askyesno("確認", "削除します。よろしいですか？"):
            if self.DeleteUserData():
                messagebox.showinfo("完了", "削除しました。")
                for i in self.Tree.get_children():
                    self.Tree.delete(i)
                a = self.GetUserData()
                for i in a:

                    self.Tree.insert("", "end", values = i[0])

        
    def GetUserData(self):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"Select * from Users"
            cur.execute(sql)
            a = cur.fetchall()
            con.commit()
            con.close()
            return a
        except Exception as e:
            return e

    def RegistUser(self, Data):
        try:
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"INSERT INTO Users VALUES {Data}"
            cur.execute(sql)
            con.commit()
            con.close()
            return True
        except Exception as e:
            return False
    
    def DeleteUserData(self):
        try:
            GlobalID = self.Tree.set(self.Tree.selection())["1"]
            con = sqlite3.connect(DBName)
            cur = con.cursor()
            sql = f"DELETE FROM Users WHERE GID = '{GlobalID}'"
            cur.execute(sql)
            con.commit()
            con.close()
            return True
        except Exception as e:
            return False




if __name__ == "__main__":
    main()