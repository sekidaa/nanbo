# Import the required libraries
from tkinter import *
import traceback
from tkinter import messagebox
import tkinter as tk
import openpyxl
from openpyxl import Workbook, load_workbook
import os
import webbrowser
# Create an instance of tkinter frame or window
win=Tk()

# Set the size of the tkinter window
win.geometry("550x500")
win.title("毎月なんぼ?アプリ")
win.configure(bg="#fffacd")

def cal_sum():
   try:
      t1=int(a.get())
      t2=int(b.get())
      p1= t1 + t2
      sum=round(p1*0.24, 3)
      sum2=round(p1*0.12, 3)
      sum3=round(p1*0.08, 3)
      sum4=round(p1*0.05, 3)
      sum5=round(p1*0.04, 3)
      label1.config(text=str("・食費(24%)：")+str(sum)+str("万円"))
      label2.config(text=str("・貯蓄(12%)：")+str(sum2)+str("万円"))
      label3.config(text=str("・娯楽費(8%)：")+str(sum3)+str("万円"))
      label4.config(text=str("・交際費(5%)：")+str(sum4)+str("万円"))
      label5.config(text=str("・被服費(4%)：")+str(sum5)+str("万円"))

   except  Exception as e:
      traceback.print_exc()
      messagebox.showerror("値入力エラー", str("月収または副収入に数値を入力してください。"))

def excel_enter():
   try:
      t1=int(a.get())
      t2=int(b.get())
      p1= t1 + t2
      wb = openpyxl.Workbook()  #新規ワークブックを作成
      sheet = wb.active
      sheet.append(['月収', str(t1)+r"万円"])
      sheet.append(['副収入', str(t2)+r"万円"])
      sheet.append(['食費(24%)', str(round(p1*0.24, 3))+r"万円"])
      sheet.append(['貯蓄(12%)', str(round(p1*0.12, 3))+r"万円"])
      sheet.append(['娯楽費(8%)', str(round(p1*0.08, 3))+r"万円"])
      sheet.append(['交際費(5%)', str(round(p1*0.05, 3))+r"万円"])
      sheet.append(['被服費(4%)', str(round(p1*0.04, 3))+r"万円"])
      # ログイン名の取得
      path_a = r'C:\Users\\'
      user = os.environ['USERNAME']
      path_b = r"\Documents\shuturyoku.xlsx"
      # 取得したログイン名を表示
      path = path_a + user + path_b
      wb.save(path)
      
   except  Exception as e:
      traceback.print_exc()
      messagebox.showerror("Excelエラー", str("Excelファイルへの出力に失敗しました。"))

def web_enter():
   url = "https://docs.google.com/forms/d/e/1FAIpQLSf6tRMWe_dwTBG_h7U2X2jOAPqTMx7QqZ8ihUwikrCKlwqyWA/viewform?usp=sf_link"
   webbrowser.open(url)

def version_look():
   messagebox.showinfo("バージョン情報", str("Ver.1.0.0.0"))

label=Label(win, text="月収と副収入をそれぞれ万単位で入力してください。", font=('Calibri 10'), bg = '#fffacd')
label.pack(pady=10)

# 月収入力欄
Label(win, text=r"月収", font=('Calibri 10'), bg = '#fffacd').pack()
a=Entry(win, width=15)
a.insert(0, "0")
a.pack()

# 副収入入力欄
Label(win, text=r"副収入", font=('Calibri 10'), bg = '#fffacd').pack()
b=Entry(win, width=15)
b.insert(0, "0")
b.pack()


label=Label(win, text="なんぼ?を押すと、平均的な成人(独身)の出費のうち以下5項目の目安が出力されます。", font=('Calibri 10'), bg = '#fffacd')
label.pack(pady=10)
label1=Label(win, text="・食費(24%)", font=('Calibri 12'), bg = '#fffacd')
label1.pack(pady=10)

label2=Label(win, text="・貯蓄(12%)", font=('Calibri 12'), bg = '#fffacd')
label2.pack(pady=10)

label3=Label(win, text="・娯楽費(8%)", font=('Calibri 12'), bg = '#fffacd')
label3.pack(pady=10)

label4=Label(win, text="・交際費(5%)", font=('Calibri 12'), bg = '#fffacd')
label4.pack(pady=10)

label5=Label(win, text="・被服費(4%)", font=('Calibri 12'), bg = '#fffacd')
label5.pack(pady=10)
menubar = tk.Menu(win)

filemenu = tk.Menu(menubar, tearoff=0)
filemenu.add_command(label="ドキュメントフォルダーにExcel出力", command=excel_enter)
filemenu.add_command(label="お問い合わせ", command=web_enter)
filemenu.add_command(label="バージョン情報", command=version_look)
menubar.add_cascade(label="　メニュー　", menu=filemenu)

win.config(menu=menubar)
Button(win, text="なんぼ?", font='Calibri 15', command=cal_sum).pack()

win.mainloop()
