#!/usr/bin/env python
# coding: utf-8

# In[1]:


from tkinter import *
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import tkinter, os, requests, re, csv, datetime, glob, collections, itertools, time, random, json
import codecs, mojimoji, phonenumbers, threading, traceback, subprocess, smtplib
from pathlib import Path
from tkinter import filedialog, messagebox
from functools import partial
from concurrent import futures
from concurrent.futures.thread import ThreadPoolExecutor
from tkinter.scrolledtext import ScrolledText
from email.mime.text import MIMEText
from email.utils import formatdate
from urllib.parse import urlparse
from PIL import ImageTk, Image
from gspread_dataframe import set_with_dataframe
from pyfile import main_actions, error_actions, config


def send_log(txt_shousai, body):
    log_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    txt_shousai.insert(tkinter.END, f"[{log_time}]：{body}")
    txt_shousai.see("end")

def Maintotal(num_start,num_end,sheet_name, making_type):
    b0_4['state']= DISABLED        
    #定義
    df_sheet = pd.read_excel(config.EXCEL_PATH, sheet_name = sheet_name)
    df_sheet_xpath = pd.read_excel(config.EXCEL_PATH, sheet_name = "xpath管理")
    df = pd.merge(df_sheet, df_sheet_xpath, on='媒体名')
    df = df.iloc[num_start-2:num_end-1]
    dicts = df.to_dict(orient="record")
    if dicts:
        error = error_actions.main(num_start, dicts)
        if not error:
            num = num_start -1
            for dic in dicts: 
                num += 1
                try:
                    medium = dic["媒体名"]
                    save_path_folder = config.SAVE_EXCEL_PATH
                    log_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
                    txt_shousai.delete(0.0, tkinter.END)
                    txt_shousai.insert(tkinter.END, f"[{log_time}]：{medium}開始\n")
                    txt_shousai.see("end")     
                    log_time = datetime.datetime.now().strftime("%Y%m%d")
                    save_filename = f"{save_path_folder}/{log_time}【{medium}】.xlsx"
                    columns = list(dic.keys())[4:]
                    columns.insert(0, 'URL')
                    #メイン処理
                    #ページネーション
                    item_urls = main_actions.get_item_url(dic, txt_shousai, making_type)
                    if item_urls:
                        log_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
                        txt_shousai.insert(tkinter.END, f"[{log_time}]：アイテムループ\n")
                        txt_shousai.see("end")
                        #アイテムループ
                        item_contents = main_actions.get_responce_item(txt_shousai, item_urls)
                        if item_contents:
                            item_contents = [main_actions.main(item_content, dic) for item_content in item_contents]
                            df_out = pd.DataFrame(data = item_contents, columns = columns)
                            df_out = df_out.dropna(axis=1, how='all')
                            df_out = df_out.replace('\n', '', regex = True).replace('\t', '', regex = True)
                            for columns in df_out.columns: 
                                df_out[columns] = df_out[columns].str.strip()
                            df_out.to_excel(save_filename, index = False)
                            send_log(txt_shousai, "終了")
                            b0_4['state']= NORMAL
                        else:
                            b0_4['state']= NORMAL
                            send_log(txt_shousai, "アイテムループエラー\n【原因】：xpath違い・リクエストエラー")
                            continue
                    else:
                        b0_4['state']= NORMAL
                        send_log(txt_shousai, "ページネーションエラー\n【原因】：xpath違い・リクエストエラー")
                        continue
                except:
                    txt_shousai.insert(tkinter.END, traceback.format_exc() + '\n\n格納できなかったよ')
                    txt_shousai.see("end")

                finally:
                    b0_4['state']= NORMAL
        else:
            b0_4['state']= NORMAL
            txt_shousai.insert(tkinter.END,  "\n".join(error))
            txt_shousai.see("end")
            
    else:
        b0_4['state']= NORMAL
        txt_shousai.insert(tkinter.END,  "媒体名あっているか確認してね。")
        txt_shousai.see("end")

def change_app(window):
    window.tkraise()
    
def switch():
    global btnState
    if btnState is True:
        for x in range(100):
            frame1.place(x=-x*3, y=0)
            frame0.update()
        btnState = False
    else:
        for x in range(-100, 0):
            frame1.place(x=x*3, y=0)
            frame0.update()
        btnState = True
        
def make_combo(frame, file_kinds, x, y,width,height):
    combo = ttk.Combobox(frame,state = "readonly", width = 10, justify = "center", font=("HG丸ｺﾞｼｯｸM-PRO", 12, "bold"))
    combo["values"] = file_kinds
    combo.current("0")
    combo.place(x = x, y = y,width = width,height = height)
    return combo

# ★バグ対応用の関数を追加
def fixed_map(option):
    return [elm for elm in style.map('Treeview', query_opt=option) if
            elm[:2] != ('!disabled', '!selected')]

def btn_click(path_list):
    sheet_name,num_start, num_end, making_type = combo.get(), entry1.get(), entry2.get(), combo2.get()
    threading.Thread(target = Maintotal, args=(int(num_start), int(num_end),sheet_name, making_type)).start()

def click_close():
    if messagebox.askokcancel("確認", "終了しますか？"):
        window.destroy()
        
path_list = []
window = Tk()
window.geometry("1000x600")
window.configure(bg = "#ffffff")
window.title("")
window.grid_rowconfigure(0, weight=1)
window.grid_columnconfigure(0, weight=1)
window.iconbitmap(default='asset/DashBord/favicon.ico')
window.option_add("*TCombobox*Listbox*Background", "gray93")
window.option_add("*TCombobox*Listbox*Font", "HG丸ｺﾞｼｯｸM-PRO")
window.option_add("*TCombobox*Listbox*selectBackground", "#000000")   # 背景色
window.option_add("*TCombobox*Listbox*foreground", "gray53")   # 文字色


frame0  = ttk.Frame(window) #base
frame3  = ttk.Frame(window) #about

frame0.grid( row=0, column=0, sticky="nsew", pady=0)
frame3.grid( row=0, column=0, sticky="nsew", pady=0)

style = ttk.Style()
style.theme_create('combostyle23', parent='alt',
                   settings = {'TCombobox':{'configure':{'selectbackground': 'gray93',
                                                         'fieldbackground': 'gray93',
                                                         'background': 'gray93',
                                                         "foreground":'gray53',
                                                        "selectforeground":'gray53',
                                                        "bordercolor":[("readonly", "!focus", "white"),("readonly", "focus", "white")]
                                                        }}})
style.theme_use('combostyle23') 



canvas  = Canvas(frame0,bg = "#ffffff",height = 600,width = 1000,bd = 0,highlightthickness = 0,relief = "ridge")
canvas3 = Canvas(frame3,bg = "#ffffff",height = 600,width = 1000,bd = 0,highlightthickness = 0,relief = "ridge")

canvas.place(x = 0, y = 0)
canvas3.place(x = 0, y = 0)


canvas.create_text(380, 550.0,text = "↓入力するExcelシート",fill = "#000000",font = ("Montserrat-Bold", int(10.0)))
canvas.create_text(650, 550.0,text = "↓出力されるExcelシート",fill = "#000000",font = ("Montserrat-Bold", int(10.0)))
canvas.create_rectangle(0, 0, 0+1000, 0+101,fill = "#000000",outline = "")
canvas3.create_rectangle(40, 14, 40+269, 14+1,fill = "#efefef",outline = "")
canvas3.create_rectangle(84, 240, 84+113, 240+5,fill = "#ffffff",outline = "")
canvas3.create_text(196.5, 59.0,text = "Created By",fill = "#000000",font = ("Montserrat-Bold", int(26.0)))
#ほうこく君
image = Image.open("asset/DashBord/houkokukunn.png")
photo = ImageTk.PhotoImage(image, master=window)
canvas.create_image(920, 127, image=photo)

style.configure("Treeview.Heading", font=("HG丸ｺﾞｼｯｸM-PRO", 12), background='#B8C8E8')
style.configure("Treeview", font=("HG丸ｺﾞｼｯｸM-PRO", 10), bordercolor="white", background='gray98', foreground='gray50')
style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))
style.configure("Horizontal.TScrollbar", gripcount=0,
                background="#000000", darkcolor="white", lightcolor="white",
                troughcolor="gray93", bordercolor="white", arrowcolor="white")


img0_1  = PhotoImage(file = f"asset/DashBord/img1.png",  master=frame0) #From
img0_2  = PhotoImage(file = f"asset/DashBord/img2.png",  master=frame0) #To
img0_3  = PhotoImage(file = f"asset/DashBord/img3.png",  master=frame0) #Homeロゴ
img0_4  = PhotoImage(file = f"asset/DashBord/img4.png",  master=frame0) #Action
img0_5  = PhotoImage(file = f"asset/DashBord/img5.png",  master=frame0) #type

img0_6  = PhotoImage(file = f"asset/DashBord/img6.png",  master=frame0) #sheet
img0_7  = PhotoImage(file = f"asset/DashBord/img7.png",  master=frame0) #
img0_8  = PhotoImage(file = f"asset/DashBord/img8.png",  master=frame0) #menuvar

img3_0 = PhotoImage(file = f"asset/about/img0.png",  master=frame3) #戻るボタン
img3_1 = PhotoImage(file = f"asset/about/img1.png",  master=frame3) #紹介

Labe0_1 = tk.Label(frame0, image=img0_1, background ="white")
Labe0_2 = tk.Label(frame0, image=img0_2, background ="white")
Labe0_3 = tk.Label(frame0, image=img0_3, background ="#000000")
Labe0_5 = tk.Label(frame0, image=img0_5, background ="white")
Labe0_6 = tk.Label(frame0, image=img0_6, background ="white")
Labe0_7 = tk.Label(frame0, image=img0_7, background ="white")
Labe3_1 = tk.Label(frame3, image=img3_1, background ="white")


b0_4  = Button(frame0, image = img0_4,borderwidth = 0,highlightthickness = 0,relief = "flat",bg = "white",cursor = "hand2")
b0_8  = Button(frame0, image = img0_8,borderwidth = 0,highlightthickness = 0,relief = "flat",bg = "#000000",cursor = "hand2", activebackground="gray93")
b3_0  = Button(frame3, image = img3_0,borderwidth = 0,highlightthickness = 0,relief = "flat",bg = "white",cursor = "hand2")

Labe0_1.place(x = 33, y = 257,width = 125,height = 62)
Labe0_2.place(x = 199, y = 257,width = 125,height = 62)
Labe0_3.place(x = 785, y = 32,width = 184,height = 40)
Labe0_5.place(x = 200, y = 157,width = 125,height = 62)
Labe0_6.place(x = 33, y = 157,width = 125,height = 62)
Labe0_7.place(x = 340, y = 138,width = 650,height = 390)

Labe3_1.place(x = 40, y = 120,width = 424,height = 289)

b0_4.place(x = 80, y = 450,width = 190,height = 48)
b0_8.place(x = 42, y = 42,width = 34,height = 20)
b3_0.place(x = 40, y = 33,width = 53,height = 53)

sheet, making_type_list = ["Sheet1","Sheet2"], ["本番","テスト"]
combo = make_combo(frame0, sheet, 45, 190,100,20)
combo2 = make_combo(frame0, making_type_list, 210, 190,100,20)

entry1 = Entry(frame0,justify=tkinter.CENTER,  bd = 0,bg="gray94",highlightthickness = 0,font=("HG丸ｺﾞｼｯｸM-PRO", 12, "bold"), foreground='gray50') #from
entry2 = Entry(frame0,justify=tkinter.CENTER,  bd = 0,bg="gray94",highlightthickness = 0,font=("HG丸ｺﾞｼｯｸM-PRO", 12, "bold"), foreground='gray50') # to
entry6 = Entry(frame0,bd = 0,bg="white",highlightthickness = 0,font=("HG丸ｺﾞｼｯｸM-PRO", 8, "bold"), foreground='gray40') # pagenum
entry7 = Entry(frame0,bd = 0,bg="white",highlightthickness = 0,font=("HG丸ｺﾞｼｯｸM-PRO", 8, "bold"), foreground='gray40') # pagenum
entry1.place(x = 60, y = 290, width = 60, height = 20)
entry2.place(x = 230, y = 290, width = 60, height = 20)
entry6.place(x = 330, y = 570, width = 200, height = 20)
entry7.place(x = 600, y = 570, width = 200, height = 20)
entry6.insert(0, "source/input.xlsx")
entry7.insert(0, "output/*****.xlsx")
entry6.configure(state='readonly')
entry7.configure(state='readonly')

txt_shousai = tkinter.scrolledtext.ScrolledText(frame0, font=("HG丸ｺﾞｼｯｸM-PRO", 11, "bold"), height=23, width=45, bg = "gray94", bd = 0, foreground='gray50')
txt_shousai.place(x = 385, y = 160)

b0_4['command'] = command=lambda: btn_click(path_list) 
b0_8['command'] = command=lambda: switch() 
b3_0['command'] = command=lambda: change_app(frame0) #rooms


"""
---navgate-----
"""
btnState = False       
frame1 = tk.Frame(frame0, bg="#ffffff", height=1000, width=250)
frame1.place(x=-300, y=0)

canvas1 = Canvas(frame1,bg = "#000000",height = 101,width = 250,bd = 0,highlightthickness = 0,relief = "ridge")
canvas1.place(x = 0, y = 0)

img1_0 = PhotoImage(file = f"asset/nav/img0.png",  master=frame1) # nav
img1_1 = PhotoImage(file = f"asset/nav/img1.png",  master=frame1) # DashBord
img1_3 = PhotoImage(file = f"asset/nav/img3.png",  master=frame1) # About
img1_5 = PhotoImage(file = f"asset/nav/img5.png",  master=frame1) # 下地

Labe1_5 = tk.Label(frame1, image=img1_5, background ="white")

b1_0  = Button(frame1, image = img1_0,borderwidth = 0,highlightthickness = 0,relief = "flat",bg = "#000000",cursor = "hand2", activebackground="gray93")
b1_1  = Button(frame1, image = img1_1,borderwidth = 0,highlightthickness = 0,relief = "flat",bg = "gray94",cursor = "hand2", activebackground="#000000")
b1_3  = Button(frame1, image = img1_3,borderwidth = 0,highlightthickness = 0,relief = "flat",bg = "gray94",cursor = "hand2", activebackground="#000000")

Labe1_5.place(x = 11, y = 111,width = 211,height = 446)
b1_0.place(x = 182, y = 41,width = 34,height = 20)
b1_1.place(x = 45, y = 123,width = 150,height = 55)
b1_3.place(x = 25, y = 183,width = 150,height = 55)

b1_0['command'] = command=lambda: switch() 
b1_1['command'] = command=lambda: switch()  #DashBord
b1_3['command'] = command=lambda: change_app(frame3) #about

frame0.tkraise()
window.resizable(False, False)
window.protocol("WM_DELETE_WINDOW", click_close); #終了確認
window.mainloop()

