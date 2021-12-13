#!/usr/bin/env python
# coding: utf-8

import time
from tkinter.filedialog import * # 不提前引用会导致第一次运行失败？
from tkinter import *
from tkinter import ttk
# Jeffrey - add sqlite 
import sqlite3


# 系统设计所需库
# from tkinter.filedialog import *
import tkinter.messagebox
import ctypes
from ctypes import windll
from PIL import Image, ImageTk
from sqlalchemy import create_engine

    
from PIL import Image, ImageTk
import xlwt
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.pylab import mpl
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg,NavigationToolbar2Tk
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure
plt.rcParams['font.family'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False


# 系统逻辑所需库
import re
import xlrd
import pandas as pd
import numpy as np
import pyodbc
import os
import datetime
from dateutil.relativedelta import relativedelta
import math
import pymssql
# pymssql依赖包缺失
from pymssql import _mssql
from pymssql import _pymssql
import uuid
import decimal

# for i in range(50,100):
#     # 每次更新加1
#     progressbarOne['value'] = i + 1
#     # 更新画面
#     start_root.update()
#     time.sleep(0.02)
# # 动画关闭
# start_root.destroy()
# start_root.mainloop()

# 关闭警告
import warnings
warnings.filterwarnings('ignore')

# Jeffrey - 添加其他文件中的python
from prism_database_operation import PrismDatabaseOperation


# Master表信息维护
class MasterData:
    # 查询code数据
    def master_search():
        SQL_SELECT = "SELECT Material FROM ProductMaster"
        material = PrismDatabaseOperation.Prism_select(SQL_SELECT).values
        return material
    
    # 主数据校验,校验是否存在，若否，则返回相应code
    def master_check(df):
        master_code = pd.DataFrame(columns=['Material'],data=MasterData.master_search())
        lack_code = df[~df['Material'].isin(master_code['Material'])]['Material']
        if lack_code.empty:
            return 0
        else:
            return lack_code

    # 修改数据（数据校验修改前后是否存在重复）
    def master_update(item_text):
        #print(item_text)
        # 数据修改窗口
        item_text = item_text
        modify_master = Tk()
        modify_master.title('主数据修改')
        modify_master.geometry('1050x250')

        # 获取主数据列名
        columns = ['规格型号','包装规格','分类Level3','分类Level4','ABC','分类','不含税单价',
                   '预测状态','MOQ','安全库存天数']
        modify_treeview = ttk.Treeview(modify_master, height=1, show="headings", columns=columns)
        modify_treeview.place(x=20,y=20)

        for i in range(len(columns)):
            modify_treeview.column(columns[i], width=100, anchor='center')

        # 显示列名
        for i in range(len(columns)):
            modify_treeview.heading(columns[i], text=columns[i])
        modify_treeview.insert('', 1, values=item_text)
        # 合并输入数据
        def set_value(event): 
            # 获取鼠标所选item
            for item in modify_treeview.selection():
                item_text = modify_treeview.item(item, "values")

            column = modify_treeview.identify_column(event.x)# 所在列
            row = modify_treeview.identify_row(event.y)# 所在行，返回
            cn = int(str(column).replace('#',''))
            #             rn = int(str(row).replace('I',''))
            #             print(row,rn,column,cn)
            entryedit = Entry(modify_master,width=13)
            entryedit.insert(0,str(item_text[cn-1]))
            entryedit.place(x=150, y=150)
            #             entryedit.place(x=20+(cn-1)*100, y=45) # 点击相应地方进行修改
            Label_select = Label(modify_master,text=str(item_text[cn-1]),width=20,anchor="w")
            Label_select.place(x=150, y=100)
            # 将编辑好的信息更新到数据库中
            def save_edit():
                # 获取
                modify_treeview.set(item, column=column, value=entryedit.get())
                entryedit.destroy()
                btn_input.destroy()
                btn_cancal.destroy()
                Label_select.destroy()
                
            btn_input = Button(modify_master, text='OK', width=7, command=save_edit)
            btn_input.place(x=260,y=150)
            
            # 取消输入
            def cancal_edit():
                entryedit.destroy()
                btn_input.destroy()
                btn_cancal.destroy()
                Label_select.destroy()
            
            btn_cancal = Button(modify_master, text='Cancel', width=7, command=cancal_edit)
            btn_cancal.place(x=350,y=150)
            
        # 触发双击事件
        modify_treeview.bind('<Double-1>', set_value)

        # 显示文本数据
        Label(modify_master,
              text="Tips：包装规格、不含税单价、MOQ、安全库存天数必须为数字！",
              fg='red').place(x=100,y=200)
        # Label(modify_master,text="例如：数字").place(x=10,y=90)
        Label(modify_master,text="修改前：").place(x=100,y=100)
        Label(modify_master,text="修改后：").place(x=100,y=150)

        def db_update():
            # 获取所有最新数据,直接更新所有数据
            # 先删除，再直接附加~更简单~

            # 遍历获取所有数据，并生成df
            try:
                t = modify_treeview.get_children()
                a = list()
                for i in t:
                    a.append(list(modify_treeview.item(i,'values')))
                df_now = pd.DataFrame(a,columns=columns)
                df_now.rename(columns={"规格型号":"Material","不含税单价":"GTS",
                                       "预测状态":"FCST_state"},inplace=True)
                for i in ["GTS","包装规格","MOQ","安全库存天数"]:
                    df_now[i] = df_now[i].astype(float)
                for i in range(len(df_now)):
                    SQL_delete = "DELETE FROM ProductMaster WHERE Material ='"+df_now['Material'].iloc[i]+"';"
                    PrismDatabaseOperation.Prism_delete(SQL_delete)
                insert = PrismDatabaseOperation.Prism_insert('ProductMaster',df_now)
                if insert != "error":
                    tkinter.messagebox.showinfo("提示","成功！")
            except:
                tkinter.messagebox.showerror("错误","修改失败！请检查数据格式")
            
            master_maintain()
            modify_master.destroy()

        Button(modify_master,text="确认修改",font=("黑体",12,'bold'),bg='slategrey',fg='white',
               width=9,height=1,borderwidth=5,command=db_update).place(x=920,y=200)

        modify_master.mainloop()

    # 批量修改数据
    def master_update_batch(df):
        # 数据修改窗口
        modify_master = Tk()
        modify_master.title('主数据修改')
        # 窗体大小随df变化而变化
        modify_master.geometry('1050x500')
        
        columns = list(df.columns)
        modify_treeview = ttk.Treeview(modify_master, height=10, show="headings", 
                                       columns=columns)
        modify_treeview.place(x=20,y=20)
        sb = ttk.Scrollbar(modify_master,command=modify_treeview.yview)
        sb.config(command=modify_treeview.yview)
        sb.place(x=975,y=0,in_ = modify_treeview,height=230)
        modify_treeview.config(yscrollcommand=sb.set)
        
        #  表示列,不显示,文本靠左，数字靠右
        for i in columns:
        # print(i,ProductMaster[i].dtypes)
            if df[i].dtypes == float or df[i].dtypes == int:
                modify_treeview.column(i, width=97, anchor='center')
            else:
                modify_treeview.column(i, width=97, anchor='center')
                
        # 显示列
        for i in range(len(df.columns)):
            modify_treeview.heading(str(df.columns[i]), text=str(df.columns[i]))
        
        # 插入数据
        for i in range(len(df)):
            modify_treeview.insert('', i, 
                                   values=(df[df.columns[0]].iloc[i],
                                           df[df.columns[1]].iloc[i],
                                           df[df.columns[2]].iloc[i],
                                           df[df.columns[3]].iloc[i],
                                           df[df.columns[4]].iloc[i],
                                           df[df.columns[5]].iloc[i],
                                           df[df.columns[6]].iloc[i],
                                           df[df.columns[7]].iloc[i],
                                           df[df.columns[8]].iloc[i],
                                           df[df.columns[9]].iloc[i]))
        
        # Treeview、列名、排列方式
        def treeview_sort_column(tv, col, reverse):  
            L = [(tv.set(k, col), k) for k in tv.get_children('')]
            try:
                for i in range(len(L)):
                    L[i] = (float(L[i][0].replace(',', '')),L[i][1])
            except:
                pass
            L.sort(reverse=reverse)  # 排序方式
            # 根据排序后索引移动
            for index, (val, k) in enumerate(L):
                tv.move(k, '', index)
            # 重写标题，使之成为再点倒序的标题
            tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))  

        # 绑定函数，使表头可排序
        for col in columns:
            modify_treeview.heading(col, 
                                    text=col, 
                                    command=lambda _col=col: treeview_sort_column(modify_treeview, 
                                                                                  _col, 
                                                                                  False))
        # 合并输入数据
        def set_value(event): 
            # 获取鼠标所选item
            for item in modify_treeview.selection():
                item_text = modify_treeview.item(item, "values")

            column = modify_treeview.identify_column(event.x)# 所在列
            cn = int(str(column).replace('#',''))
            
            Label_select = Label(modify_master,text=str(item_text[cn-1]),width=20)
            Label_select.place(x=150, y=300)
            entryedit = Text(modify_master,width=15,height = 1)
            entryedit.place(x=150, y=350)

            # 将编辑好的信息更新到数据库中
            def save_edit():
                # 获取
                modify_treeview.set(item, column=column, value=entryedit.get(0.0, "end")[:-1])
                entryedit.destroy()
                btn_input.destroy()
                btn_cancal.destroy()
                Label_select.destroy()
                
            btn_input = Button(modify_master, text='OK', width=7, command=save_edit)
            btn_input.place(x=260,y=350)
            
            # 取消输入
            def cancal_edit():
                entryedit.destroy()
                btn_input.destroy()
                btn_cancal.destroy()
                Label_select.destroy()
            
            btn_cancal = Button(modify_master, text='Cancel', width=7, command=cancal_edit)
            btn_cancal.place(x=350,y=350)

        # 触发双击事件
        modify_treeview.bind('<Double-1>', set_value)

        # 显示文本数据
        Label(modify_master,text="修改前：").place(x=100,y=300)
        Label(modify_master,text="修改后：").place(x=100,y=350)

        def db_update():
            # 获取所有最新数据,直接更新所有数据
            # 先删除，再直接附加~更简单~
            # 遍历获取行列
            t = modify_treeview.get_children()
            a = list()
            for i in t:
                a.append(list(modify_treeview.item(i,'values')))
            column = list(PrismDatabaseOperation.Prism_select('SELECT TOP 1 * FROM ProductMaster').columns)
            df_now = pd.DataFrame(a,columns=column)
            for i in ["GTS","包装规格","MOQ","安全库存天数"]:
                df_now[i] = df_now[i].astype(float)
            # 删除已有material
            for i in range(len(df_now)):
                SQL_delete = "DELETE FROM ProductMaster WHERE Material ='"+                 df_now['Material'].iloc[i]+"';"
                PrismDatabaseOperation.Prism_delete(SQL_delete)
            insert = PrismDatabaseOperation.Prism_insert('ProductMaster',df_now)
            if insert != "error":
                tkinter.messagebox.showinfo("提示","成功！")
            modify_master.destroy()
            master_maintain()
            
        Button(modify_master,text="确认修改",font=("黑体",12,'bold'),bg='slategrey',fg='white',
               width=9,height=1,borderwidth=5,command=db_update).place(x=920,y=400)
        
        # 提示文本
        Label(modify_master,text="Tips：包装规格、不含税单价、MOQ、安全库存天数必须为数字！",
              fg='red').place(x=100,y=400)

        modify_master.mainloop()



# 链接数据库 python 对sql数据库操作进行封装
#-- coding: utf-8 --
# Prism 数据库操作，输入sql语句，返回查询df、修改数据、删除数据等
# Jeffrey - 多次服用，已经拆出来一个单的文件和类
class PrismDatabaseOperation:
    """
    执行Prism数据库sql语句的函数，可进行增、删、改、查
    注：增加数据时不直接通过读取excel文件输入，而是先read处理后再insert
    """
    
    # 查询
    def Prism_select(SQL, df=""):
        try:
            # 定义数据库参数
            server = "localhost"          # 服务器名称
            user = "ryan"                 # 用户名
            password = "prism123+"        # 链接密码
            database = "Prism"            # 数据库名
            charset = "cp936"               # charset编码规则中文
            try:
                # 连接数据库
                #-- coding: utf-8 --
                # conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+
                #                       ';DATABASE='+database+';UID='+user+';PWD='+password)
                # conn = pymssql.connect(server=server, user=user, password=password, database=database)
                conn = sqlite3.connect('prism_data.db')
                cursor = conn.cursor() # 句柄
            except:
                tkinter.messagebox.showerror("错误","连接数据库失败！")
            # 单次查询指定sql
            if df == "":
                cursor.execute(SQL) # 执行语句
                data = cursor.fetchall()
                columnDes = cursor.description #获取连接对象的描述信息
                columnNames = [columnDes[i][0] for i in range(len(columnDes))]
                df = pd.DataFrame([list(i) for i in data],columns=columnNames)
                cursor.close() # 关闭句柄
                return df
        except:
            tkinter.messagebox.showerror("错误","数据查询失败！")

    # 修改
    def Prism_update(SQL, df=""):
        try:
            # 定义数据库参数
            server = "localhost"          # 服务器名称
            user = "ryan"                 # 用户名
            password = "prism123+"        # 链接密码
            database = "Prism"            # 数据库名
            charset = "gbk"               # charset编码规则中文
            try:
                # 连接数据库
                # conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+
                #                       ';DATABASE='+database+';UID='+user+';PWD='+password)
                conn = sqlite3.connect('prism_data.db')
                cursor = conn.cursor() # 句柄
                
            except:
                tkinter.messagebox.showerror("错误","连接数据库失败！")
                
            cursor.execute(SQL)
            cursor.commit()
            cursor.close() # 关闭句柄
        except:
            tkinter.messagebox.showerror("错误","数据更新失败！")

    # 删除
    def Prism_delete(SQL, df=""):
        try:
            # 定义数据库参数
            server = "localhost"          # 服务器名称
            user = "ryan"                 # 用户名
            password = "prism123+"        # 链接密码
            database = "Prism"            # 数据库名
            charset = "gbk"               # charset编码规则中文
            try:
                # 连接数据库
                # conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+
                #                       ';DATABASE='+database+';UID='+user+';PWD='+password)
                conn = sqlite3.connect('prism_data.db')
                cursor = conn.cursor() # 句柄
            except:
                tkinter.messagebox.showerror("错误","连接数据库失败！")
            cursor.execute(SQL)
            cursor.commit()
            cursor.close() # 关闭句柄
        except:
            tkinter.messagebox.showerror("错误","数据删除失败！")

    # 增加,特殊方法，直接附加df,如有null，替换为0
    def Prism_insert(db_sheetname, df=""):
        try:
            try:
                #  定义数据库参数,连接数据库
                # conn = create_engine('mssql+pymssql://ryan:prism123+@localhost/Prism')
                conn = sqlite3.connect('prism_data.db')
            except:
                tkinter.messagebox.showerror("错误","连接数据库失败！")
            # 附加df
            df.to_sql(db_sheetname, con=conn, if_exists='append', index=False)
            conn.dispose() # 关闭链接
            return "right"
#             tkinter.messagebox.showinfo("提示",db_sheetname+"数据插入成功！")
        except:
            tkinter.messagebox.showerror("错误",db_sheetname+"数据插入失败！")
            return "error"


# 部分全局函数

# 主界面文本风格设置
# Jeffrey - 该函数没用？？
def s_2():
    place_x = 0
    place_y = 0
    font = ('黑体',16)
    bg='WhiteSmoke'
    fg='DimGray'
    return place_x,place_y,font,bg,fg

# 返回前N月的JNJ_Month
# Jeffrey - 是否确保按照JNJ_Date排序
def JNJ_Month(n):
    SQL_select_date = "SELECT JNJ_Date From Outbound"
    # 建议修改成 select JNJ_Date FROM Outbound ORDER BY JNJ_Date
    Outbound_month = list(PrismDatabaseOperation.Prism_select(SQL_select_date)['JNJ_Date'].unique())
    last_month = sorted(Outbound_month)[-n:]
    return last_month

# df含千分位字符转数字
# Jeffrey - 该函数没用？
def convert(item):
    if isinstance(item, str):
        if ',' not in item: 
            return float(item)
        s = ''
        tmp = item.strip().split(',')
        for i in range(len(tmp)):
            s += tmp[i]
        return float(s)
    else:
        return 'Type transformed'

# python 自带round精度问题，四舍五入不准确（_float输入浮点数, _len小数点位数）
# Jeffrey - 可以采用Decimal模块直接实现
def new_round(_float, _len):
    if isinstance(_float, float):
        if str(_float)[::-1].find('.') <= _len:
            return(_float)
        if str(_float)[-1] == '5':
            return(round(float(str(_float)[:-1]+'6'), _len))
        else:
            return(round(_float, _len))
    else:
        return(round(_float, _len))

# 鼠标移动提醒文本
# https://blog.csdn.net/qq_46329012/article/details/115767178
# 对该控件的定义
# Jeffrey - 该函数没用？？
class ToolTip(object):

    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0
    #当光标移动指定控件是显示消息
    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx()+30
        y = y + cy + self.widget.winfo_rooty()+30
        self.tipwindow = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = Label(tw, text=self.text,justify=LEFT,
                      background="white", relief=SOLID, borderwidth=1,
                      font=("华文细黑", "15"))
        label.pack(side=BOTTOM)
    #当光标移开时提示消息隐藏
    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

#创建该控件的函数
"""
第一个参数：是定义的控件的名称
第二个参数，是要显示的文字信息
"""
def CreateToolTip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

# 重设小数位
# Jeffrey - 或者直接用 "{:.2f}".format(XXXXXX)
def re_round(a):
    remain_amount = "%.2f" % a
    remain_amount_format =re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,", remain_amount)
    return remain_amount_format
# 右键菜单
# def cut(editor, event=None):
#     editor.event_generate("<<Cut>>")

# #复制功能的实现
# def copy(editor, event=None):
#     editor.event_generate("<<Copy>>")

# #粘贴功能的实现
# def paste(editor, event=None):
#     editor.event_generate('<<Paste>>')

# #右键帮定的函数
# '''
# 使用的时候定义一个Menubar控件，然后将其传给rightkey()函数
# '''
# def rightKey(menubar, event, editor):
#     menubar.delete(0, END)
#     menubar.add_command(label='复制', command=lambda: copy(editor))
#     menubar.add_separator()
#     menubar.add_command(label='剪切', command=lambda: cut(editor))
#     menubar.add_separator()
#     menubar.add_command(label='粘贴', command=lambda: paste(editor))
#     menubar.post(event.x_root, event.y_root)



# ---------------- Prism用户主界面 -----------------#
window = Tk()
window.title('Prism v2.1')
window.geometry('1285x725+100+100')
# 去掉window自带窗体和不显示在任务栏
window.overrideredirect(True)
# 界面显示到任务栏
def set_appwindow(window):
    GWL_EXSTYLE=-20
    WS_EX_APPWINDOW=0x00040000
    WS_EX_TOOLWINDOW=0x00000080
    hwnd = windll.user32.GetParent(window.winfo_id())
    style = windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    style = style & ~WS_EX_TOOLWINDOW
    style = style | WS_EX_APPWINDOW
    res = windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
    # re-assert the new window style
    window.wm_withdraw()
    window.after(10, lambda: window.wm_deiconify())

window.after(10, lambda: set_appwindow(window))

window.configure(bg='Gainsboro')
# 图标
window.iconbitmap(os.path.abspath('.')+'\\Picture\\PrismLogo.ico')

# 提高清晰度
# # 告诉操作系统使用程序自身的dpi适配
# ctypes.windll.shcore.SetProcessDpiAwareness(1)
# # 获取屏幕的缩放因子
# ScaleFactor=ctypes.windll.shcore.GetScaleFactorForDevice(0)
# # 设置程序缩放 75%
# window.tk.call('tk', 'scaling', ScaleFactor/70)


# 导入系统图片,设置背景等图标
#  背景设置
img_bg = Image.open(os.path.abspath('.')+'\\Picture\\Background\\背景(无按键).png')
img_bg_png = ImageTk.PhotoImage(img_bg)
BgLabel = Label(window,justify = LEFT,image = img_bg_png,compound = CENTER)
BgLabel.place(x=0,y=0)
# logo
img_logo = Image.open(os.path.abspath('.')+
                      '\\Picture\\Button\\一级目录Button\\未选中（灰色）\\1.PrismLogo_Gray.png')
img_logo_png = ImageTk.PhotoImage(img_logo)
LogoLabel = Label(window,justify = LEFT,image = img_logo_png,compound = CENTER,height=40,width=40)
LogoLabel.place(x=15,y=10)
# CreateToolTip(LogoLabel, "Prism")
# 搜索框（待更新）
# img_search = Image.open(os.path.abspath('.')+'\\Picture\\Background\\搜索.png')
# img_search_png = ImageTk.PhotoImage(img_search)
# SearchLabel = Label(window,width=158,height=33,bg='WhiteSmoke',image = img_search_png)
# SearchLabel.place(x=90,y=14)
# CreateToolTip(SearchLabel, "待更新")

# 窗体移动
def window_move(event):
    window_x = 200 # 鼠标在窗体中的位置
    window_y = 10
    new_x = event.x  + window.winfo_x() - window_x
    new_y = event.y + window.winfo_y() - window_y
#     print(event.x,event.y)
    geo_str=f"{'1285x725'}+{new_x}+{new_y}"
    window.geometry(geo_str)

lb_move = Label(window,width=100,height=2,bg='WhiteSmoke')
lb_move.place(x=500,y=10)
lb_move.bind('<B1-Motion>', window_move)
CreateToolTip(lb_move, "点击按住此处即可拖动窗体")
# 按钮
# 更新按钮:按下前后
img_btn_update_1 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\一级目录Button\\未选中（灰色）\\2.DataUpdate_Gray.png')
img_btn_update_png_1 = ImageTk.PhotoImage(img_btn_update_1)
img_btn_update_2 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\一级目录Button\\选中（蓝色）\\2.DataUpdate_Blue.png')
img_btn_update_png_2 = ImageTk.PhotoImage(img_btn_update_2)
#  月度数据更新按钮背景 
img_btn_update_3 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\数据更新\\数据上传（蓝 6x6).png')
img_btn_update_png_3 = ImageTk.PhotoImage(img_btn_update_3)

# 预测按钮
img_btn_FCST_1 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\一级目录Button\\未选中（灰色）\\3.DemandFCST_Gray.png')
img_btn_FCST_png_1 = ImageTk.PhotoImage(img_btn_FCST_1)
img_btn_FCST_2 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\一级目录Button\\选中（蓝色）\\3.DemandFCST_Blue.png')
img_btn_FCST_png_2 = ImageTk.PhotoImage(img_btn_FCST_2)

# 补货按钮
img_btn_Rep_1 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\一级目录Button\\未选中（灰色）\\4.Rep_Gray.png')
img_btn_Rep_png_1 = ImageTk.PhotoImage(img_btn_Rep_1)
img_btn_Rep_2 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\一级目录Button\\选中（蓝色）\\4.Rep_Blue.png')
img_btn_Rep_png_2 = ImageTk.PhotoImage(img_btn_Rep_2)

# 进出库记录按钮
img_btn_Track_1 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\一级目录Button\\未选中（灰色）\\5.Track_Gray.png')
img_btn_Track_png_1 = ImageTk.PhotoImage(img_btn_Track_1)
img_btn_Track_2 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\一级目录Button\\选中（蓝色）\\5.Track_Blue.png')
img_btn_Track_png_2 = ImageTk.PhotoImage(img_btn_Track_2)

# 设置按钮
img_btn_Set_1 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\一级目录Button\\未选中（灰色）\\6.Set_Gray.png')
img_btn_Set_png_1 = ImageTk.PhotoImage(img_btn_Set_1)
img_btn_Set_2 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\一级目录Button\\选中（蓝色）\\6.Set_Blue.png')
img_btn_Set_png_2 = ImageTk.PhotoImage(img_btn_Set_2)

# 更多按钮
img_btn_More_1 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\一级目录Button\\未选中（灰色）\\7.More_Gray.png')
img_btn_More_png_1 = ImageTk.PhotoImage(img_btn_More_1)
img_btn_More_2 = Image.open(
    os.path.abspath('.')+'\\Picture\\Button\\一级目录Button\\选中（蓝色）\\7.More_Blue.png')
img_btn_More_png_2 = ImageTk.PhotoImage(img_btn_More_2)

# 文本

# 输入框

# 下拉框

# 主数据维护界面
def master_maintain():
    """主数据维护界面同步按钮发生变化"""
    # 标题
    lb_title = Label(window,text='主数据维护 ',font=('华文中宋',14),bg='WhiteSmoke',
                                  fg='black',width=10,height=2)
    lb_title.place(x=280,y=10)
    
    # 表格数据
    ProductMaster = PrismDatabaseOperation.Prism_select('SELECT * FROM ProductMaster')
    
    # 规范格式
    ProductMaster.fillna("-",inplace=True)
    ProductMaster.replace("nan","-",inplace=True)
    ProductMaster.replace("","-",inplace=True)
    ProductMaster.rename(columns={"Material":"规格型号","GTS":"不含税单价","FCST_state":"预测状态"}
                         ,inplace=True)
    # 数据小数位规整：包装规格、MOQ整数，安全库存天数为1位，不含税单价2位
    for i in ['包装规格','MOQ']:
        try:
            if "-" in ProductMaster[i]:
                for i in range(len(ProductMaster)):
                    ProductMaster[i].iloc[i] = int(ProductMaster[i].iloc[i])
            else:
                ProductMaster[i] = ProductMaster[i].astype(int)
        except:
            pass

    try:
        for i in range(len(ProductMaster)):
            ProductMaster['不含税单价'].iloc[i] = '%.2f' %ProductMaster['不含税单价'].iloc[i]
    except:
        pass

    try:
        for i in range(len(ProductMaster)):
            ProductMaster['安全库存天数'].iloc[i] = '%.1f' % ProductMaster['安全库存天数'].iloc[i]
    except:
        pass


    # 出现product—master主要数据，可双击treeview显示编辑
    frame = Frame(window,height=655,width=1015,bg='WhiteSmoke')
    # label，提示双击可以进行相应数据的编辑
    Label(frame,text='* 操作提示：双击相应数据可以进行编辑；点击列名即可排序',bg='WhiteSmoke',
          font=("黑体",10)).place(x=0,y=628)
    columns = list(ProductMaster.columns)
    
    # 设置样式
    style_head = ttk.Style()
    style_head.configure("MyStyle.Treeview.Heading",rowheight=50,font=("华文中宋",12))
    style_value = ttk.Style()
    style_value.configure("MyStyle.Treeview", rowheight=24)
    treeview = ttk.Treeview(frame, height=21, show="headings",selectmode="extended",
                            columns=columns,style='MyStyle.Treeview')
    
    # 添加滚动条
    # 竖向滚动条
    sb_y = ttk.Scrollbar(frame,command=treeview.yview)
    sb_y.config(command=treeview.yview)
    sb_y.place(in_=treeview,relx=1, rely=0,relheight=1)
    treeview.config(yscrollcommand=sb_y.set)
    # 横向滚动条
    sb_x = ttk.Scrollbar(frame,command=treeview.xview,orient="horizontal")
    sb_x.config(command=treeview.xview)
    sb_x.place(in_=treeview,relx=0, rely=1,relwidth=1)
    treeview.config(xscrollcommand=sb_x.set)
    treeview.place(x=0,y=70,relwidth=0.98)
    frame.place(x=267,y=61)
    
    # 行交替颜色
    def fixed_map(option):# 重要！无此步骤则无法显示
        return [elm for elm in style.map("Treeview", query_opt=option)
                if elm[:2] != ("!disabled", "!selected")]
    style = ttk.Style()
    style.map("Treeview",foreground=fixed_map("foreground"),background=fixed_map("background"))
    
    treeview.tag_configure('oddrow', background='LightGrey')
    treeview.tag_configure('evenrow', background='white')
    
    # 行坐标重排
    def odd_even_color():
        for index,row in enumerate(treeview.get_children()):
            if index % 2 == 0:
                treeview.item(row,tags="evenrow")
            else:
                treeview.item(row,tags="oddrow")
    
    #  表示列,不显示,文本靠左，数字靠右
    for i in columns:
    #         print(i,ProductMaster[i].dtypes)
        if ProductMaster[i].dtypes == float or ProductMaster[i].dtypes == int:
            treeview.column(i, width=95, anchor='center')
        else:
            treeview.column(i, width=95, anchor='center')
    
    # 显示表头
    for i in columns:
        treeview.heading(str(i), text=str(i))
    
    # 插入数据
    for i in range(len(ProductMaster)):
        if i % 2 == 0:
            tag = "evenrow"
        else:
            tag = "oddrow"
        treeview.insert('', i, values=list(ProductMaster.iloc[i,:]),tags=tag)
        
    # Treeview、列名、排列方式
    def treeview_sort_column(tv, col, reverse):
        L = [(tv.set(k, col), k) for k in tv.get_children('')]
        try:
            for i in range(len(L)):
                if L[i][0] == "-":
                    L[i] = (float(0),L[i][1])
                else:
                    L[i] = (float(L[i][0].replace(',', '')),L[i][1])
        except:
            pass
        #         print(L)
        L.sort(reverse=reverse)  # 排序方式
        
        # 根据排序后索引移动
        for index, (val, k) in enumerate(L):
            tv.move(k, '', index)
        # 重写标题，使之成为再点倒序的标题
        tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))
        odd_even_color() # 行坐标标签重排
    
    # 绑定函数，使表头可排序
    for col in columns:
        treeview.heading(col, text=col, command=
                         lambda _col=col: treeview_sort_column(treeview, _col, False))
    
    # 双击进入编辑状态，弹出编辑界面
    def set_cell_value(event):
        item_text = treeview.item(treeview.selection(), "values")
        MasterData.master_update(item_text)
    
    treeview.bind('<Double-1>', set_cell_value)
    
    # 添code按钮
    def add_material():
        to_updata = pd.DataFrame(data=[],columns=columns)
        to_updata.loc[0] = ["规格型号","数字","","","","","数字","MTS","数字","数字"]
        MasterData.master_update_batch(to_updata)
    
    btn_add_material = Button(frame,text="单个添加",font=("heiti",10,'bold'),bg='slategrey',
                              fg='white',width=9,height=1,borderwidth=5,
                              command=add_material)
    btn_add_material.place(x=800,y=30)
    
    # 批量上传code
    def batch_material():
        filename = tkinter.filedialog.askopenfilename().replace("/","\\")
        batch_code = pd.read_excel(filename,dtype={"包装规格":float,"不含税单价":float})
        # 批量插入数据，如已存在，则直接覆盖已有数据
        MasterData.master_update_batch(batch_code)
        
    btn_batch_material = Button(frame,text="批量上传",font=("heiti",10,'bold'),bg='slategrey',
                                fg='white',width=9,height=1,borderwidth=5,
                                command=batch_material)
    btn_batch_material.place(x=700,y=30)

        # 删除指定行code
    #     def del_material():
    #         pass
    #     btn_del_material = Button(frame,text="选中删除",bg="grey",font=("heiti",10),width=9,
    #                               height=1,borderwidth=5,command=del_material)
    #     btn_del_material.place(x=600,y=30)
    
    # 搜索code功能
    Label(frame,text="筛选字段：",bg='WhiteSmoke',font=("黑体",12)).place(x=0,y=35)
    cbx = ttk.Combobox(frame,font=("黑体",11),width=10) #筛选字段
    comvalue = tkinter.StringVar()
    cbx["values"] = ["全局搜索"] + columns
    cbx.current(1)
    cbx.place(x=80,y=37)
    entry_search = Entry(frame,font=("黑体",11),width=12) # 筛选内容
    entry_search.insert(0, "请输入信息")
    entry_search.place(x=190,y=37)
    #     CreateToolTip(entry_search, "请注意大小写输入！")
    # 先清空表格，再插入数据，当字段选择为空、内容为空则显示全部
    def search_material():
        search_all = ProductMaster.copy()
        for i in search_all.columns:
            search_all[i] = search_all[i].apply(str)# 必须转字符，否则无法全局搜索
        
        # 清空
        for item in treeview.get_children():
            treeview.delete(item)
        # 查找并插入数据
        if entry_search.get() != "":
            search_content = str(entry_search.get())
            # 全局搜索
            if cbx.get() == "全局搜索":
                search_df = pd.DataFrame(columns=search_all.columns)
                for i in range(len(search_all.columns)):
                    search_df = search_df.append(search_all[search_all[
                        search_all.columns[i]].str.contains(search_content)])                
                search_df.drop_duplicates(subset=["规格型号"], keep='first',inplace=True)
                #                 print(search_df)
            # 指定字段搜索
            else:
                appoint = str(cbx.get())
                search_df = search_all[search_all[appoint].str.contains(search_content)]
                #                 print(search_df)
            # 插入表格
            for i in range(len(search_df)):
                if i % 2 == 0:
                    tag = "evenrow"
                else:
                    tag = "oddrow"
                treeview.insert('', i, values=list(search_df.iloc[i,:]),tags=tag)
        # 若输入值为空则显示全部内容
        else:
            # 插入
            for i in range(len(search_all)):
                if i % 2 == 0:
                    tag = "evenrow"
                else:
                    tag = "oddrow"
                treeview.insert('', i, values=list(search_all.iloc[i,:]),tags=tag)
            
    btn_search_material = Button(frame,text="查找",font=("黑体",10,'bold'),bg='slategrey',
                                 fg='white',width=9,height=1,borderwidth=5,
                                 command=search_material)
    btn_search_material.place(x=330,y=30)
    
    # 选择路径，输出保存
    def output_FCST():
        filename = tkinter.filedialog.asksaveasfilename()
        # 遍历获取所有数据，并生成df
        # 改变文本存储的数字
        t = treeview.get_children()
        a = list()
        for i in t:
            a.append(list(treeview.item(i,'values')))
        df_now = pd.DataFrame(a,columns=columns)

        # 指定列
        for i in range(0,len(df_now.columns)):
            try:
                df_now[df_now.columns[i]] = df_now.loc[:,df_now.columns[i]].apply(
                    lambda x: float(x.replace(",", "")))
            except:
                pass

        df_now.to_excel(filename+".xls",index=False)

    btn_output = Button(frame,text="导出",font=("heiti",10,'bold'),bg='slategrey',
                        fg='white',width=9,height=1,borderwidth=5,command=output_FCST)
    btn_output.place(x=900,y=30)


# 模板下载
def create_model():
    # 创建df模板
    model_df = pd.DataFrame(data=[["请输入不带空格值","请输入数值"]],columns=["规格型号","数量"])
    # 获取指定路径
    path = filedialog.askdirectory().replace("/","\\")
    # 获取上个月时间
    last_month = (datetime.datetime.today()+relativedelta(months=-1)).strftime("%Y%m")
    model_df.to_excel(path+"\\缺货_"+last_month+".xls",index=False)
    model_df.to_excel(path+"\\可发量_"+last_month+".xls",index=False)
    model_df.to_excel(path+"\\销售出库_"+last_month+".xls",index=False)
    model_df.to_excel(path+"\\在途_"+last_month+".xls",index=False)
    model_df.to_excel(path+"\\预入库_"+last_month+".xls",index=False)


# 读取数据文件所在文件夹,并进行数据更新,导入数据时附加强生年月，强生年月根据文件名,并维护主数据
def read_dir():
    """
    选择数据所在文件路径，读取并添加入数据库
    """
    # 提示导入成功
    # 不用先做是否存在的判断，直接merge后进行判断主数据的”预测状态“是否为空
    # merge后，弹出报错提示框，与现有的masterdata对比，显示出错的mastercode信息，为空则让他自己添加
    # 在点击确认按钮之后再进行对比，直至没有报错。
    
    path = filedialog.askdirectory().replace("/","\\")
    # 返回导入失败后的文件名
    error_input = []
    # 返回已有数据的文件名
    already_input = []
    # 返回缺失的文件名
    lack_input = ["在途","可发量","预入库","缺货","销售出库"]
    for name in os.listdir(path):
        # 定义强生年月、文件路径
        jnj_date = name[name.find("_")+1:name.find("_")+7]
        file_path = os.path.join(path, name)
        # 导入销售库存信息
        if "在途" in name:
            lack_input.remove("在途")
            try:
                Intransit = pd.read_excel(file_path,dtype={'规格型号':str})
                Intransit['JNJ_Date'] = jnj_date
                # 统一命名
                Intransit.rename(columns={'规格型号':'Material','数量':'Intransit_QTY'},inplace=True)
                # 使用prism类，导入df数据，导入之前判断数据库的强生年月（文件名中含有）中是否
                # 已经存在，若有，则跳过导入数据；若无，则直接添加；数据库为空则直接添加
                JNJ_Date_exist = PrismDatabaseOperation.Prism_select('SELECT JNJ_Date FROM Intransit')
                Intransit = Intransit[['Material','Intransit_QTY','JNJ_Date']]
                if pd.isnull(Intransit.at[0,'Material'])==False:
                    Intransit.drop(Intransit[Intransit["Material"].isnull()].index,inplace=True)
                    Intransit.fillna(0,inplace=True)
                    if JNJ_Date_exist.empty:
                        insert_Intransit = PrismDatabaseOperation.Prism_insert('Intransit',Intransit)
                        if insert_Intransit == "error":
                            error_input.append(name)
                    else:
                        if jnj_date in JNJ_Date_exist['JNJ_Date'].unique():
                            already_input.append(name)
                            SQL_select = "SELECT * FROM Intransit WHERE JNJ_Date = '"+jnj_date+"'"
                            Intransit = PrismDatabaseOperation.Prism_select(SQL_select)
                        else:
                            insert_Intransit = PrismDatabaseOperation.Prism_insert('Intransit',Intransit)
                            if insert_Intransit == "error":
                                error_input.append(name)
                else:
                    Intransit.drop(Intransit[Intransit["Material"].isnull()].index,inplace=True)
                    lack_input.append(name)
            except:
                error_input.append(name)

        # 导入可发量库存信息
        elif "可发量" in name:
            lack_input.remove("可发量")
            try:
                Onhand = pd.read_excel(file_path,dtype={'规格型号':str})
                Onhand['JNJ_Date'] = jnj_date
                # 数据库中列名小括号引发列名不一致问题
                Onhand.rename(columns={'规格型号':'Material','数量':'Onhand_QTY'},inplace=True)
                JNJ_Date_exist = PrismDatabaseOperation.Prism_select('SELECT JNJ_Date FROM Onhand')
                Onhand = Onhand[['Material','Onhand_QTY','JNJ_Date']]
                if pd.isnull(Onhand.at[0,'Material'])==False:
                    Onhand.drop(Onhand[Onhand["Material"].isnull()].index,inplace=True)
                    Onhand.fillna(0,inplace=True)
                    if JNJ_Date_exist.empty:
                        insert_Onhand = PrismDatabaseOperation.Prism_insert('Onhand',Onhand)
                        if insert_Onhand == "error":
                            error_input.append(name)
                    else:
                        if jnj_date in JNJ_Date_exist['JNJ_Date'].unique():
                            already_input.append(name)
                            SQL_select = "SELECT * FROM Onhand WHERE JNJ_Date = '"+jnj_date+"'"
                            Onhand = PrismDatabaseOperation.Prism_select(SQL_select)
                        else:
                            insert_Onhand = PrismDatabaseOperation.Prism_insert('Onhand',Onhand)
                            if insert_Onhand == "error":
                                error_input.append(name)
                else:
                    Onhand.drop(Onhand[Onhand["Material"].isnull()].index,inplace=True)
                    lack_input.append(name)
            except:
                error_input.append(name)

        # 导入预入库信息
        elif "预入库" in name:
            lack_input.remove("预入库")
            try:
                Putaway = pd.read_excel(file_path,dtype={'规格型号':str})
                Putaway['JNJ_Date'] = jnj_date
                Putaway.rename(columns={'规格型号':'Material','数量':'Putaway_QTY'},inplace=True)
                JNJ_Date_exist = PrismDatabaseOperation.Prism_select('SELECT JNJ_Date FROM Putaway')
                Putaway = Putaway[['Material','Putaway_QTY','JNJ_Date']]
                # 数据源不会空
                if pd.isnull(Putaway.at[0,'Material'])==False:
                    # 删除规格型号为空的行
                    Putaway.drop(Putaway[Putaway["Material"].isnull()].index,inplace=True)
                    Putaway.fillna(0,inplace=True)
                    if JNJ_Date_exist.empty:
                        insert_Putaway = PrismDatabaseOperation.Prism_insert('Putaway',Putaway)
                        if insert_Putaway == "error":
                            error_input.append(name)
                    else:
                        if jnj_date in JNJ_Date_exist['JNJ_Date'].unique():
                            already_input.append(name)
                            SQL_select = "SELECT * FROM Putaway WHERE JNJ_Date = '"+jnj_date+"'"
                            Putaway = PrismDatabaseOperation.Prism_select(SQL_select)
                        else:
                            insert_Putaway = PrismDatabaseOperation.Prism_insert('Putaway',Putaway)
                            if insert_Putaway == "error":
                                error_input.append(name)
                else:
                    Putaway.drop(Putaway[Putaway["Material"].isnull()].index,inplace=True)
                    lack_input.append(name)
            except:
                error_input.append(name)
                
        # 导入销售出库信息
        elif "销售出库" in name:
            lack_input.remove("销售出库")
            try:
                Outbound = pd.read_excel(file_path,dtype={'规格型号':str})
                Outbound['JNJ_Date'] = jnj_date
                Outbound.rename(columns={'规格型号':'Material','数量':'Outbound_QTY'},inplace=True)
                # 销售数据直接先保存，当进行merge的时候，需要进行groupby，当前不做处理
                Outbound = Outbound[['Material','Outbound_QTY','JNJ_Date']]
                Outbound = Outbound.groupby(by=['JNJ_Date','Material'],
                                              as_index=False).sum()
                JNJ_Date_exist = PrismDatabaseOperation.Prism_select('SELECT JNJ_Date FROM Outbound')
                if pd.isnull(Outbound.at[0,'Material'])==False:
                    Outbound.drop(Outbound[Outbound["Material"].isnull()].index,inplace=True)
                    Outbound.fillna(0,inplace=True)
                    if JNJ_Date_exist.empty:
                        insert_Outbound = PrismDatabaseOperation.Prism_insert('Outbound',Outbound)
                        if insert_Outbound == "error":
                            error_input.append(name)
                    else:
                        if jnj_date in JNJ_Date_exist['JNJ_Date'].unique():
                            already_input.append(name)
                            SQL_select = "SELECT * FROM Outbound WHERE JNJ_Date = '"+jnj_date+"'"
                            Outbound = PrismDatabaseOperation.Prism_select(SQL_select)
                        else:
                            insert_Outbound = PrismDatabaseOperation.Prism_insert('Outbound',Outbound)
                            if insert_Outbound == "error":
                                error_input.append(name)
                else:
                    Outbound.drop(Outbound[Outbound["Material"].isnull()].index,inplace=True)
                    lack_input.append(name)
            except:
                error_input.append(name)           

        # 导入缺货数据
        elif "缺货" in name:
            lack_input.remove("缺货")
            try:
                Backorder = pd.read_excel(file_path)
                Backorder['JNJ_Date'] = jnj_date
                Backorder.rename(columns={'规格型号':'Material','数量':'Backorder_QTY'},inplace=True)
                JNJ_Date_exist = PrismDatabaseOperation.Prism_select('SELECT JNJ_Date FROM Backorder')
                Backorder = Backorder[['Material','Backorder_QTY','JNJ_Date']]
                if pd.isnull(Backorder.at[0,'Material'])==False:
                    Backorder.drop(Backorder[Backorder["Material"].isnull()].index,inplace=True)
                    Backorder.fillna(0,inplace=True)
                    if JNJ_Date_exist.empty:
                        insert_Backorder = PrismDatabaseOperation.Prism_insert('Backorder',Backorder)
                        if insert_Backorder == "error":
                            error_input.append(name)
                    else:
                        if jnj_date in JNJ_Date_exist['JNJ_Date'].unique():
                            already_input.append(name)
                            SQL_select = "SELECT * FROM Backorder WHERE JNJ_Date = '"+jnj_date+"';"
                            Backorder = PrismDatabaseOperation.Prism_select(SQL_select)
                        else:
                            insert_Backorder = PrismDatabaseOperation.Prism_insert('Backorder',Backorder)
                            if insert_Backorder == "error":
                                error_input.append(name)
                else:
                    Backorder.drop(Backorder[Backorder["Material"].isnull()].index,inplace=True)
                    lack_input.append(name)
            except:
                error_input.append(name)  
        else:
            tkinter.messagebox.showinfo("提示",name +"此文件非指定信息,请检查文件命名是否规范！")
    
    # 判断是否存在错误数据和缺失数据，显示到主界面
    ft = ("华文中宋",12)
#     Label()
    if error_input == [] and lack_input == []:
        # 判断是否已导入
        if already_input == []:
            Label(window,width=82,height=6,font=ft,justify="left",wraplength=800,anchor="nw",
                   bg="white",text="数据全部导入成功").place(x=360,y=280)
#             tkinter.messagebox.showinfo("提示","数据全部导入成功")
        else:
            Label(window,width=82,height=6,font=ft,justify="left",wraplength=800,anchor="nw",
                   bg="white",text="请注意:\n"+str(already_input)+"已存在数据库！"
                 ).place(x=360,y=280)
#             tkinter.messagebox.showinfo("提示",str(already_input)+"已存在数据库！")
    else:
        Label(window,width=82,height=6,font=ft,justify="left",wraplength=800,anchor="nw",
              bg="white",text="请注意:\n"+str(error_input)+"导入失败！\n"+str(lack_input)+
              "文件或数据缺失！").place(x=360,y=280)
#         tkinter.messagebox.showinfo("警告","请注意\n"+str(error_input)+"导入失败！\n"
#                                      +str(lack_input)+"文件或数据缺失！")
    # 各个数据进行merge
    merge_O_B = pd.merge(Outbound.loc[:,['JNJ_Date','Material','Outbound_QTY']],
                         Backorder,
                         on=['JNJ_Date','Material'],
                         how='outer')
    merge_O_B_P = pd.merge(merge_O_B,Putaway,
                           on=['JNJ_Date','Material'],
                           how='outer')
    merge_O_B_P_O = pd.merge(merge_O_B_P,Onhand,
                             on=['JNJ_Date','Material'],
                             how='outer')
    merge_O_B_P_O_I = pd.merge(merge_O_B_P_O,Intransit,on=['JNJ_Date','Material'],
                               how='outer')
    merge_all = merge_O_B_P_O_I[['JNJ_Date','Material','Outbound_QTY','Backorder_QTY',
                                'Putaway_QTY','Onhand_QTY','Intransit_QTY']]
#     merge_all.to_excel(r"merge_all.xlsx",index=False)
    merge_all.fillna(0,inplace=True)
    try:
        # 删除合计行
        merge_all.drop(index=merge_all[merge_all['Material']=='合计'].index,inplace=True)
        # 删除可能出现的重复值
        merge_all.drop_duplicates(inplace=True)
        # 删除因空表产生的0值
        merge_all.drop(index=merge_all[merge_all['Material']==0].index,inplace=True)
    except:
        pass
    merge_all['预测状态']='MTS'
    
#     merge_all.to_excel(r"merge_all.xlsx",index=False)
    # 获取当前Product Master
    ProductMaster = PrismDatabaseOperation.Prism_select('SELECT * FROM ProductMaster')
    # 合并
    merge_lack = pd.merge(ProductMaster,merge_all, on='Material',how='outer')
#     merge_lack.to_excel(r"merge_lack.xlsx",index=False)
#     to_open = merge_lack[merge_lack['预测状态']=='MTS'][merge_lack['FCST_state']=='关']
    to_add = merge_lack[merge_lack['预测状态']=='MTS'][merge_lack['FCST_state'].isnull()]
#     print(to_open)
#     print(to_add)
#     to_updata = to_open.append(to_add)
    to_updata = to_add.copy()
    to_updata = to_updata[list(ProductMaster.columns)]
    to_updata.rename(columns={'Material':'规格型号','GTS':'不含税单价','FCST_state':'预测状态'},
                     inplace=True)

    # 检测是code是否缺失和预测状态是否打开
    if to_updata.empty == False:
        MasterData.master_update_batch(to_updata)


# 需求数据预测
def forecast():
    """
    1.先计算预测模型所需的模型数值
    2.加权平均、乘以季节因子，计算出预测值，并计算共三个月的预测值（系统建议值），存入数据库
    3.需求修改，将修改过的值和原因存入数据库
    """
    
    # 计算预测模型值，并将计算出来的预测及预测之后3月数据显示到主窗口上
    def acl_FCSTModel():
        # 读取当月缺货、销售出库、季节因子数据，并且进行合并和计算，并导入相应的数据库表
        last_month = JNJ_Month(1)[0]
        Outbound = PrismDatabaseOperation.Prism_select("SELECT * FROM Outbound WHERE JNJ_Date='"+last_month+"';")
        Backorder = PrismDatabaseOperation.Prism_select("SELECT * FROM Backorder WHERE JNJ_Date='"+
                                          last_month+"';")
        B_O = pd.merge(Outbound,Backorder,how="outer",on=['Material','JNJ_Date'])
        SeasonFactor = PrismDatabaseOperation.Prism_select("SELECT * FROM SeasonFactor;")
        ActDemand = pd.merge(B_O,SeasonFactor,how="left",on=['JNJ_Date'])
        ActDemand['Backorder_QTY'].fillna(0,inplace=True)
        ActDemand['Outbound_QTY'].fillna(0,inplace=True)
        ActDemand['ActDemand_QTY'] = (ActDemand['Outbound_QTY']+
                                      ActDemand['Backorder_QTY'])/ActDemand['season_factor']
        ActDemand = ActDemand[['JNJ_Date','Material','ActDemand_QTY']]
        # 插入之前判断是否已有数据
        ActDemand_db = PrismDatabaseOperation.Prism_select("SELECT * FROM ActDemand WHERE JNJ_Date='"+
                                             last_month+"';")
        #         ActDemand.to_excel(r"ActDemand.xlsx")
        missing = []
        if ActDemand_db.empty:
            PrismDatabaseOperation.Prism_insert('ActDemand',ActDemand)
        else:
            ActDemand = ActDemand_db
            missing.append("模型数据")
        #             tkinter.messagebox.showinfo("提示","模型数据已存在！")

        month_12 = JNJ_Month(12)
        FCSTmodel = pd.DataFrame()
        for i in range(12):
            SQL_select = "SELECT * FROM ActDemand WHERE JNJ_Date='"+month_12[i]+"';"
            FCSTmodel = FCSTmodel.append(PrismDatabaseOperation.Prism_select(SQL_select))

        # 获取所有FCST_state为开的ProductMaster数据
        SQL_state = "SELECT Material,分类Level4 FROM ProductMaster WHERE FCST_state = 'MTS'"
        state_MTS  = PrismDatabaseOperation.Prism_select(SQL_state)

        # 链接数据，得到所需的计算12个月的数据
        acl_FCSTDemand = pd.merge(state_MTS,FCSTmodel,how="outer",on=['Material'])
        acl_FCSTDemand.fillna(0,inplace=True)
        #         print(acl_FCSTDemand)

        # 获取权值
        FCSTWeight = PrismDatabaseOperation.Prism_select("SELECT * FROM FCSTWeight")
        w1 = FCSTWeight['值'].iloc[0]
        w2 = FCSTWeight['值'].iloc[1]
        w3 = FCSTWeight['值'].iloc[2]

        # 数透，material为行，JNJ_Date为列
        acl_FCSTDemand_12 = acl_FCSTDemand.pivot_table(index='Material',columns='JNJ_Date')
        acl_FCSTDemand_12.fillna(0,inplace=True)

        # 换列名，方便直接取出相应月份的数据
        acl_FCSTDemand_12_col = []
        for i in range(len(acl_FCSTDemand_12.columns)):
            if type(acl_FCSTDemand_12.columns.values[i][1]) == str :
                 acl_FCSTDemand_12_col.append(acl_FCSTDemand_12.columns.values[i][1])
            else:
                acl_FCSTDemand_12_col.append(str(acl_FCSTDemand_12.columns.values[i][1]))
        #         print(acl_FCSTDemand_12_col)
        acl_FCSTDemand_12.columns = acl_FCSTDemand_12_col
        try:
            del acl_FCSTDemand_12["0"] # 当master里有而需求中没有就会出现不必要的0列，删除即可
        except:
            pass

        # 循环计算3次，以预测值预测后三月内容，nice！
        for i in range(3):
            # 获取相应月数据，并计算相应结果列
            model_3 = acl_FCSTDemand_12[acl_FCSTDemand_12.columns[-3:]]
            acl_model_3 = model_3[list(model_3.columns)[0]] # 求和3月值
            for i in range(1,len(model_3.columns)):
                acl_model_3 = acl_model_3 + model_3[list(model_3.columns)[i]]
            acl_model_3 = acl_model_3/3*w1

            model_6 = acl_FCSTDemand_12[acl_FCSTDemand_12.columns[-6:]]
            acl_model_6 = model_6[list(model_6.columns)[0]] # 求和6月值
            for i in range(1,len(model_6.columns)):
                acl_model_6 = acl_model_6 + model_6[list(model_6.columns)[i]]
            acl_model_6 = acl_model_6/6*w2

            model_12 = acl_FCSTDemand_12[acl_FCSTDemand_12.columns[-12:]]
            acl_model_12 = model_12[list(model_12.columns)[0]] # 求和12月值
            for i in range(1,len(model_12.columns)):
                acl_model_12 = acl_model_12 + model_12[list(model_12.columns)[i]]
            acl_model_12 = acl_model_12/12*w3
            FCSTModel_QTY = acl_model_3+acl_model_6+acl_model_12

            # 获取计算月的后一个月
            month_after_1 = datetime.datetime.strptime(acl_FCSTDemand_12.columns[-1],
                                                       '%Y%m')+relativedelta(months=+1)
            month_after_1 = month_after_1.strftime('%Y%m')
            # 将计算出来的最新一个月数据赋值给12月的计算数据
            acl_FCSTDemand_12[month_after_1] = FCSTModel_QTY

        # 取出预测的3个值，并排好列序，以防插入数据库时失败
        FCST_Demand = acl_FCSTDemand_12.iloc[:,-3:].reset_index() # 将数透格式转为df
        FCST_Demand.columns = ['Material','FCST_Demand1','FCST_Demand2','FCST_Demand3']
        FCST_Demand['JNJ_Date'] = len(FCST_Demand)*JNJ_Month(1)

        # 前面除过的季节因子，现在乘回来，需获取当月的JNJ_Date并获取后三个月的季节因子
        try:
            SeasonFactor = PrismDatabaseOperation.Prism_select("SELECT * FROM SeasonFactor")
            months_1 = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                        relativedelta(months=1)).strftime("%Y%m")
            months_2 = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                        relativedelta(months=2)).strftime("%Y%m")
            months_3 = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                        relativedelta(months=3)).strftime("%Y%m")
            SeasonFactor_1 = SeasonFactor[SeasonFactor['JNJ_Date']==months_1
                                         ]['season_factor'].iloc[0]
            SeasonFactor_2 = SeasonFactor[SeasonFactor['JNJ_Date']==months_2
                                         ]['season_factor'].iloc[0]
            SeasonFactor_3 = SeasonFactor[SeasonFactor['JNJ_Date']==months_3
                                         ]['season_factor'].iloc[0]
        except:
            tkinter.messagebox.showerror("错误","季节因子不存在，请维护季节因子数据")

        #         print(FCST_Demand[FCST_Demand['Material']=='W9932'])
        FCST_Demand['FCST_Demand1'] = FCST_Demand['FCST_Demand1']*SeasonFactor_1
        FCST_Demand['FCST_Demand2'] = FCST_Demand['FCST_Demand2']*SeasonFactor_2
        FCST_Demand['FCST_Demand3'] = FCST_Demand['FCST_Demand3']*SeasonFactor_3
        #         FCST_Demand.to_excel(r"FCST_Demand.xlsx")
        FCSTDemand = FCST_Demand[['JNJ_Date','Material','FCST_Demand1','FCST_Demand2',
                                  'FCST_Demand3']]
        FCSTDemand.fillna(0,inplace=True)
        # 全部取整
        FCSTDemand['FCST_Demand1'] = new_round(FCSTDemand['FCST_Demand1'],0)
        FCSTDemand['FCST_Demand2'] = new_round(FCSTDemand['FCST_Demand2'],0)
        FCSTDemand['FCST_Demand3'] = new_round(FCSTDemand['FCST_Demand3'],0)

        # 判断数据库中是否已存在，将计算出来的数值插入到数据库中
        SQL_select = "SELECT JNJ_Date FROM FCSTDemand"
        FCST_Demand_Date = PrismDatabaseOperation.Prism_select(SQL_select)['JNJ_Date']
        if JNJ_Month(1) not in FCST_Demand_Date.unique() or FCST_Demand_Date.empty:
            PrismDatabaseOperation.Prism_insert('FCSTDemand',FCSTDemand)
        else:
            SQL = "SELECT * FROM FCSTDemand WHERE JNJ_Date ='"+JNJ_Month(1)[0]+"';"
            FCSTDemand = PrismDatabaseOperation.Prism_select(SQL)
            missing.append("预测数据")
        #             tkinter.messagebox.showinfo("提示",JNJ_Month(1)[0]+"预测数据已存在！")

        # 断数据库中是否已存在，将FCST_Demand1存入预测需求调整表格中
        AdjustFCSTDemand = FCSTDemand[['JNJ_Date','Material','FCST_Demand1']]
        AdjustFCSTDemand["Remark"] = ""
        SQL_select = "SELECT JNJ_Date FROM AdjustFCSTDemand"
        Adjust_FCST_Demand_Date = PrismDatabaseOperation.Prism_select(SQL_select)['JNJ_Date']
        if JNJ_Month(1)  not in Adjust_FCST_Demand_Date.unique() or Adjust_FCST_Demand_Date.empty:
            PrismDatabaseOperation.Prism_insert('AdjustFCSTDemand',AdjustFCSTDemand)
        else:
            SQL = "SELECT * FROM AdjustFCSTDemand WHERE JNJ_Date ='"+JNJ_Month(1)[0]+"';"
            AdjustFCSTDemand = PrismDatabaseOperation.Prism_select(SQL)
            missing.append("已调整预测需求")
        #             tkinter.messagebox.showinfo("提示",JNJ_Month(1)[0]+"已调整预测需求已存在")        

        # 如果missing已存在，则提示哪些已存在
        if missing != []:
            tkinter.messagebox.showinfo("提示",str(missing)+"已存在!")

        # 将预测结果显示到主窗口，规格型号、产品家族、出库记录（6个月）、置信度（6个月数据，千分位）
        SQL_select = "SELECT [Material],[分类Level4] FROM ProductMaster WHERE FCST_state = 'MTS'"
        ProductMaster = PrismDatabaseOperation.Prism_select(SQL_select)

        # 获取出库记录（6个月）
        Outbound_QTY = pd.DataFrame()
        for i in range(6):
            SQL_select_outbound = "SELECT * FROM Outbound WHERE JNJ_Date = '"+             JNJ_Month(6)[i]+"';"
            Outbound_QTY = Outbound_QTY.append(PrismDatabaseOperation.Prism_select(SQL_select_outbound))
        Outbound_QTY = Outbound_QTY.pivot_table(index='Material',columns="JNJ_Date")
        Outbound_QTY_col = [] # 换列名，方便直接取出相应月份的数据
        for i in range(6):
            Outbound_QTY_col.append(Outbound_QTY.columns.values[i][1])
        Outbound_QTY.columns = Outbound_QTY_col
        Outbound_QTY = Outbound_QTY.reset_index() # 重置index

        # 合并并计算置信度
        P_O = pd.merge(ProductMaster,Outbound_QTY,how='left',on="Material")
        P_O_A = pd.merge(P_O,FCSTDemand,how='left',on="Material")
        FCST = pd.merge(P_O_A,AdjustFCSTDemand,how='left',on="Material")
        del FCST['JNJ_Date_x'],FCST['JNJ_Date_y']
        miu = np.mean(FCST[[JNJ_Month(6)[0],JNJ_Month(6)[1],JNJ_Month(6)[2],
                            JNJ_Month(6)[3],JNJ_Month(6)[4],JNJ_Month(6)[5]]].iloc[:],axis=1)
        sigma = np.std(FCST[[JNJ_Month(6)[0],JNJ_Month(6)[1],JNJ_Month(6)[2],
                             JNJ_Month(6)[3],JNJ_Month(6)[4],JNJ_Month(6)[5]]].iloc[:],axis=1)
        FCST.rename(columns={'Material':'规格型号','FCST_Demand1_x':'模型预测值',
                             'FCST_Demand1_y':'最终预测值','Remark':'修改原因'},
                    inplace=True)
        # # 历史数据为0，置信度应该为高
        # # 根据出库记录，计算置信度，按可信程度区分，切片替换,|x-μ|的距离判断即可
        FCST['置信度'] = ""
        FCST.loc[(abs(FCST['最终预测值']-miu)>abs(3*sigma)),'置信度'] = '低'
        FCST.loc[(abs(FCST['最终预测值']-miu)<=abs(3*sigma)) & 
                 (abs(FCST['最终预测值']-miu)>abs(2*sigma)),'置信度'] = '较低'
        FCST.loc[(abs(FCST['最终预测值']-miu)<=abs(2*sigma)) & 
                 (abs(FCST['最终预测值']-miu)>abs(sigma)),'置信度'] = '较高'
        FCST.loc[(abs(FCST['最终预测值']-miu)<=abs(sigma)),'置信度'] = "高"
        FCST.loc[((np.mean(FCST.iloc[:,-8:-2],axis=1))<=abs(0.1)),'置信度'] = "高"
        #         print(round(FCST['FCST_Demand1']))
        FCST.fillna(0,inplace=True)
        # 重新截取数据并排序
        FCST = FCST[["规格型号","分类Level4","模型预测值","最终预测值","修改原因","置信度",
                     FCST.columns[7],FCST.columns[6],FCST.columns[5],
                     FCST.columns[4],FCST.columns[3],FCST.columns[2]]]
        
        # 转换格式
        for i in FCST.columns:
            try:
                for j in range(len(FCST)):
                    FCST[i].iloc[j] = "{:,}".format(int(float(FCST[i].iloc[j])))
            except:
                FCST[i] = FCST[i].astype(str)
        #         print(FCST)
        #         FCST.to_excel(r'FCST.xlsx')


        # ------------------主界面-------------------#
        frame = Frame(window,height=600,width=1010,bg='WhiteSmoke')
        frame.place(x=270,y=120)
        columns = list(FCST.columns)
        
        # 设置样式
        style_head = ttk.Style()
        style_head.configure("MyStyle.Treeview.Heading",rowheight=50,font=("华文中宋",12))
        style_value = ttk.Style()
        style_value.configure("MyStyle.Treeview", rowheight=24)
        treeview = ttk.Treeview(frame, height=20, show="headings",selectmode="extended",
                                columns=columns,style='MyStyle.Treeview')

        # 添加滚动条
        # 竖向滚动条
        sb_y = ttk.Scrollbar(frame,command=treeview.yview)
        sb_y.config(command=treeview.yview)
        sb_y.place(in_=treeview,relx=1, rely=0,relheight=1)
        treeview.config(yscrollcommand=sb_y.set)
        # 横向滚动条
        sb_x = ttk.Scrollbar(frame,command=treeview.xview,orient="horizontal")
        sb_x.config(command=treeview.xview)
        sb_x.place(in_=treeview,relx=0, rely=1,relwidth=1)
        treeview.config(xscrollcommand=sb_x.set)
        treeview.place(x=0,y=50,relwidth=0.98)
        
        Label(frame,text='* 操作提示:双击相应数据可以进行编辑',font=('黑体',10),
              bg='WhiteSmoke').place(in_=treeview,x=10,y=530)

        # 表示列,不显示
        for i in range(len(FCST.columns)):
        #             print(str(FCST.columns[i]))
            treeview.column(str(FCST.columns[i]), width=90, anchor='center') 

        # 显示表头
        for i in range(len(FCST.columns)):
            treeview.heading(str(FCST.columns[i]), text=str(FCST.columns[i]))

        # 行交替颜色
        def fixed_map(option):# 重要！无此步骤则无法显示
            return [elm for elm in style.map("Treeview", query_opt=option)
                    if elm[:2] != ("!disabled", "!selected")]
        style = ttk.Style()
        style.map("Treeview",foreground=fixed_map("foreground"),background=fixed_map("background"))

        treeview.tag_configure('oddrow', background='LightGrey')
        treeview.tag_configure('evenrow', background='white')

        # 行坐标重排
        def odd_even_color():
            for index,row in enumerate(treeview.get_children()):
                if index % 2 == 0:
                    treeview.item(row,tags="evenrow")
                else:
                    treeview.item(row,tags="oddrow")
        
        # 插入数据，数字显示为千分位
        for i in range(len(FCST)):
            if i % 2 == 0:
                tag = "evenrow"
            else:
                tag = "oddrow"
            treeview.insert('', i, values=list(FCST.iloc[i,:]),tags=tag)

        # Treeview、列名、排列方式
        def treeview_sort_column(tv, col, reverse):  
            L = [(tv.set(k, col), k) for k in tv.get_children('')]
            try:
                for i in range(len(L)):
                    L[i] = (float(L[i][0].replace(',', '')),L[i][1])
            except:
                pass
            L.sort(reverse=reverse)  # 排序方式
            # 根据排序后索引移动
            for index, (val, k) in enumerate(L):
                tv.move(k, '', index)
            # 重写标题，使之成为再点倒序的标题
            tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))
            odd_even_color()

        # 绑定函数，使表头可排序
        for col in columns:
            treeview.heading(col, text=col, command=
                             lambda _col=col: treeview_sort_column(treeview, _col, False))

        # 双击进入编辑状态，弹出编辑界面
        def set_cell_value(event): 
            item_text = treeview.item(treeview.selection(), "values")
            # 编辑修改界面
            modify_FCST = Tk()
            modify_FCST.title('修改和记录')
            modify_FCST.geometry('500x300')

            #             # 输入值控制函数
            #             def entry_num():
            #                 if Entry_item.get().isdigit():
            #                     print("success")
            #                 else:
            #                     tkinter.messagebox.showerror("错误","请输入数字！")
                    
            # 修改前后
            Label(modify_FCST,text="型号规格：").place(x=100,y=10)
            Label(modify_FCST,text=str(item_text[0]),width=20).place(x=180,y=10)
            
            Label(modify_FCST,text="修改前：").place(x=100,y=50)
            Label(modify_FCST,text=str(item_text[3]),width=20).place(x=180,y=50)

            Label(modify_FCST,text="修改后：").place(x=100,y=90)
            Entry_item = Entry(modify_FCST,width=15)
            Entry_item.insert(0, item_text[3])
            #             Entry_item = Entry(modify_FCST,width=15,validate="focusout",
            #                                validatecommand=entry_num)
            Entry_item.place(x=180, y=90)

            Label(modify_FCST,text="修改原因：").place(x=100,y=130)
            Entry_remark = Entry(modify_FCST,width=15)
            Entry_remark.insert(0, item_text[4])
            Entry_remark.place(x=180, y=130)

            # 将编辑好的信息更新到界面和数据库中
            # 需求数据修改成实际需求后存入数据库，并之后的拆周计算以此为依据
            def save_edit():
                # 将编辑好的数字信息更新到数据库中,数量插入失败则提示
                try:
                    Entry_item_num = float(Entry_item.get().replace(',', ''))
                    item_text = treeview.item(treeview.selection(), "values")
                    #                     print(item_text,FCST.columns)
                    SQL_update_1 = "UPDATE AdjustFCSTDemand SET FCST_Demand1 = "+                     str(Entry_item_num)+" WHERE Material ='"+str(item_text[0])+                     "' AND JNJ_date = '"+str(FCST.columns[6])+"';" 
                    PrismDatabaseOperation.Prism_update(SQL_update_1)
                    # 将编辑好的数字信息更新到界面上
                    treeview.set(treeview.selection(), column=str(FCST.columns[3]),
                                 value=Entry_item.get())
                    # 将编辑好的备注信息更新到数据库中
                    SQL_update_2 =  "UPDATE AdjustFCSTDemand SET Remark = '"+                     Entry_remark.get()+"' WHERE Material ='"+str(item_text[0])+                     "' AND JNJ_date = '"+str(FCST.columns[6])+"';"
                    PrismDatabaseOperation.Prism_update(SQL_update_2)
                    # 将编辑好的字符信息更新到界面上
                    treeview.set(treeview.selection(),column=str(FCST.columns[4]),
                                 value=Entry_remark.get())
                    # 更新置信度判断
                    item_miu = np.mean([float(x.replace(',', '')) for x in item_text[6:11]])
                    item_sigma = np.std([float(x.replace(',', '')) for x in item_text[6:11]])

                    if abs(float(item_text[3].replace(',', ''))-item_miu) <= item_sigma:
                        treeview.set(treeview.selection(),column='置信度', value='高')
                    elif abs(float(item_text[3].replace(',', ''))-item_miu) <= 2*item_sigma:
                        treeview.set(treeview.selection(),column='置信度', value='较高')
                    elif abs(float(item_text[3].replace(',', ''))-item_miu) <= 3*item_sigma:
                        treeview.set(treeview.selection(),column='置信度', value='较低')
                    else:
                        treeview.set(treeview.selection(),column='置信度', value='低')
                except:
                    tkinter.messagebox.showerror("错误","请输入数字！")
                modify_FCST.destroy()
                
            btn_input = Button(modify_FCST, text='确认', width=10, command=save_edit)
            btn_input.place(x=120,y=180)

            # 取消输入
            def cancal_edit():
                modify_FCST.destroy()

            btn_cancal = Button(modify_FCST, text='取消', width=10, command=cancal_edit)
            btn_cancal.place(x=260,y=180)

            modify_FCST.mainloop()

        treeview.bind('<Double-1>', set_cell_value)
        
        # 搜索code功能
        Label(frame,text="筛选字段：",bg='WhiteSmoke',font=("黑体",12)).place(x=0,y=5)
        cbx = ttk.Combobox(frame,font=("黑体",12),width=10) #筛选字段
        comvalue = tkinter.StringVar()
        cbx["values"] = ["全局搜索"] + columns
        cbx.current(1)
        cbx.place(x=80,y=5)
        entry_search = Entry(frame,font=("黑体",12),width=12) # 筛选内容
        entry_search.insert(0, "请输入信息")
        entry_search.place(x=190,y=5)

        # 先清空表格，再插入数据，当字段选择为空、内容为空则显示全部
        def search_material():
            search_all = FCST.copy()
            # 必须转字符，否则无法全局搜索
            for i in search_all.columns:
                try:
                    search_all[i] = "{:,}".format(search_all[i].apply(int))
                except:
                    search_all[i] = search_all[i].apply(str)

            # 清空
            for item in treeview.get_children():
                treeview.delete(item)
            # 查找并插入数据
            if entry_search.get() != "":
                search_content = str(entry_search.get())
                # 全局搜索
                if cbx.get() == "全局搜索":
                    search_df = pd.DataFrame(columns=search_all.columns)
                    for i in range(len(search_all.columns)):
                        search_df = search_df.append(search_all[search_all[
                            search_all.columns[i]].str.contains(search_content)])                
                    search_df.drop_duplicates(subset=["规格型号"], keep='first',inplace=True)
                    #                 print(search_df)
                # 指定字段搜索
                else:
                    appoint = str(cbx.get())
                    search_df = search_all[search_all[appoint].str.contains(search_content)]
                    #                 print(search_df)
                # 插入表格
                for i in range(len(search_df)):
                    if i % 2 == 0:
                        tag = "evenrow"
                    else:
                        tag = "oddrow"
                    treeview.insert('',i,values=list(search_df.iloc[i,:]),tags=tag)
            # 若输入值为空则显示全部内容
            else:
                # 插入
                for i in range(len(search_all)):
                    if i % 2 == 0:
                        tag = "evenrow"
                    else:
                        tag = "oddrow"
                    treeview.insert('',i,values=list(search_all.iloc[i,:]),tags=tag)

        btn_search_material = Button(frame,text="查找",font=("黑体",10,'bold'),bg='slategrey',
                                     fg='white',width=9,height=1,borderwidth=5,
                                     command=search_material)
        btn_search_material.place(x=300,y=0)
        
        # 选择路径，输出保存
        def output_FCST():
            filename = tkinter.filedialog.asksaveasfilename()
            # 遍历获取所有数据，并生成df
            # 改变文本存储的数字
            t = treeview.get_children()
            a = list()
            for i in t:
                a.append(list(treeview.item(i,'values')))
            df_now = pd.DataFrame(a,columns=columns)
            
            # 指定列
            for i in range(0,len(df_now.columns)):
                try:
                    df_now[df_now.columns[i]] = df_now.loc[:,df_now.columns[i]].apply(
                        lambda x: float(x.replace(",", "")))
                except:
                    pass
            
            df_now.to_excel(filename+".xls",index=False)
        
        btn_output = Button(frame,text="导出结果",font=('黑体',10,'bold'),bg='slategrey',
                            fg='white',width=9,height=1,borderwidth=5,command=output_FCST)
        btn_output.place(x=890,y=0)

        def Three_content_1():
            # 二级目录按钮
            frame = Frame(window, height=200, width=185,bg='Gainsboro')# bg=""背景色透明      
            btn_replenishment = Button(frame,text='补货计划',font=s_1()[2],command=Replenishment,
                                       fg=s_1()[4],width=s_1()[6],height=s_1()[7],borderwidth=s_1()[5],
                                       compound=CENTER)
            btn_replenishment.place(x=s_1()[0],y=s_1()[1])
            btn_modify_rep = Button(frame,text='手动修改',font=s_1()[2],command=modify_Replenishment,
                                    fg=s_1()[4],width=s_1()[6],height=s_1()[7],borderwidth=s_1()[5],
                                    compound=CENTER)
            btn_modify_rep.place(x=s_1()[0],y=s_1()[1]+80)

            frame.place(x=80,y=100)

            # 覆盖按钮，替换颜色
            btn_One = Button(window,image=img_btn_update_png_1,borderwidth=0,height=45,width=45,
                             command=content.One_content)
            btn_One.place(x=17,y=100)
            btn_Two = Button(window,image=img_btn_FCST_png_1,borderwidth=0,height=45,width=45,
                             command=content.Two_content)
            btn_Two.place(x=17,y=180)
            btn_Three_2 = Button(window,image=img_btn_Rep_png_2,borderwidth=0,height=45,width=45,
                                 command=content.Three_content)
            btn_Three_2.place(x=17,y=260)
            btn_Four = Button(window,image=img_btn_Track_png_1,borderwidth=0,height=45,width=45,
                              command=content.Four_content)
            btn_Four.place(x=17,y=340)
            btn_Set = Button(window,image=img_btn_Set_png_1,borderwidth=0,height=45,width=45,
                             command=content.Five_content)
            btn_Set.place(x=17,y=640)

            # 显示提示文本
            CreateToolTip(btn_One, "更新数据")
            CreateToolTip(btn_Two, "预测需求")
            CreateToolTip(btn_Three_2, "补货拆周")
            CreateToolTip(btn_Four, "订单追踪")
            CreateToolTip(btn_Set, "设置")
            
            # 补货
            Replenishment()
        
        btn_replenishment = Button(window,text='生成补货计划',font=('黑体',12,'bold'),
                                   bg='slategrey',fg='white',width=15,height=1,
                                   borderwidth=5,command=Three_content_1,compound=CENTER)
        btn_replenishment.place(x=1100,y=75)
    
    acl_FCSTModel()


# Map and Biaos的图像
def MapeBias():
    # 读取outbound、ProductMaster、AdjustFCSTDemand数据并计算Map 和Biaos
    # 显示框
    frame = Frame(window,height=655,width=1015,bg='WhiteSmoke')
    frame.place(x=267,y=61)

    # 提示
    lb_title_f = Label(window,text="当前数据库最新至:"+JNJ_Month(1)[0],font=('黑体',12))
    lb_title_f.place(x=1000,y=25)

    # 标题
    lb_title = Label(window,text='Mape&Bias',font=('华文中宋',14),bg='WhiteSmoke',
                             fg='black',width=10,height=2)
    lb_title.place(x=280,y=10)

    lack = [] # 缺失
    # 出库数据
    last_month = JNJ_Month(1)[0]
    SQL_Outbound = "SELECT * FROM Outbound WHERE JNJ_Date='"+last_month+"';"
    Outbound = PrismDatabaseOperation.Prism_select(SQL_Outbound)
    Outbound = Outbound[["Material","Outbound_QTY"]]
    if Outbound.empty:
        lack.append("出库")

    # 调整后的需求数据
    SQL_AdjustFCSTDemand = "SELECT * FROM AdjustFCSTDemand WHERE JNJ_Date='"+JNJ_Month(2)[0]+"';"
    AdjustFCSTDemand = PrismDatabaseOperation.Prism_select(SQL_AdjustFCSTDemand)
    AdjustFCSTDemand = AdjustFCSTDemand[["Material","FCST_Demand1","Remark"]]
    if AdjustFCSTDemand.empty:
        lack.append("调整后的需求数据")

    # 缺货数据
    SQL_Backorder = "SELECT * FROM Backorder WHERE JNJ_Date='"+last_month+"';"
    Backorder = PrismDatabaseOperation.Prism_select(SQL_Backorder)
    Backorder = Backorder[["Material","Backorder_QTY"]]
    if Backorder.empty:
        lack.append("缺货")

    # 主数据
    SQL_ProductMaster = "SELECT Material,ABC,FCST_state From ProductMaster;"
    ProductMaster = PrismDatabaseOperation.Prism_select(SQL_ProductMaster)

    # 合并
    merge_1 = pd.merge(AdjustFCSTDemand,Outbound,on="Material",how="outer")
    merge_2 = pd.merge(merge_1,Backorder,on="Material",how="outer")
    merge_all = pd.merge(merge_2,ProductMaster,on="Material",how="outer")
    merge_all = merge_all[merge_all["FCST_state"]=="MTS"]
    merge_all.fillna(0,inplace=True)

    # 计算Gap 
    merge_all["Gap"] = merge_all["Outbound_QTY"]-merge_all["FCST_Demand1"]
    total_Mape = sum(abs(merge_all["Gap"]))/sum(merge_all["Outbound_QTY"])
    total_Bias = sum(merge_all["Gap"])/sum(merge_all["Outbound_QTY"])
    merge_all["Mape"] = abs(merge_all["Gap"])/merge_all["Outbound_QTY"]
    
    # mape剔除Outbound_QTY=0的情况
    Mape_df = merge_all[merge_all["Outbound_QTY"]!=0].sort_values(
        by=["Mape"],ascending=False)
    
    # 替换空值重命名、排序、截取所需信息、切换格式等
    merge_all.fillna(0,inplace=True)
    merge_all.rename(columns={"Material":"规格型号","Outbound_QTY":"实际出库",
                          "FCST_Demand1":"需求预测值","Remark":"调整原因",
                           "Backorder_QTY":"缺货"},inplace=True)
    merge_all = merge_all[["规格型号","ABC","实际出库","需求预测值","Gap","Mape","调整原因"]]
    merge_all.loc[np.isinf(merge_all["Mape"]),"Mape"]="无实际出库"
    for i in range(len(merge_all)):
        try:
            merge_all["Mape"].iloc[i] = "{:.1%}".format(merge_all["Mape"].iloc[i])
        except:
            pass

    # ***************主界面显示*****************#
    # Mape20
    Mape20 = Mape_df.sort_values(by=["Mape"],ascending=False)
    lt_Mape20 = Listbox(frame,selectmode=tkinter.EXTENDED,height=20)
    lt_Mape20.place(relx=0.02,rely=0.37)
    # 竖向滚动条
    sb_y = ttk.Scrollbar(frame,command=lt_Mape20.yview)
    sb_y.config(command=lt_Mape20.yview)
    sb_y.place(in_=lt_Mape20,relx=1, rely=0,relheight=1)
    lt_Mape20.config(yscrollcommand=sb_y.set)
    
    # 筛选
    Label(frame,text="MAPE TOP20",font=('黑体',15,'bold'),fg='white',bg='slategrey',relief=RIDGE,width=15).place(relx=0.02,rely=0.26)
    Label(frame,text="ABC:",bg='WhiteSmoke',font=('黑体',10)).place(relx=0.02,rely=0.32)    
    cbx_Mape20 = ttk.Combobox(frame,font=("黑体",11),width=5) #筛选字段
    comvalue = tkinter.StringVar()
    cbx_Mape20["values"] = list(Mape20["ABC"].unique())
    cbx_Mape20.current(0)
    cbx_Mape20.place(relx=0.06,rely=0.32)
    def search_Mape20():
        df_Mape = Mape_df.copy()
        df_Mape20 = df_Mape[df_Mape["ABC"]==cbx_Mape20.get()
                           ].sort_values(by=["Mape"],ascending=False).iloc[0:20,:]
        lt_Mape20.delete(0,END)
        if cbx_Mape20.get() != "":
            for i in df_Mape20["Material"]:
                lt_Mape20.insert(tkinter.END,i)
    
    search_Mape20()# 第一次默认值
    btn_Mape20 = Button(frame,text="筛选",font=("黑体",10,'bold'),bg='slategrey',
                                 fg='white',height=1,borderwidth=5,command=search_Mape20)
    btn_Mape20.place(relx=0.14,rely=0.31)
    
    # 总mape
    Label(frame,text="预测准确率回顾",font=('黑体',14,'bold'),bg='slategrey',relief=RIDGE,
          fg='white',width=15).place(relx=0.02,rely=0.05)
    Label(frame,text="Total Mape :",anchor='w',font=('Times New Roman',12,'bold'),bg='WhiteSmoke',fg='Black',
         width=10,height=2).place(relx=0.02,rely=0.1)
    lb_total_Mape = Label(frame,text='{:.1%}'.format(total_Mape),font=('Times New Roman',12,'bold'),bg='WhiteSmoke',fg='Brown',
                          anchor='e',width=7,height=2)
    lb_total_Mape.place(relx=0.1,rely=0.1)
    
    # 总bias
    Label(frame,text="Total Bias :",anchor='w',font=('Times New Roman',12,'bold'),bg='WhiteSmoke',fg='Black',
         width=10,height=2).place(relx=0.02,rely=0.17)
    lb_total_Bias = Label(frame,text='{:.1%}'.format(total_Bias),font=('Times New Roman',12,'bold'),bg='WhiteSmoke',fg='Brown',
                          anchor='e',width=7,height=2)
    lb_total_Bias.place(relx=0.1,rely=0.17)

    # *******Mape20清单界面******#
    Label(frame,text="MAPE 明细",font=('黑体',15,'bold'),fg='white',bg='slategrey',relief=RIDGE,
          width=73,anchor='center').place(relx=0.2,rely=0.05)
    columns = list(merge_all.columns)

    # 设置样式
    style_head = ttk.Style()
    style_head.configure("MyStyle.Treeview.Heading",rowheight=50,font=("华文中宋",12))
    style_value = ttk.Style()
    style_value.configure("MyStyle.Treeview", rowheight=24)
    treeview = ttk.Treeview(frame, height=20, show="headings",selectmode="extended",
                            columns=columns,style='MyStyle.Treeview')

    # 添加滚动条
    # 竖向滚动条
    sb_y = ttk.Scrollbar(frame,command=treeview.yview)
    sb_y.config(command=treeview.yview)
    sb_y.place(in_=treeview,relx=1, rely=0,relheight=1)
    treeview.config(yscrollcommand=sb_y.set)
    # 横向滚动条
    sb_x = ttk.Scrollbar(frame,command=treeview.xview,orient="horizontal")
    sb_x.config(command=treeview.xview)
    sb_x.place(in_=treeview,relx=0, rely=1,relwidth=1)
    treeview.config(xscrollcommand=sb_x.set)
    treeview.place(relx=0.2,rely=0.15,relwidth=0.78)

    # 表示列,不显示
    for i in range(0,len(merge_all.columns)):
        treeview.column(str(merge_all.columns[i]), width=80, anchor='center') 

    # 显示表头
    for i in range(len(merge_all.columns)):
        treeview.heading(str(merge_all.columns[i]), text=str(merge_all.columns[i]))

    # 行交替颜色
    def fixed_map(option):# 重要！无此步骤则无法显示
        return [elm for elm in style.map("Treeview", query_opt=option)
                if elm[:2] != ("!disabled", "!selected")]
    style = ttk.Style()
    style.map("Treeview",foreground=fixed_map("foreground"),background=fixed_map("background"))

    treeview.tag_configure('oddrow', background='LightGrey')
    treeview.tag_configure('evenrow', background='white')

    # 行坐标重排
    def odd_even_color():
        for index,row in enumerate(treeview.get_children()):
            if index % 2 == 0:
                treeview.item(row,tags="evenrow")
            else:
                treeview.item(row,tags="oddrow")

    # 插入数据，数字显示为千分位
    for i in range(len(merge_all)):
        if i % 2 == 0:
            tag = "evenrow"
        else:
            tag = "oddrow"
        treeview.insert('', i, 
                          values=(merge_all[merge_all.columns[0]].iloc[i],
                                  merge_all[merge_all.columns[1]].iloc[i],
                "{:,}".format(int(merge_all[merge_all.columns[2]].iloc[i])),
                "{:,}".format(int(merge_all[merge_all.columns[3]].iloc[i])),   
                                  merge_all[merge_all.columns[4]].iloc[i],
                                  merge_all[merge_all.columns[5]].iloc[i],
                                  merge_all[merge_all.columns[6]].iloc[i])
                       ,tags=tag)

    # Treeview、列名、排列方式
    def treeview_sort_column(tv, col, reverse):  
        L = [(tv.set(k, col), k) for k in tv.get_children('')]
        try:
            for i in range(len(L)):
                if L[i][0] == "无实际出库":
                    L[i] = (float(-1),L[i][1])
                elif "%" in L[i][0]:
                    L[i] = (float(L[i][0].replace('%', '')),L[i][1])
                else:
                    L[i] = (float(L[i][0].replace(',', '')),L[i][1])
        except:
            pass
        L.sort(reverse=reverse)  # 排序方式
        # 根据排序后索引移动
        for index, (val, k) in enumerate(L):
            tv.move(k, '', index)
        # 重写标题，使之成为再点倒序的标题
        tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))
        odd_even_color()

    # 绑定函数，使表头可排序
    for col in columns:
        treeview.heading(col, text=col, command=
                         lambda _col=col: treeview_sort_column(treeview, _col, False))

    # 选择路径，输出保存
    def output_plan():
        filename = tkinter.filedialog.asksaveasfilename()
        # 遍历获取所有数据，并生成df
        # 改变文本存储的数字
        t = treeview.get_children()
        a = list()
        for i in t:
            a.append(list(treeview.item(i,'values')))
        df_now = pd.DataFrame(a,columns=columns)
        # 指定列修改千分位为数字
        for i in range(0,len(df_now.columns)):
            try:
                df_now[df_now.columns[i]] = df_now.loc[:,df_now.columns[i]].apply(
                    lambda x: float(x.replace(",", "")))
            except:
                pass
        df_now.to_excel(filename+'.xls',index=False)

    btn_output = Button(frame,text="下载Mape&Bias",font=('黑体',10,'bold'),width=15,height=1,
                        bg='slategrey',fg='white',borderwidth=5,command=output_plan)
    btn_output.place(relx=0.85,rely=0.1)

    # 搜索功能
    Label(frame,text="筛选字段：",bg='WhiteSmoke',font=("黑体",10)).place(relx=0.2,rely=0.11)
    cbx = ttk.Combobox(frame,font=("黑体",11),width=10) #筛选字段
    comvalue = tkinter.StringVar()
    cbx["values"] = ["全局搜索"] + columns
    cbx.current(1)
    cbx.place(relx=0.28,rely=0.11)
    entry_search = Entry(frame,font=("黑体",11),width=12) # 筛选内容
    entry_search.insert(0, "请输入信息")
    entry_search.place(relx=0.4,rely=0.11)

    # 先清空表格，再插入数据，当字段选择为空、内容为空则显示全部
    def search_material():
        search_all = merge_all.copy()
        for i in search_all.columns:
            # 必须转字符，否则无法全局搜索
            try:
                search_all[i] = search_all[i].map(lambda x:format(int(x),','))
            except:
                search_all[i] = search_all[i].apply(str)

        # 清空
        for item in treeview.get_children():
            treeview.delete(item)
        # 查找并插入数据
        if entry_search.get() != "":
            search_content = str(entry_search.get())
            # 全局搜索
            if cbx.get() == "全局搜索":
                search_df = pd.DataFrame(columns=search_all.columns)
                for i in range(len(search_all.columns)):
                    search_df = search_df.append(search_all[search_all[
                        search_all.columns[i]].str.contains(search_content)])                
                search_df.drop_duplicates(subset=["规格型号"], keep='first',inplace=True)
                #                 print(search_df)
            # 指定字段搜索
            else:
                appoint = str(cbx.get())
                search_df = search_all[search_all[appoint].str.contains(search_content)]
                #                 print(search_df)
            # 插入表格
            for i in range(len(search_df)):
                if i % 2 == 0:
                    tag = "evenrow"
                else:
                    tag = "oddrow"
                treeview.insert('', i, values=list(search_df.iloc[i,:]),tags=tag)
        # 若输入值为空则显示全部内容
        else:
            # 插入
            for i in range(len(search_all)):
                if i % 2 == 0:
                    tag = "evenrow"
                else:
                    tag = "oddrow"
                treeview.insert('', i, values=list(search_all.iloc[i,:]),tags=tag)

    btn_search_material = Button(frame,text="查找",font=("黑体",10,'bold'),bg='slategrey',
                                 fg='white',width=9,height=1,borderwidth=5,
                                 command=search_material)
    btn_search_material.place(relx=0.52,rely=0.10)        
                
    # 点击跳转需求回顾数据
    def History_data():
        # 计算历史数据
        JNJ_12Month = []
        for i in range(12):
            JNJ_12Month.append(JNJ_Month(12)[i])
        JNJ_12Month = str(tuple(JNJ_12Month))
        
        lack = [] # 缺失

        # 主数据
        SQL_ProductMaster = "SELECT Material,ABC,FCST_state From ProductMaster;"
        ProductMaster = PrismDatabaseOperation.Prism_select(SQL_ProductMaster)

        # 出库数据
        SQL_Outbound = "SELECT * FROM Outbound WHERE JNJ_Date in "+JNJ_12Month+";"
        Outbound = PrismDatabaseOperation.Prism_select(SQL_Outbound)
        if Outbound.empty:
            lack.append("出库")

        # 缺货数据
        SQL_Backorder = "SELECT * FROM Backorder WHERE JNJ_Date in "+JNJ_12Month+";"
        Backorder = PrismDatabaseOperation.Prism_select(SQL_Backorder)
        if Backorder.empty:
            lack.append("缺货")

        # 调整后的需求数据
        JNJ_13Month = []
        for i in range(12):
            JNJ_13Month.append(JNJ_Month(13)[i])
        JNJ_13Month = str(tuple(JNJ_13Month))
        SQL_AdjustFCSTDemand = "SELECT * FROM AdjustFCSTDemand WHERE JNJ_Date in "+JNJ_13Month+";"
        AdjustFCSTDemand = PrismDatabaseOperation.Prism_select(SQL_AdjustFCSTDemand)
        # 为使相应数据合并，所有月份加1
        for i in range(len(AdjustFCSTDemand)):
            AdjustFCSTDemand.loc[i,"JNJ_Date"] = (datetime.datetime.strptime(
                AdjustFCSTDemand.loc[i,"JNJ_Date"],"%Y%m")+relativedelta(months=1)).strftime("%Y%m") 

        if AdjustFCSTDemand.empty:
            lack.append("调整后的需求数据")

        # 合并
        merge_O_B = pd.merge(Outbound,Backorder,how="outer",on=["Material","JNJ_Date"])
        merge_O_B_A = pd.merge(merge_O_B,AdjustFCSTDemand,how="outer",on=["Material","JNJ_Date"])
        merge_all = pd.merge(merge_O_B_A,ProductMaster,how="outer",on="Material")
        merge_all = merge_all[merge_all["FCST_state"]=="MTS"]

        merge_all["Outbound_QTY"].fillna(0,inplace=True)
        merge_all["Backorder_QTY"].fillna(0,inplace=True)
        merge_all["FCST_Demand1"].fillna(0,inplace=True)
        merge_all["Remark"].fillna("",inplace=True)
        merge_all["total"] = merge_all["Outbound_QTY"] + merge_all["Backorder_QTY"]

        # 聚合
        total = merge_all.groupby(merge_all["JNJ_Date"]).sum()
        total.reset_index(inplace=True)

        # ********##*主界面*********** #
        review = Tk()
        review.title("历史数据回顾")
        review.geometry('1500x650+20+80')
        review.configure(bg='WhiteSmoke')
        # set a figure
        fig = Figure(figsize=(10.1, 3.5), dpi=100,facecolor='WhiteSmoke')
        pic = fig.add_subplot(111)
        pic.set_facecolor("WhiteSmoke")
        x = total["JNJ_Date"]
        y1 = total["total"]
        y2 = total["Outbound_QTY"]
        y3 = total["Backorder_QTY"]
        y4 = total["FCST_Demand1"]
        plt.figure()
        pic.plot(x,y1,'b',label="实际出库+缺货",marker='o',
                 markerfacecolor='b',markersize=5,linewidth=2)
        pic.plot(x,y2,'r',label="实际出库",marker='o',linestyle='-.',
                 markerfacecolor='r',markersize=5,linewidth=3)
        pic.plot(x,y3,'y',label="缺货",marker='o',linestyle='--',
                 markerfacecolor='y',markersize=5,linewidth=2)
        pic.plot(x,y4,'g',label="预测需求",marker='o',linestyle=':',
                 markerfacecolor='g',markersize=5,linewidth=2)
        pic.legend()# 图例
        # pic.set_xticks(x,fontproperties = 'yahei', size = 10)

        # 坐标轴描述
        pic.set_ylabel("数量") 
        pic.set_xlabel("年月")
        # 标题
        pic.set_title("历史回顾（过去12个月）")
        # 数字标签
        for a, b in zip(x, y1):
            pic.text(a, b, b, ha='center', va='bottom', fontsize=10)

        # 创造并显示画布
        canvas = FigureCanvasTkAgg(fig,review)
        canvas.draw()
        canvas.get_tk_widget().place(x=5,rely=0.1,relheight=0.85,relwidth=0.65)

        # 定义并绑定键盘事件处理函数
        def on_key_event(event):
            print('you press %s' %event.key)
            key_press_handler(event, canvas, toolbar)
        canvas.mpl_connect('key_press_event', on_key_event)

        # 当筛选code的时候，只显示相应code的历史情况
        Label(review,text="筛选：",bg='WhiteSmoke',font=("黑体",12)).place(x=0,rely=0.06)
        entry_find = Entry(review,font=("黑体",12),width=15) # 筛选内容
        entry_find.insert(0, "请输入code")
        entry_find.place(x=60,rely=0.06)
        def updata_frame():
            # treeview数据同步变化
            search_material()
            # 截取相应数据
            search_df = merge_all.copy()
            search_code = entry_find.get()
            if search_code != "":
                search_info = search_df[search_df["Material"]==
                                        search_code].groupby(search_df["JNJ_Date"]).sum()
            else:
                search_info = search_df.groupby(search_df["JNJ_Date"]).sum()
            search_info.reset_index(inplace=True)

            # 覆盖画图
            fig = Figure(figsize=(10.1, 4), dpi=100)
            pic = fig.add_subplot(111)
            canvas = FigureCanvasTkAgg(fig,review)
            canvas.draw()
            canvas.get_tk_widget().place(x=5,rely=0.1,relheight=0.85,relwidth=0.65)

            # 工具栏
            toolbar = NavigationToolbar2Tk(canvas, review)
            toolbar.place(x=0,rely=0)
            # 定义并绑定键盘事件处理函数
            def on_key_event(event):
                print('you press %s' %event.key)
                key_press_handler(event, canvas, toolbar)
            canvas.mpl_connect('key_press_event', on_key_event)

            x = search_info["JNJ_Date"]
            y1 = search_info["total"]
            y2 = search_info["Outbound_QTY"]
            y3 = search_info["Backorder_QTY"]
            y4 = search_info["FCST_Demand1"]
            plt.figure()
            pic.plot(x,y1,'b',label="实际出库+缺货",marker='o',
                     markerfacecolor='b',markersize=5,linewidth=2)
            pic.plot(x,y2,'r',label="实际出库",marker='o',linestyle='-.',
                     markerfacecolor='r',markersize=5,linewidth=3)
            pic.plot(x,y3,'y',label="缺货",marker='o',linestyle='--',
                     markerfacecolor='y',markersize=5,linewidth=2)
            pic.plot(x,y4,'g',label="预测需求",marker='o',linestyle=':',
                     markerfacecolor='g',markersize=5,linewidth=2)
            pic.legend()# 图例

            # 坐标轴描述
            pic.set_ylabel("数量") 
            pic.set_xlabel("年月")
            pic.set_title("历史回顾（过去12个月）")
            # 数字标签
            for a, b in zip(x, y1):
                pic.text(a, b, b, ha='center', va='bottom', fontsize=10)

        btn_search_find = Button(review,text="查找",font=("黑体",10,'bold'),bg='slategrey',
                                 fg='white',width=9,height=1,borderwidth=5,command=updata_frame)
        btn_search_find.place(x=200,rely=0.05)

        # 工具栏
        toolbar = NavigationToolbar2Tk(canvas, review)
        toolbar.place(x=0,rely=0)

        # Remark界面显示
        # 抓取有Remark的
        Remark = merge_all[merge_all["Remark"]!=""]
        for i in range(len(Remark)):
            try:
                Remark["Mape"].iloc[i] = "{:.1%}".format(Remark["Mape"].iloc[i])
            except:
                pass
        # 重命名及排序
        Remark.rename(columns={"Material":"规格型号","Outbound_QTY":"实际出库",
                               "Backorder_QTY":"缺货","FCST_Demand1":"预测需求",
                               "JNJ_Date":"年月"},
                     inplace=True)
        Remark = Remark[["规格型号","ABC","年月","Remark","实际出库","缺货","预测需求"]]

        # *********界面显示*********#
        # 标题
        Label(review,text="Remark",font=("heiti",15,"bold")).place(relx=0.66,rely=0.02)
        columns = list(Remark.columns)

        # 设置样式
        style_head = ttk.Style()
        style_head.configure("MyStyle.Treeview.Heading",rowheight=50,font=("华文中宋",12))
        style_value = ttk.Style()
        style_value.configure("MyStyle.Treeview", rowheight=24)
        treeview = ttk.Treeview(review, height=20, show="headings",selectmode="extended",
                                columns=columns,style='MyStyle.Treeview')

        # 添加滚动条
        # 竖向滚动条
        sb_y = ttk.Scrollbar(review,command=treeview.yview)
        sb_y.config(command=treeview.yview)
        sb_y.place(in_=treeview,relx=1, rely=0,relheight=1)
        treeview.config(yscrollcommand=sb_y.set)
        # 横向滚动条
        sb_x = ttk.Scrollbar(review,command=treeview.xview,orient="horizontal")
        sb_x.config(command=treeview.xview)
        sb_x.place(in_=treeview,relx=0, rely=1,relwidth=1)
        treeview.config(xscrollcommand=sb_x.set)
        treeview.place(relx=0.66,rely=0.1,relwidth=0.33,relheight=0.85)

        # 表示列,不显示
        for i in range(0,len(Remark.columns)):
            treeview.column(str(Remark.columns[i]), width=10, anchor='center') 

        # 显示表头
        for i in range(len(Remark.columns)):
            treeview.heading(str(Remark.columns[i]), text=str(Remark.columns[i]))

        # 行交替颜色
        def fixed_map(option):# 重要！无此步骤则无法显示
            return [elm for elm in style.map("Treeview", query_opt=option)
                    if elm[:2] != ("!disabled", "!selected")]
        style = ttk.Style()
        style.map("Treeview",foreground=fixed_map("foreground"),background=fixed_map("background"))

        treeview.tag_configure('oddrow', background='LightGrey')
        treeview.tag_configure('evenrow', background='white')

        # 行坐标重排
        def odd_even_color():
            for index,row in enumerate(treeview.get_children()):
                if index % 2 == 0:
                    treeview.item(row,tags="evenrow")
                else:
                    treeview.item(row,tags="oddrow")

        # 插入数据，数字显示为千分位
        for i in range(len(Remark)):
            if i % 2 == 0:
                tag = "evenrow"
            else:
                tag = "oddrow"
            treeview.insert('', i, values=list(Remark.iloc[i,:]),tags=tag)

        # Treeview、列名、排列方式
        def treeview_sort_column(tv, col, reverse):  
            L = [(tv.set(k, col), k) for k in tv.get_children('')]
            try:
                for i in range(len(L)):
                    if L[i][0] == "无实际出库":
                        L[i] = (float(-1),L[i][1])
                    elif "%" in L[i][0]:
                        L[i] = (float(L[i][0].replace('%', '')),L[i][1])
                    else:
                        L[i] = (float(L[i][0].replace(',', '')),L[i][1])
            except:
                pass
            L.sort(reverse=reverse)  # 排序方式
            # 根据排序后索引移动
            for index, (val, k) in enumerate(L):
                tv.move(k, '', index)
            # 重写标题，使之成为再点倒序的标题
            tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))
            odd_even_color()

        # 绑定函数，使表头可排序
        for col in columns:
            treeview.heading(col, text=col, command=
                             lambda _col=col: treeview_sort_column(treeview, _col, False))

        # 选择路径，输出保存
        def output_plan():
            filename = tkinter.filedialog.asksaveasfilename()
            # 遍历获取所有数据，并生成df
            # 改变文本存储的数字
            t = treeview.get_children()
            a = list()
            for i in t:
                a.append(list(treeview.item(i,'values')))
            df_now = pd.DataFrame(a,columns=columns)
            # 指定列修改千分位为数字
            for i in range(0,len(df_now.columns)):
                try:
                    df_now[df_now.columns[i]] = df_now.loc[:,df_now.columns[i]].apply(
                        lambda x: float(x.replace(",", "")))
                except:
                    pass
            df_now.to_excel(filename+'.xls',index=False)

        btn_output = Button(review,text="下载Remark",font=('黑体',10,'bold'),width=10,height=1,
                            bg='slategrey',fg='white',borderwidth=5,command=output_plan)
        btn_output.place(relx=0.93,rely=0.05)

#         # 搜索功能
#         Label(review,text="筛选字段：",bg='WhiteSmoke',font=("黑体",12)).place(relx=0.66,rely=0.06)
#         cbx = ttk.Combobox(review,font=("黑体",11),width=10) #筛选字段
#         comvalue = tkinter.StringVar()
#         cbx["values"] = ["全局搜索"] + columns
#         cbx.current(1)
#         cbx.place(relx=0.72,rely=0.06)
#         entry_search = Entry(review,font=("黑体",11),width=12) # 筛选内容
#         entry_search.insert(0, "请输入信息")
#         entry_search.place(relx=0.8,rely=0.06)

        # 先清空表格，再插入数据，当字段选择为空、内容为空则显示全部
        def search_material():
            search_all = Remark.copy()
            for i in search_all.columns:
                try:
                    search_all[i] = search_all[i].apply(str)
                except:
                    pass

            # 清空
            for item in treeview.get_children():
                treeview.delete(item)
            # 查找并插入数据
            if entry_search.get() != "":
#                 search_content = str(entry_search.get())
                search_content = str(entry_find.get())
                # 全局搜索
                if cbx.get() == "全局搜索":
                    search_df = pd.DataFrame(columns=search_all.columns)
                    for i in range(len(search_all.columns)):
                        search_df = search_df.append(search_all[search_all[
                            search_all.columns[i]].str.contains(search_content)])                
                    search_df.drop_duplicates(subset=["规格型号"], keep='first',inplace=True)
        #                 print(search_df)
                # 指定字段搜索
                else:
                    appoint = str(cbx.get())
                    search_df = search_all[search_all[appoint].str.contains(search_content)]
        #                 print(search_df)
                # 插入表格
                for i in range(len(search_df)):
                    if i % 2 == 0:
                        tag = "evenrow"
                    else:
                        tag = "oddrow"
                    treeview.insert('', i, values=list(search_df.iloc[i,:]),tags=tag)
            # 若输入值为空则显示全部内容
            else:
                # 插入
                for i in range(len(search_all)):
                    if i % 2 == 0:
                        tag = "evenrow"
                    else:
                        tag = "oddrow"
                    treeview.insert('', i, values=list(search_all.iloc[i,:]),tags=tag)

#         btn_search_material = Button(review,text="查找",font=("黑体",10,'bold'),bg='slategrey',
#                                      fg='white',width=9,height=1,borderwidth=5,
#                                      command=search_material)
#         btn_search_material.place(relx=0.87,rely=0.05) 

        review.mainloop()
        
    btn_History = Button(frame,text=" >> 历史数据",font=('黑体',12,'bold'),bg='slategrey',fg='white',width=15,
                         borderwidth=5,command=History_data)
    btn_History.place(relx=0.84,rely=0.95)
    

# 补货、拆周模块
def Replenishment():
    # 显示框
    frame = Frame(window,height=655,width=1015,bg='WhiteSmoke')
    frame.place(x=267,y=61)
    lb_title_f = Label(window,text="当前数据库最新至:"+JNJ_Month(1)[0],font=('黑体',12))
    lb_title_f.place(x=1000,y=25)
    
    # 补货计划标题
    lb_title = Label(window,text='补货计划 ',font=('华文中宋',14),bg='WhiteSmoke',
                     fg='Black',width=10,height=2)
    lb_title.place(x=280,y=10)
    
    # 读取ProductMaster、Outbound（3个月）、Intransit、Safetystockday、AdjustFCSTDemand、Intransit、
    # Onhand、Putaway、Backorder，并返回计算结果
    def read_db():
        # 数据缺失
        lack = []
        # 主数据
        SQL_ProductMaster = "SELECT Material,GTS,ABC,MOQ From ProductMaster " +         "WHERE FCST_state='MTS';"
        ProductMaster = PrismDatabaseOperation.Prism_select(SQL_ProductMaster)
        ProductMaster.rename(columns={'ABC':"Class"},inplace=True)
        
        # 出库数据6个月，计算置信度
        SQL_Outbound = "SELECT * From Outbound WHERE JNJ_Date = '"+JNJ_Month(6)[0]+         "' OR JNJ_Date = '"+JNJ_Month(6)[1]+"' OR JNJ_Date = '"+JNJ_Month(6)[2]+         "' OR JNJ_Date = '"+JNJ_Month(6)[3]+"' OR JNJ_Date = '"+JNJ_Month(6)[4]         +"' OR JNJ_Date = '"+JNJ_Month(6)[5]+"';"
        Outbound = PrismDatabaseOperation.Prism_select(SQL_Outbound)
        Outbound = Outbound.pivot_table(index='Material',columns='JNJ_Date')
        Outbound_QTY_col = [] # 换列名，方便直接取出相应月份的数据
        for i in range(6):
            Outbound_QTY_col.append(Outbound.columns.values[i][1])
        Outbound.columns = Outbound_QTY_col
        Outbound = Outbound.reset_index()
        if Outbound.empty:
            lack.append("出库")
        
        # Intransit
        SQL_Intransit = "SELECT Material,Intransit_QTY From Intransit WHERE JNJ_Date = '"+         JNJ_Month(3)[2]+"';"
        Intransit = PrismDatabaseOperation.Prism_select(SQL_Intransit)
        if Intransit.empty:
            lack.append("在途")
            
        # Onhand_QTY
        SQL_Onhand = "SELECT Material,Onhand_QTY From Onhand WHERE JNJ_Date = '"+         JNJ_Month(3)[2]+"';"
        Onhand = PrismDatabaseOperation.Prism_select(SQL_Onhand)
        if Onhand.empty:
            lack.append("可发")
    
        # Putaway
        SQL_Putaway = "SELECT Material,Putaway_QTY From Putaway WHERE JNJ_Date = '"+         JNJ_Month(3)[2]+"';"
        Putaway = PrismDatabaseOperation.Prism_select(SQL_Putaway)
        if Putaway.empty:
            lack.append("预入库")
        
        # Backorder
        SQL_Backorder = "SELECT Material,Backorder_QTY From Backorder WHERE JNJ_Date = '"+         JNJ_Month(3)[2]+"';"
        Backorder = PrismDatabaseOperation.Prism_select(SQL_Backorder)
        if Backorder.empty:
            lack.append("缺货")
        
        # AdjustFCSTDemand
        SQL_Adjust = "SELECT Material,FCST_Demand1,Remark From AdjustFCSTDemand WHERE JNJ_Date = '"+         JNJ_Month(3)[2]+"';"
        AdjustFCSTDemand = PrismDatabaseOperation.Prism_select(SQL_Adjust)
        AdjustFCSTDemand['FCST_Demand1'] = new_round(AdjustFCSTDemand['FCST_Demand1'],0)
        if AdjustFCSTDemand.empty:
            lack.append("需求")
        
        # SafetyStockDay安全库存天数
        SQL_SafetyStockDay = "SELECT [Class],[Safetystock_Day] From SafetyStockDay;"
        SafetyStockDay = PrismDatabaseOperation.Prism_select(SQL_SafetyStockDay)
        if SafetyStockDay.empty:
            lack.append("安全库存天数")
        
#         if lack != []:
#             tkinter.messagebox.showwarning("警告",str(lack)+"数据缺失!")

        # 合并所需信息
        merge1 = pd.merge(AdjustFCSTDemand,Intransit,how="outer",on="Material")
        merge2 = pd.merge(merge1,Onhand,how="outer",on="Material")
        merge3 = pd.merge(merge2,Putaway,how="outer",on="Material")
        merge4 = pd.merge(merge3,Backorder,how="outer",on="Material")
        merge5 = pd.merge(merge4,Outbound,how="outer",on="Material")
        merge6 = pd.merge(ProductMaster,merge5,how="left",on="Material")
        merge_all = pd.merge(merge6,SafetyStockDay,how="left",on="Class")
        merge_all.drop_duplicates(keep='first',inplace=True) # 删除因merge产生的意外重复code
        merge_all.fillna(0,inplace=True)

        # 计算安全库存量(四舍五入取整)
        merge_all['Safetystock_QTY'] = 0
        month_1 = JNJ_Month(3)[0]
        month_2 = JNJ_Month(3)[1]
        month_3 = JNJ_Month(3)[2]
        for i in range(len(merge_all)):
            merge_all['Safetystock_QTY'].iloc[i] = new_round(merge_all['Safetystock_Day'].iloc[i]*
                                                             (merge_all[month_1].iloc[i]+
                                                              merge_all[month_2].iloc[i]+
                                                              merge_all[month_3].iloc[i])/90,0)
        # 计算置信度
        miu = np.mean(merge_all[[JNJ_Month(6)[0],JNJ_Month(6)[1],JNJ_Month(6)[2],
                          JNJ_Month(6)[3],JNJ_Month(6)[4],JNJ_Month(6)[5]]].iloc[:],axis=1)
        sigma = np.std(merge_all[[JNJ_Month(6)[0],JNJ_Month(6)[1],JNJ_Month(6)[2],
                          JNJ_Month(6)[3],JNJ_Month(6)[4],JNJ_Month(6)[5]]].iloc[:],axis=1)

        # # 历史数据为0，置信度应该为高
        # # 根据出库记录，计算置信度，按可信程度区分，切片替换,|x-μ|的距离判断即可
        merge_all['置信度'] = ""
        merge_all.loc[(abs(merge_all['FCST_Demand1']-miu)>abs(3*sigma)),'置信度'] = '低'
        merge_all.loc[(abs(merge_all['FCST_Demand1']-miu)<=abs(3*sigma)) & 
                 (abs(merge_all['FCST_Demand1']-miu)>abs(2*sigma)),'置信度'] = '较低'
        merge_all.loc[(abs(merge_all['FCST_Demand1']-miu)<=abs(2*sigma)) & 
                 (abs(merge_all['FCST_Demand1']-miu)>abs(sigma)),'置信度'] = '较高'
        merge_all.loc[(abs(merge_all['FCST_Demand1']-miu)<=abs(sigma)),'置信度'] = "高"
        merge_all.loc[((np.mean(merge_all[[JNJ_Month(6)[0],JNJ_Month(6)[1],JNJ_Month(6)[2],
                          JNJ_Month(6)[3],JNJ_Month(6)[4],JNJ_Month(6)[5]]].iloc[:],
                                axis=1))<=abs(0.1)),'置信度'] = "高"
        # 计算完成，删除不需要显示的列
        merge_all = merge_all.drop([JNJ_Month(6)[0],JNJ_Month(6)[1],JNJ_Month(6)[2]], axis=1)
        
        # 计算总库存
        merge_all['TotalINV_QTY'] =  merge_all['Intransit_QTY']+ merge_all['Onhand_QTY']+                              merge_all['Putaway_QTY']

        # 计算初始补货量
        merge_all['RepV1_QTY'] =  merge_all['FCST_Demand1']+merge_all['Safetystock_QTY']+         merge_all['Backorder_QTY']-merge_all['TotalINV_QTY']

        # 补货量凑整（二期：关于整托凑整）
        merge_all['Rep_QTY'] = 0
        for i in range(len(merge_all)):
            if merge_all['RepV1_QTY'].iloc[i] >= 300:
                merge_all['Rep_QTY'].iloc[i] = new_round(merge_all['RepV1_QTY'].iloc[i]/
                                                         merge_all['MOQ'].iloc[i],
                                                         0)*merge_all['MOQ'].iloc[i]
            elif merge_all['RepV1_QTY'].iloc[i] >= 50:
                merge_all['Rep_QTY'].iloc[i] = new_round(merge_all['RepV1_QTY'].iloc[i]/
                                                         merge_all['MOQ'].iloc[i],
                                                         0)*merge_all['MOQ'].iloc[i]
            elif merge_all['RepV1_QTY'].iloc[i] >= 2:
                merge_all['Rep_QTY'].iloc[i] = math.ceil(merge_all['RepV1_QTY'].iloc[i]/
                                                         merge_all['MOQ'].iloc[i]
                                                        )*merge_all['MOQ'].iloc[i]
            elif merge_all['RepV1_QTY'].iloc[i] >= 1:
                merge_all['Rep_QTY'].iloc[i] = 1
            else:
                merge_all['Rep_QTY'].iloc[i] = 0
        
        # 计算补货金额
        merge_all['Rep_value'] = merge_all['Rep_QTY']*merge_all['GTS']
        
        return merge_all
    
    # 计算模块：计算补货、并拆周，加入输出
    def acl_rep():
        merge_all = read_db()
        # 获取补货权值
        WeeklyPattern = PrismDatabaseOperation.Prism_select("SELECT * FROM WeeklyPattern")
        WK1 = WeeklyPattern[WeeklyPattern['week']=='WK1']['pattern'].iloc[0]
        WK2 = WeeklyPattern[WeeklyPattern['week']=='WK2']['pattern'].iloc[0]
        WK3 = WeeklyPattern[WeeklyPattern['week']=='WK3']['pattern'].iloc[0]
        WK4 = WeeklyPattern[WeeklyPattern['week']=='WK4']['pattern'].iloc[0]
        merge_all['W1'] = 0
        merge_all['W2'] = 0
        merge_all['W3'] = 0
        merge_all['W4'] = 0
        
        # 拆周
        for i in range(len(merge_all)):
            # 第一、二、三周
            if merge_all['Rep_QTY'].iloc[i] <= 300:
                if merge_all['TotalINV_QTY'].iloc[i] <= merge_all['FCST_Demand1'].iloc[i]*(WK1+WK2):
                    merge_all['W1'].iloc[i] = merge_all['Rep_QTY'].iloc[i]
                elif (merge_all['TotalINV_QTY'].iloc[i]>merge_all['FCST_Demand1'].iloc[i]*(WK1+WK2)
                      and merge_all['TotalINV_QTY'].iloc[i] < merge_all['FCST_Demand1'].iloc[i]):
                    merge_all['W2'].iloc[i] = merge_all['Rep_QTY'].iloc[i]
                else:
                    merge_all['W3'].iloc[i] = merge_all['Rep_QTY'].iloc[i]
            elif merge_all['Rep_QTY'].iloc[i] < 900:
                if merge_all['TotalINV_QTY'].iloc[i]<=merge_all['FCST_Demand1'].iloc[i]*(WK1+WK2):
                    merge_all['W1'].iloc[i] = new_round(merge_all['Rep_QTY'].iloc[i]*(WK1+WK2)/
                                                        merge_all['MOQ'].iloc[i],0
                                                       )*merge_all['MOQ'].iloc[i]
                    merge_all['W3'].iloc[i] = new_round(merge_all['Rep_QTY'].iloc[i]*(1-WK1-WK2)/
                                                        merge_all['MOQ'].iloc[i],0
                                                       )*merge_all['MOQ'].iloc[i]
                else:
                    merge_all['W2'].iloc[i] = new_round(merge_all['Rep_QTY'].iloc[i]*(WK1+WK2)/
                                                        merge_all['MOQ'].iloc[i],0
                                                       )*merge_all['MOQ'].iloc[i]
                    merge_all['W3'].iloc[i] = new_round(merge_all['Rep_QTY'].iloc[i]*(1-WK1-WK2)/
                                                        merge_all['MOQ'].iloc[i],0
                                                       )*merge_all['MOQ'].iloc[i]
            else:
                merge_all['W1'].iloc[i] = new_round(merge_all['Rep_QTY'].iloc[i]*WK1/
                                                    merge_all['MOQ'].iloc[i],0
                                                   )*merge_all['MOQ'].iloc[i]
                merge_all['W2'].iloc[i] = new_round(merge_all['Rep_QTY'].iloc[i]*WK2/
                                                    merge_all['MOQ'].iloc[i],0
                                                   )*merge_all['MOQ'].iloc[i]
                merge_all['W3'].iloc[i] = new_round(merge_all['Rep_QTY'].iloc[i]*WK3/
                                                    merge_all['MOQ'].iloc[i],0
                                                   )*merge_all['MOQ'].iloc[i]  

            # 第四周
            merge_all['W4'].iloc[i] = (merge_all['Rep_QTY'].iloc[i]-merge_all['W1'].iloc[i]-
                                           merge_all['W2'].iloc[i]-merge_all['W3'].iloc[i])
            if merge_all['W4'].iloc[i] < 0 :
                merge_all['W4'].iloc[i] = 0
                merge_all['W3'].iloc[i] = merge_all['W3'].iloc[i] - merge_all['W4'].iloc[i]
        
        # 输出结果
        rep_result = merge_all[['Material','FCST_Demand1','置信度','Rep_QTY','W1','W2','W3','W4',
                                JNJ_Month(3)[0],JNJ_Month(3)[1],JNJ_Month(3)[2],'Backorder_QTY',
                                'TotalINV_QTY','Onhand_QTY','Putaway_QTY','Intransit_QTY',
                                'Safetystock_QTY']]
        rep_result.rename(columns={'Material':'规格型号','FCST_Demand1':'二级需求',
                                   'Backorder_QTY':'缺货量','TotalINV_QTY':'总库存',
                                   'Onhand_QTY':'可发量','Putaway_QTY':'预入库',
                                   'Intransit_QTY':'在途','Safetystock_QTY':'安全库存',
                                   'Rep_QTY':'补货量'},inplace=True)
        
        
        # 显示计算金额并进行比较，若有上个月数据，则显示，若无则显示为0
        next_month = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                  relativedelta(months=1)).strftime("%Y%m") # 下个月时间 
        if next_month in list(PrismDatabaseOperation.Prism_select("SELECT JNJ_Date FROM OrderTarget")['JNJ_Date']):
            SQL_select_target = "SELECT * FROM OrderTarget WHERE JNJ_Date = '"+next_month+"';"
            Target_amount = PrismDatabaseOperation.Prism_select(SQL_select_target)['order_target'].iloc[0]
            lb_target_amount = Label(frame,text=re_round(Target_amount),anchor="e",
                                     font=('黑体',15),width=15,height=2,bg='WhiteSmoke')
            lb_target_amount.place(x=820,y=3)
        else:
            Target_amount = 0
            lb_target_amount = Label(frame,text=re_round(Target_amount),anchor="e",
                                     font=('黑体',15),width=15,height=2,bg='WhiteSmoke')
            lb_target_amount.place(x=820,y=3)
        
        Label(frame,text="月度指标金额：",font=('华文中宋',14),width=12,height=1,bg='WhiteSmoke'
             ,anchor="w").place(x=675,y=10)
        Label(frame,text="当前补货金额：",font=('华文中宋',14),width=12,height=1,bg='WhiteSmoke'
             ,anchor="w").place(x=675,y=55)
        
        # 判断颜色，大于指标红色，小于指标绿色
        if int(Target_amount) > int(sum(merge_all['Rep_value'])):
            lb_amount = Label(frame,text=re_round(sum(merge_all['Rep_value'])),anchor="e",
                              font=('黑体',15),width=15,height=2,fg='green',bg='WhiteSmoke')
            lb_amount.place(x=820,y=46)
        else:
            lb_amount = Label(frame,text=re_round(sum(merge_all['Rep_value'])),anchor="e",
                              font=('黑体',15),width=15,height=2,fg='red',bg='WhiteSmoke')
            lb_amount.place(x=820,y=46)
        
        
        # 将得到的计算结果展示在界面
        columns = list(rep_result.columns)

        # 设置样式
        style_head = ttk.Style()
        style_head.configure("MyStyle.Treeview.Heading",rowheight=50,font=("华文中宋",12))
        style_value = ttk.Style()
        style_value.configure("MyStyle.Treeview", rowheight=24)
        treeview = ttk.Treeview(frame, height=21, show="headings",selectmode="extended",
                                columns=columns,style='MyStyle.Treeview')

        # 添加滚动条
        # 竖向滚动条
        sb_y = ttk.Scrollbar(frame,command=treeview.yview)
        sb_y.config(command=treeview.yview)
        sb_y.place(in_=treeview,relx=1, rely=0,relheight=1)
        treeview.config(yscrollcommand=sb_y.set)
        # 横向滚动条
        sb_x = ttk.Scrollbar(frame,command=treeview.xview,orient="horizontal")
        sb_x.config(command=treeview.xview)
        sb_x.place(in_=treeview,relx=0, rely=1,relwidth=1)
        treeview.config(xscrollcommand=sb_x.set)
        treeview.place(x=0,y=100,relwidth=0.98)

        # 表示列,不显示
#         treeview.column(str(rep_result.columns[0]), width=80, anchor='center')
        
        for i in range(0,len(rep_result.columns)):
#             print(str(FCST.columns[i]))
            treeview.column(str(rep_result.columns[i]), width=100, anchor='center') 

        # 显示表头
        for i in range(len(rep_result.columns)):
            treeview.heading(str(rep_result.columns[i]), text=str(rep_result.columns[i]))

        # 行交替颜色
        def fixed_map(option):# 重要！无此步骤则无法显示
            return [elm for elm in style.map("Treeview", query_opt=option)
                    if elm[:2] != ("!disabled", "!selected")]
        style = ttk.Style()
        style.map("Treeview",foreground=fixed_map("foreground"),background=fixed_map("background"))

        treeview.tag_configure('oddrow', background='LightGrey')
        treeview.tag_configure('evenrow', background='white')

        # 行坐标重排
        def odd_even_color():
            for index,row in enumerate(treeview.get_children()):
                if index % 2 == 0:
                    treeview.item(row,tags="evenrow")
                else:
                    treeview.item(row,tags="oddrow")
            
        # 插入数据，数字显示为千分位
        for i in range(len(rep_result)):
            if i % 2 == 0:
                tag = "evenrow"
            else:
                tag = "oddrow"
            treeview.insert('', i, 
                            values=(rep_result[rep_result.columns[0]].iloc[i],
                                    "{:,}".format(int(rep_result[rep_result.columns[1]].iloc[i])),
                                    rep_result[rep_result.columns[2]].iloc[i],
                                    "{:,}".format(int(rep_result[rep_result.columns[3]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[4]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[5]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[6]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[7]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[8]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[9]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[10]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[11]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[12]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[13]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[14]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[15]].iloc[i])),
                                    "{:,}".format(int(rep_result[rep_result.columns[16]].iloc[i])))
                           ,tags=tag)

        # Treeview、列名、排列方式
        def treeview_sort_column(tv, col, reverse):  
            L = [(tv.set(k, col), k) for k in tv.get_children('')]
            try:
                for i in range(len(L)):
                    L[i] = (float(L[i][0].replace(',', '')),L[i][1])
            except:
                pass
            L.sort(reverse=reverse)  # 排序方式
            # 根据排序后索引移动
            for index, (val, k) in enumerate(L):
                tv.move(k, '', index)
            # 重写标题，使之成为再点倒序的标题
            tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))
            odd_even_color()

        # 绑定函数，使表头可排序
        for col in columns:
            treeview.heading(col, text=col, command=
                             lambda _col=col: treeview_sort_column(treeview, _col, False))
        
        # 保存
        def save_rep_plan():
            # 将计算好的数据保存进数据库、如有，则删除，再添加
            t = treeview.get_children() # 获取当前显示的表格所有数据
            a = list()
            for i in t:
                a.append(list(treeview.item(i,'values')))
            Rep_plan = pd.DataFrame(a,columns=columns)
            Rep_plan["JNJ_Date"] = next_month
            Rep_plan = Rep_plan[["JNJ_Date","规格型号","W1","W2","W3","W4"]]
            Rep_plan.rename(columns={"规格型号":"Material"},inplace=True)
            Rep_plan = Rep_plan.melt(id_vars =['Material','JNJ_Date'], var_name = 'week_No', 
                                     value_name = 'RepWeek_QTY')
            # 改变千分位字符为数字
            Rep_plan['RepWeek_QTY'] = Rep_plan.loc[:,'RepWeek_QTY'].apply(
                lambda x: float(x.replace(",", "")))
    #         print(Rep_plan)
            # 判断当前数据库中的月份
        
            if next_month in list(PrismDatabaseOperation.Prism_select("SELECT JNJ_Date FROM RepPlan")['JNJ_Date']):
                SQL_delete = "DELETE FROM RepPlan WHERE JNJ_Date = '"+next_month+"';"
                PrismDatabaseOperation.Prism_delete(SQL_delete)
                PrismDatabaseOperation.Prism_insert('RepPlan',Rep_plan)
                # 将微调后的保存至数据库的调整计划表
                SQL_delete_Adjust = "DELETE FROM AdjustRepPlan WHERE JNJ_Date = '"+next_month+"';"
                PrismDatabaseOperation.Prism_delete(SQL_delete_Adjust)
                Adj_Rep_plan = Rep_plan.copy()
                Adj_Rep_plan["Rep_Remark"] = ""
                PrismDatabaseOperation.Prism_insert('AdjustRepPlan',Rep_plan)
            else:
                PrismDatabaseOperation.Prism_insert('AdjustRepPlan',Rep_plan)
                PrismDatabaseOperation.Prism_insert('RepPlan',Rep_plan)
                
            tkinter.messagebox.showinfo("提示","保存成功！")
        
        btn_save = Button(frame,text="调整数据暂存",font=('黑体',12,'bold'),bg='slategrey',
                          fg='white',height=1,borderwidth=5,command=save_rep_plan)
        btn_save.place(x=330,y=15)
        CreateToolTip(btn_save, "所有调整已完成，保存信息至数据库中！（注：会覆盖已有数据）")
        
        # 搜索功能
        Label(frame,text="筛选字段：",bg='WhiteSmoke',font=("黑体",12)).place(x=5,y=65)
        cbx = ttk.Combobox(frame,font=("黑体",11),width=10) #筛选字段
        comvalue = tkinter.StringVar()
        cbx["values"] = ["全局搜索"] + columns
        cbx.current(1)
        cbx.place(x=85,y=65)
        entry_search = Entry(frame,font=("黑体",11),width=12) # 筛选内容
        entry_search.insert(0, "请输入信息")
        entry_search.place(x=195,y=65)
    #     CreateToolTip(entry_search, "请注意大小写输入！")
        # 先清空表格，再插入数据，当字段选择为空、内容为空则显示全部
        def search_material():
            search_all = rep_result.copy()
            for i in search_all.columns:
                search_all[i] = search_all[i].apply(str)# 必须转字符，否则无法全局搜索

            # 清空
            for item in treeview.get_children():
                treeview.delete(item)
            # 查找并插入数据
            if entry_search.get() != "":
                search_content = str(entry_search.get())
                # 全局搜索
                if cbx.get() == "全局搜索":
                    search_df = pd.DataFrame(columns=search_all.columns)
                    for i in range(len(search_all.columns)):
                        search_df = search_df.append(search_all[search_all[
                            search_all.columns[i]].str.contains(search_content)])                
                    search_df.drop_duplicates(subset=["规格型号"], keep='first',inplace=True)
    #                 print(search_df)
                # 指定字段搜索
                else:
                    appoint = str(cbx.get())
                    search_df = search_all[search_all[appoint].str.contains(search_content)]
    #                 print(search_df)
                # 插入表格
                for i in range(len(search_df)):
                    if i % 2 == 0:
                        tag = "evenrow"
                    else:
                        tag = "oddrow"
                    treeview.insert('', i, values=list(search_df.iloc[i,:]),tags=tag)
            # 若输入值为空则显示全部内容
            else:
                # 插入
                for i in range(len(search_all)):
                    if i % 2 == 0:
                        tag = "evenrow"
                    else:
                        tag = "oddrow"
                    treeview.insert('', i, values=list(search_all.iloc[i,:]),tags=tag)

        btn_search_material = Button(frame,text="查找",font=("黑体",10,'bold'),bg='slategrey',
                                     fg='white',width=9,height=1,borderwidth=5,
                                     command=search_material)
        btn_search_material.place(x=320,y=60)
        
        # 输入预测月指标
        def order_target(event):
            entry_input = Entry(frame,font=('黑体',15),width=15)
            entry_input.place(x=820,y=10)
            # 插入数据库中的目标金额
            SQL_OrderTarget = "SELECT * FROM OrderTarget WHERE JNJ_Date = '"+next_month+"';"
            OrderTarget_amount = PrismDatabaseOperation.Prism_select(SQL_OrderTarget)
            entry_input.insert(0,OrderTarget_amount["order_target"].iloc[0])
            
            # 确认输入函数，并将输入数据更新至数据库
            def input_target_amount():
                # 保存数据库并以最后的标准为主，方法：先插入下个月信息，再更新输入信息
                # 如果已存在于数据库，则直接更新，否则新增空的过度变量再更新
                if next_month in list(
                    PrismDatabaseOperation.Prism_select("SELECT JNJ_Date FROM OrderTarget")["JNJ_Date"]):
                    # 覆盖输入数据
                    SQL_delete = "DELETE FROM OrderTarget WHERE JNJ_Date = '"+next_month+"';"
                    PrismDatabaseOperation.Prism_delete(SQL_delete)
                    OrderTarget = pd.DataFrame(data={'JNJ_Date':[next_month],
                                                     'order_target':[float(entry_input.get())]})
                    PrismDatabaseOperation.Prism_insert('OrderTarget',OrderTarget)
                else:
                    OrderTarget = pd.DataFrame(data={'JNJ_Date':[next_month],
                                                     'order_target':[float(entry_input.get())]})
                    PrismDatabaseOperation.Prism_insert('OrderTarget',OrderTarget)


                SQL_select_target = "SELECT * FROM OrderTarget WHERE JNJ_Date = '"+next_month+"';"
                Target_amount = PrismDatabaseOperation.Prism_select(SQL_select_target)['order_target'].iloc[0]
                lb_target_amount = Label(frame,text=re_round(Target_amount),anchor="e",
                                         font=('黑体',15),width=15,height=2,bg='WhiteSmoke')
                lb_target_amount.place(x=820,y=3)
                # 双击提示
                CreateToolTip(lb_target_amount, "双击此处即可编辑")
                lb_target_amount.bind('<Double-1>',order_target)
                # 输入指标后变色
                if int(Target_amount) > int(sum(merge_all['Rep_value'])):
                    lb_amount = Label(frame,text=re_round(sum(merge_all['Rep_value'])),anchor="e",
                                      font=('黑体',15),width=15,height=2,fg='green',bg='WhiteSmoke')
                    lb_amount.place(x=820,y=46)
                else:
                    lb_amount = Label(frame,text=re_round(sum(merge_all['Rep_value'])),anchor="e",
                                      font=('黑体',15),width=15,height=2,fg='red',bg='WhiteSmoke')
                    lb_amount.place(x=820,y=46)

                btn_input_cancel.destroy()
                btn_input_target.destroy()
                entry_input.destroy()

            btn_input_target = Button(frame,text="OK",command=input_target_amount)
            btn_input_target.place(x=820,y=30)
            # 输入取消
            def input_cancel():
                btn_input_cancel.destroy()
                btn_input_target.destroy()
                entry_input.destroy()

            btn_input_cancel = Button(frame,text="No",command=input_cancel)
            btn_input_cancel.place(x=920,y=30)

        lb_target_amount.bind('<Double-1>',order_target)
        CreateToolTip(lb_target_amount, "双击此处即可编辑")
    #     btn_order =Button(frame,text='输入预测月指标',font=('黑体',12,'bold'),bg='slategrey',
    #                       fg='white',width=15,height=1,borderwidth=5,compound=CENTER,
    #                       command=order_target)
    #     btn_order.place(x=10,y=15)
        
    acl_rep() # 执行计算，并方便更新覆盖
    
    # 修改安全库存天数
    def modify_safeday():
        # 获取安全库存天数信息
        SQL_SafetyStockDay = "SELECT [Class],[Safetystock_Day] From SafetyStockDay;"
        SafetyStockDay = PrismDatabaseOperation.Prism_select(SQL_SafetyStockDay)
        # 修改安全库存天数界面
        modify_SafetyStockDay = Tk()
        modify_SafetyStockDay.title('安全库存天数修改')
        modify_SafetyStockDay.geometry('400x450')

        columns = list(SafetyStockDay.columns)
        modify_treeview = ttk.Treeview(modify_SafetyStockDay, height=16, show="headings", columns=columns)
        modify_treeview.place(x=20,y=20)

        # 定义表头
        modify_treeview.column(str(SafetyStockDay.columns[0]), width=50, anchor='center') 
        modify_treeview.column(str(SafetyStockDay.columns[1]), width=70, anchor='center') 

        # 显示表头
        for i in range(len(SafetyStockDay.columns)):
            modify_treeview.heading(str(SafetyStockDay.columns[i]), text=str(SafetyStockDay.columns[i]))

        # 插入数据
        for i in range(len(SafetyStockDay)):
            modify_treeview.insert('', i, values=(SafetyStockDay[SafetyStockDay.columns[0]].iloc[i],
                                                  SafetyStockDay[SafetyStockDay.columns[1]].iloc[i],))

        # 合并输入数据
        def set_value(event): 
            # 获取鼠标所选item
            for item in modify_treeview.selection():
                item_text = modify_treeview.item(item, "values")

            Label_select = Label(modify_SafetyStockDay,text=str(item_text[1]),width=10)
            Label_select.place(x=230, y=50)
            entryedit = Entry(modify_SafetyStockDay,width=10)
            entryedit.place(x=230, y=100)
            # 将编辑好的信息更新到数据库中
            def save_edit():
                # 获取
                modify_treeview.set(item, column=str(SafetyStockDay.columns[1]), 
                                    value=entryedit.get())                    
                entryedit.destroy()
                btn_input.destroy()
                btn_cancal.destroy()
                Label_select.destroy()

            btn_input = Button(modify_SafetyStockDay, text='OK', width=10, command=save_edit)
            btn_input.place(x=150,y=150)

            # 取消输入
            def cancal_edit():
                entryedit.destroy()
                btn_input.destroy()
                btn_cancal.destroy()
                Label_select.destroy()

            btn_cancal = Button(modify_SafetyStockDay, text='Cancal', width=10, command=cancal_edit)
            btn_cancal.place(x=280,y=150)

        # 触发双击事件
        modify_treeview.bind('<Double-1>', set_value)

        # 显示文本数据
        Label(modify_SafetyStockDay,text="修改前：").place(x=150,y=50)
        Label(modify_SafetyStockDay,text="修改后：").place(x=150,y=100)

        # 将编辑好的信息更新到数据库中
        def db_update():
            # 先删除，再直接插入~更简单~
            item_text = modify_treeview.item(modify_treeview.selection(), "values")
            SQL_delete = "DELETE FROM SafetyStockDay ;"
            PrismDatabaseOperation.Prism_delete(SQL_delete)
            # 遍历获取所有数据，并生成df
            t = modify_treeview.get_children()
            a = list()
            for i in t:
                a.append(list(modify_treeview.item(i,'values')))
            df_now = pd.DataFrame(a,columns=columns)
            PrismDatabaseOperation.Prism_insert('SafetyStockDay',df_now)
#             modify_SafetyStockDay.destroy()
            # 再次计算,更新金额,覆盖界面
            acl_rep()

        Label(modify_SafetyStockDay,text="Tips：安全库存天数必须为数字！", fg='red').place(x=20,y=400)
        btn_confirm = Button(modify_SafetyStockDay,text="确认修改",font=("黑体",12,'bold'),
                             bg='slategrey',fg='white',height=1,borderwidth=5,
                             width=15,command=db_update)
        btn_confirm.place(x=200,y=300)
        mainloop()
        
    btn_modify_safeday = Button(frame,text='安全库存天数',font=('黑体',12,'bold'),
                                bg='slategrey',fg='white',command=modify_safeday,width=15,
                                height=1,borderwidth=5,compound=CENTER)
    btn_modify_safeday.place(x=170,y=15)   
    
    # 修改WeeklyParttern
    def modify_WeeklyPattern():
        # 获取安全库存天数信息
        SQL_SafetyStockDay = "SELECT [Class],[Safetystock_Day] From SafetyStockDay;"
        SafetyStockDay = PrismDatabaseOperation.Prism_select(SQL_SafetyStockDay)
        
        # 获取数据库中的补货权值
        WeeklyPattern = PrismDatabaseOperation.Prism_select("SELECT * FROM WeeklyPattern")
        
        # 修改窗体
        modify_WeeklyPattern = Tk()
        modify_WeeklyPattern.title('拆周比例修改')
        modify_WeeklyPattern.geometry('400x350')
        
        columns = list(WeeklyPattern.columns)
        modify_treeview = ttk.Treeview(modify_WeeklyPattern, height=7, show="headings",
                                       columns=columns)
        modify_treeview.place(x=20,y=20)
        
        # 定义表头
        modify_treeview.column(str(WeeklyPattern.columns[0]), width=50, anchor='center') 
        modify_treeview.column(str(WeeklyPattern.columns[1]), width=70, anchor='center') 

        # 显示表头
        for i in range(len(WeeklyPattern.columns)):
            modify_treeview.heading(str(WeeklyPattern.columns[i]),
                                    text=str(WeeklyPattern.columns[i]))

        # 插入数据
        for i in range(len(WeeklyPattern)):
            modify_treeview.insert('', i, values=list(WeeklyPattern.iloc[i,:]))
    
        # 合并输入数据
        def set_value(event): 
            # 获取鼠标所选item
            for item in modify_treeview.selection():
                item_text = modify_treeview.item(item, "values")

            Label_select = Label(modify_WeeklyPattern,text=str(item_text[1]),width=10)
            Label_select.place(x=230, y=50)
            entryedit = Entry(modify_WeeklyPattern,width=10)
            entryedit.place(x=230, y=100)
            # 将编辑好的信息更新到数据库中
            def save_edit():
                # 获取
                modify_treeview.set(item, column=str(WeeklyPattern.columns[1]), 
                                    value=entryedit.get())
                entryedit.destroy()
                btn_input.destroy()
                btn_cancal.destroy()
                Label_select.destroy()

            btn_input = Button(modify_WeeklyPattern, text='OK', width=10, command=save_edit)
            btn_input.place(x=150,y=150)

            # 取消输入
            def cancal_edit():
                entryedit.destroy()
                btn_input.destroy()
                btn_cancal.destroy()
                Label_select.destroy()

            btn_cancal = Button(modify_WeeklyPattern, text='Cancal', width=10, command=cancal_edit)
            btn_cancal.place(x=280,y=150)

        # 触发双击事件
        modify_treeview.bind('<Double-1>', set_value)

        # 显示文本数据
        Label(modify_WeeklyPattern,text="修改前：").place(x=150,y=50)
        Label(modify_WeeklyPattern,text="修改后：").place(x=150,y=100)
        
        # 将编辑好的信息更新到数据库中
        def db_update():
            # 先删除，再直接插入~更简单~
            item_text = modify_treeview.item(modify_treeview.selection(), "values")
            SQL_delete = "DELETE FROM WeeklyPattern ;"
            PrismDatabaseOperation.Prism_delete(SQL_delete)
            # 遍历获取所有数据，并生成df
            t = modify_treeview.get_children()
            a = list()
            for i in t:
                a.append(list(modify_treeview.item(i,'values')))
            df_now = pd.DataFrame(a,columns=columns)
            PrismDatabaseOperation.Prism_insert('WeeklyPattern',df_now)
            # 再次计算,更新金额,覆盖界面
            acl_rep()
        
        # 显示文本数据
        Label(modify_WeeklyPattern,text="Tips：修改拆周比例必须为数字（非百分比格式）！",
              fg='red').place(x=20,y=200)
        btn_confirm = Button(modify_WeeklyPattern,text="确认修改",width=15,command=db_update)
        btn_confirm.place(x=200,y=300)
        mainloop()
        
    btn_modify_weekly = Button(frame,text="二级订货节奏",font=('黑体',12,'bold'),bg='slategrey',
                               fg='white',width=15,height=1,borderwidth=5,compound=CENTER,
                               command=modify_WeeklyPattern)
    btn_modify_weekly.place(x=10,y=15)
    
    
# 修改补货计划
def modify_Replenishment():
    # 显示框
    frame = Frame(window,height=655,width=1015,bg='WhiteSmoke')
    frame.place(x=267,y=61)
    # 补货计划标题
    lb_title = Label(window,text='手动修改',font=('华文中宋',14),bg='WhiteSmoke',
                     fg='Black',width=10,height=2)
    lb_title.place(x=280,y=10)
    lb_title_f = Label(window,text="当前数据库最新至:"+JNJ_Month(1)[0],font=('黑体',12))
    lb_title_f.place(x=1000,y=25)
    # 获取修改计划数据
    def read_adjust_db():
        # 数据缺失
        lack = []
        # 主数据
        SQL_ProductMaster = "SELECT Material,GTS,ABC,MOQ From ProductMaster " +         "WHERE FCST_state='MTS';"
        ProductMaster = PrismDatabaseOperation.Prism_select(SQL_ProductMaster)
        ProductMaster.rename(columns={'ABC':"Class"},inplace=True)
        # 出库数据6个月，计算置信度
        SQL_Outbound = "SELECT * From Outbound WHERE JNJ_Date = '"+JNJ_Month(6)[0]+         "' OR JNJ_Date = '"+JNJ_Month(6)[1]+"' OR JNJ_Date = '"+JNJ_Month(6)[2]+         "' OR JNJ_Date = '"+JNJ_Month(6)[3]+"' OR JNJ_Date = '"+JNJ_Month(6)[4]         +"' OR JNJ_Date = '"+JNJ_Month(6)[5]+"';"
        Outbound = PrismDatabaseOperation.Prism_select(SQL_Outbound)
        Outbound = Outbound.pivot_table(index='Material',columns='JNJ_Date')
        Outbound_QTY_col = [] # 换列名，方便直接取出相应月份的数据
        for i in range(6):
            Outbound_QTY_col.append(Outbound.columns.values[i][1])
        Outbound.columns = Outbound_QTY_col
        Outbound = Outbound.reset_index()
        if Outbound.empty:
            lack.append("出库")

        # Intransit
        SQL_Intransit = "SELECT Material,Intransit_QTY From Intransit WHERE JNJ_Date = '"+         JNJ_Month(3)[2]+"';"
        Intransit = PrismDatabaseOperation.Prism_select(SQL_Intransit)
        if Intransit.empty:
            lack.append("在途")

        # Onhand_QTY
        SQL_Onhand = "SELECT Material,Onhand_QTY From Onhand WHERE JNJ_Date = '"+         JNJ_Month(3)[2]+"';"
        Onhand = PrismDatabaseOperation.Prism_select(SQL_Onhand)
        if Onhand.empty:
            lack.append("可发")

        # Putaway
        SQL_Putaway = "SELECT Material,Putaway_QTY From Putaway WHERE JNJ_Date = '"+         JNJ_Month(3)[2]+"';"
        Putaway = PrismDatabaseOperation.Prism_select(SQL_Putaway)
        if Putaway.empty:
            lack.append("预入库")

        # Backorder
        SQL_Backorder = "SELECT Material,Backorder_QTY From Backorder WHERE JNJ_Date = '"+         JNJ_Month(3)[2]+"';"
        Backorder = PrismDatabaseOperation.Prism_select(SQL_Backorder)
        if Backorder.empty:
            lack.append("缺货")

        # AdjustFCSTDemand
        SQL_Adjust = "SELECT Material,FCST_Demand1,Remark From AdjustFCSTDemand WHERE JNJ_Date = '"+         JNJ_Month(3)[2]+"';"
        AdjustFCSTDemand = PrismDatabaseOperation.Prism_select(SQL_Adjust)
        AdjustFCSTDemand['FCST_Demand1'] = new_round(AdjustFCSTDemand['FCST_Demand1'],0)
        if AdjustFCSTDemand.empty:
            lack.append("需求")

        # SafetyStockDay安全库存天数
        SQL_SafetyStockDay = "SELECT [Class],[Safetystock_Day] From SafetyStockDay;"
        SafetyStockDay = PrismDatabaseOperation.Prism_select(SQL_SafetyStockDay)
        if SafetyStockDay.empty:
            lack.append("安全库存天数")

        # 读取补货计划
        next_month = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                  relativedelta(months=1)).strftime("%Y%m") # 下个月时间 
        SQL_AdjustRepPlan = "SELECT * From AdjustRepPlan WHERE JNJ_Date= "+next_month+";"
        AdjustRepPlan = PrismDatabaseOperation.Prism_select(SQL_AdjustRepPlan)
        if AdjustRepPlan.empty:
            lack.append("补货计划")
        # Rep_Remark存储在W1中
        AdjustRepPlan_m = AdjustRepPlan[["Material","JNJ_Date","week_No","Rep_Remark"]]
        AdjustRepPlan = pd.pivot_table(AdjustRepPlan,index=["Material","JNJ_Date"],
                                       columns=["week_No"],
                                       values=["RepWeek_QTY"])
        AdjustRepPlan_col = []
        for i in range(len(AdjustRepPlan.columns)):
            if type(AdjustRepPlan.columns.values[i][1]) == str :
                AdjustRepPlan_col.append(AdjustRepPlan.columns.values[i][1])
            else:
                AdjustRepPlan_col.append(str(AdjustRepPlan.columns.values[i][1]))
        AdjustRepPlan.columns = AdjustRepPlan_col
        AdjustRepPlan = AdjustRepPlan.reset_index()# 重排索引

        # 合并备注数据
        AdjustRepPlan = pd.merge(AdjustRepPlan,AdjustRepPlan_m,on=["Material","JNJ_Date"],
                                 how="left")
        # 切片，并只保留含W1的数据
        AdjustRepPlan = AdjustRepPlan[AdjustRepPlan["week_No"]=="W1"]
        AdjustRepPlan = AdjustRepPlan[["Material","W1","W2","W3","W4","Rep_Remark"]]

#         if lack != []:
#             tkinter.messagebox.showwarning("警告",str(lack)+"数据缺失!")

        # 合并所需信息
        merge1 = pd.merge(AdjustFCSTDemand,Intransit,how="outer",on="Material")
        merge2 = pd.merge(merge1,Onhand,how="outer",on="Material")
        merge3 = pd.merge(merge2,Putaway,how="outer",on="Material")
        merge4 = pd.merge(merge3,Backorder,how="outer",on="Material")
        merge5 = pd.merge(merge4,Outbound,how="outer",on="Material")
        merge6 = pd.merge(ProductMaster,merge5,how="left",on="Material")
        merge7 = pd.merge(merge6,AdjustRepPlan,how="left",on="Material")
        merge_all = pd.merge(merge7,SafetyStockDay,how="left",on="Class")
        merge_all.drop_duplicates(keep='first',inplace=True) # 删除因merge产生的意外重复code
        merge_all["Rep_Remark"].fillna("-",inplace=True)
        merge_all.fillna(0,inplace=True)

        # 计算安全库存量(四舍五入取整)
        merge_all['Safetystock_QTY'] = 0
        month_1 = JNJ_Month(3)[0]
        month_2 = JNJ_Month(3)[1]
        month_3 = JNJ_Month(3)[2]
        for i in range(len(merge_all)):
            merge_all['Safetystock_QTY'].iloc[i] = new_round(merge_all['Safetystock_Day'].iloc[i]*
                                                             (merge_all[month_1].iloc[i]+
                                                              merge_all[month_2].iloc[i]+
                                                              merge_all[month_3].iloc[i])/90,0)
        # 计算置信度
        miu = np.mean(merge_all[[JNJ_Month(6)[0],JNJ_Month(6)[1],JNJ_Month(6)[2],
                          JNJ_Month(6)[3],JNJ_Month(6)[4],JNJ_Month(6)[5]]].iloc[:],axis=1)
        sigma = np.std(merge_all[[JNJ_Month(6)[0],JNJ_Month(6)[1],JNJ_Month(6)[2],
                          JNJ_Month(6)[3],JNJ_Month(6)[4],JNJ_Month(6)[5]]].iloc[:],axis=1)

        # # 历史数据为0，置信度应该为高
        # # 根据出库记录，计算置信度，按可信程度区分，切片替换,|x-μ|的距离判断即可
        merge_all['置信度'] = ""
        merge_all.loc[(abs(merge_all['FCST_Demand1']-miu)>abs(3*sigma)),'置信度'] = '低'
        merge_all.loc[(abs(merge_all['FCST_Demand1']-miu)<=abs(3*sigma)) & 
                 (abs(merge_all['FCST_Demand1']-miu)>abs(2*sigma)),'置信度'] = '较低'
        merge_all.loc[(abs(merge_all['FCST_Demand1']-miu)<=abs(2*sigma)) & 
                 (abs(merge_all['FCST_Demand1']-miu)>abs(sigma)),'置信度'] = '较高'
        merge_all.loc[(abs(merge_all['FCST_Demand1']-miu)<=abs(sigma)),'置信度'] = "高"
        merge_all.loc[((np.mean(merge_all[[JNJ_Month(6)[0],JNJ_Month(6)[1],JNJ_Month(6)[2],
                          JNJ_Month(6)[3],JNJ_Month(6)[4],JNJ_Month(6)[5]]].iloc[:],
                                axis=1))<=abs(0.1)),'置信度'] = "高"
        # 计算完成，删除不需要显示的列
        merge_all = merge_all.drop([JNJ_Month(6)[0],JNJ_Month(6)[1],JNJ_Month(6)[2]], axis=1)

        # 计算总库存
        merge_all['TotalINV_QTY'] =  merge_all['Intransit_QTY']+ merge_all['Onhand_QTY']+                              merge_all['Putaway_QTY']

        # 计算初始补货量
        merge_all['Rep_QTY'] =  merge_all['W1']+merge_all['W2']+merge_all['W3']+merge_all['W4']

        # 计算补货金额
        merge_all['Rep_value'] = merge_all['Rep_QTY']*merge_all['GTS']

        # 重命名及排序

        rep_result = merge_all[['Material','FCST_Demand1','置信度','Rep_QTY','W1','W2','W3','W4',
                                'Rep_Remark',JNJ_Month(3)[0],JNJ_Month(3)[1],JNJ_Month(3)[2],
                                'Backorder_QTY','TotalINV_QTY','Onhand_QTY','Putaway_QTY',
                                'Intransit_QTY','Safetystock_QTY']]
        rep_result.rename(columns={'Material':'规格型号','FCST_Demand1':'二级需求',
                                   'Backorder_QTY':'缺货量','TotalINV_QTY':'总库存',
                                   'Onhand_QTY':'可发量','Putaway_QTY':'预入库',
                                   'Intransit_QTY':'在途','Safetystock_QTY':'安全库存',
                                   'Rep_QTY':'补货量','Rep_Remark':'备注'},inplace=True)
        
        # **************** 主界面 ***************#
        next_month = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                  relativedelta(months=1)).strftime("%Y%m") # 下个月时间 
        if next_month in list(PrismDatabaseOperation.Prism_select("SELECT JNJ_Date FROM OrderTarget")['JNJ_Date']):
            SQL_select_target = "SELECT * FROM OrderTarget WHERE JNJ_Date = '"+next_month+"';"
            Target_amount = PrismDatabaseOperation.Prism_select(SQL_select_target)['order_target'].iloc[0]
            lb_target_amount = Label(frame,text=re_round(Target_amount),font=('黑体',15),
                                     width=15,height=2,anchor="e",bg='WhiteSmoke')
            lb_target_amount.place(x=820,y=3)
        else:
            Target_amount = 0
            lb_target_amount = Label(frame,text=re_round(0),font=('黑体',15),width=15,height=2,
                                     anchor="e",bg='WhiteSmoke')
            lb_target_amount.place(x=820,y=3)
        
        Label(frame,text="月度指标金额：",font=('华文中宋',14),width=12,height=1,bg='WhiteSmoke'
             ).place(x=675,y=10)
        Label(frame,text="当前补货金额：",font=('华文中宋',14),width=12,height=1,bg='WhiteSmoke'
             ).place(x=675,y=55)
        
        # 判断颜色，大于指标红色，小于指标绿色
        if int(Target_amount) > int(sum(merge_all['Rep_value'])):
            lb_amount = Label(frame,text=re_round(sum(merge_all['Rep_value'])),anchor="e",
                              font=('黑体',15),width=15,height=2,fg='green',bg='WhiteSmoke')
            lb_amount.place(x=820,y=46)
        else:
            lb_amount = Label(frame,text=re_round(sum(merge_all['Rep_value'])),anchor="e",
                              font=('黑体',15),width=15,height=2,fg='red',bg='WhiteSmoke')
            lb_amount.place(x=820,y=46)
        
        
        # 将得到的计算结果展示在界面
        columns = list(rep_result.columns)

        # 设置样式
        style_head = ttk.Style()
        style_head.configure("MyStyle.Treeview.Heading",rowheight=50,font=("华文中宋",12))
        style_value = ttk.Style()
        style_value.configure("MyStyle.Treeview", rowheight=24)
        treeview = ttk.Treeview(frame, height=20, show="headings",selectmode="extended",
                                columns=columns,style='MyStyle.Treeview')

        # 添加滚动条
        # 竖向滚动条
        sb_y = ttk.Scrollbar(frame,command=treeview.yview)
        sb_y.config(command=treeview.yview)
        sb_y.place(in_=treeview,relx=1, rely=0,relheight=1)
        treeview.config(yscrollcommand=sb_y.set)
        # 横向滚动条
        sb_x = ttk.Scrollbar(frame,command=treeview.xview,orient="horizontal")
        sb_x.config(command=treeview.xview)
        sb_x.place(in_=treeview,relx=0, rely=1,relwidth=1)
        treeview.config(xscrollcommand=sb_x.set)
        treeview.place(x=0,y=100,relwidth=0.98)

        # Tips
        Label(frame,text="* 操作提示：双击相应数据可以且仅能修改每周的补货量及其备注",bg='WhiteSmoke',
          font=("黑体",10)).place(in_=treeview,x=0,y=530)
        # 表示列,不显示
        for i in range(0,len(rep_result.columns)):
            treeview.column(str(rep_result.columns[i]), width=100, anchor='center') 

        # 显示表头
        for i in range(len(rep_result.columns)):
            treeview.heading(str(rep_result.columns[i]), text=str(rep_result.columns[i]))

        # 行交替颜色
        def fixed_map(option):# 重要！无此步骤则无法显示
            return [elm for elm in style.map("Treeview", query_opt=option)
                    if elm[:2] != ("!disabled", "!selected")]
        style = ttk.Style()
        style.map("Treeview",foreground=fixed_map("foreground"),background=fixed_map("background"))

        treeview.tag_configure('oddrow', background='LightGrey')
        treeview.tag_configure('evenrow', background='white')

        # 行坐标重排
        def odd_even_color():
            for index,row in enumerate(treeview.get_children()):
                if index % 2 == 0:
                    treeview.item(row,tags="evenrow")
                else:
                    treeview.item(row,tags="oddrow")
            
        # 插入数据，数字显示为千分位
        for i in range(len(rep_result)):
            if i % 2 == 0:
                tag = "evenrow"
            else:
                tag = "oddrow"
            treeview.insert('', i, 
                              values=(rep_result[rep_result.columns[0]].iloc[i],
                    "{:,}".format(int(rep_result[rep_result.columns[1]].iloc[i])),
                                      rep_result[rep_result.columns[2]].iloc[i],
                    "{:,}".format(int(rep_result[rep_result.columns[3]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[4]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[5]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[6]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[7]].iloc[i])),
                                      rep_result[rep_result.columns[8]].iloc[i],
                    "{:,}".format(int(rep_result[rep_result.columns[9]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[10]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[11]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[12]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[13]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[14]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[15]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[16]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[17]].iloc[i])))
                           ,tags=tag)

        # Treeview、列名、排列方式
        def treeview_sort_column(tv, col, reverse):  
            L = [(tv.set(k, col), k) for k in tv.get_children('')]
            try:
                for i in range(len(L)):
                    L[i] = (float(L[i][0].replace(',', '')),L[i][1])
            except:
                pass
            L.sort(reverse=reverse)  # 排序方式
            # 根据排序后索引移动
            for index, (val, k) in enumerate(L):
                tv.move(k, '', index)
            # 重写标题，使之成为再点倒序的标题
            tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))
            odd_even_color()

        # 绑定函数，使表头可排序
        for col in columns:
            treeview.heading(col, text=col, command=
                             lambda _col=col: treeview_sort_column(treeview, _col, False))
        
        # 选择路径，输出保存
        def output_plan():
            filename = tkinter.filedialog.asksaveasfilename()
            # 遍历获取所有数据，并生成df
            # 改变文本存储的数字
            t = treeview.get_children()
            a = list()
            for i in t:
                a.append(list(treeview.item(i,'values')))
            df_now = pd.DataFrame(a,columns=columns)
            # 按列名输出
            df_now = df_now[["规格型号","补货量","W1","W2","W3","W4"]]
            # 指定列修改千分位为数字
            for i in range(0,len(df_now.columns)):
                try:
                    df_now[df_now.columns[i]] = df_now.loc[:,df_now.columns[i]].apply(
                        lambda x: float(x.replace(",", "")))
                except:
                    pass
            df_now.to_excel(filename+'.xls',index=False)
        
        btn_output = Button(frame,text="下载补货计划",font=('黑体',10,'bold'),width=15,height=1,
                            bg='slategrey',fg='white',borderwidth=5,
                            command=output_plan)
#         CreateToolTip(btn_output,"请确认已保存再下载！")
        btn_output.place(x=520,y=50)
        
        # 搜索功能
        Label(frame,text="筛选字段：",bg='WhiteSmoke',font=("黑体",12)).place(x=5,y=55)
        cbx = ttk.Combobox(frame,font=("黑体",11),width=10) #筛选字段
        comvalue = tkinter.StringVar()
        cbx["values"] = ["全局搜索"] + columns
        cbx.current(1)
        cbx.place(x=85,y=55)
        entry_search = Entry(frame,font=("黑体",11),width=12) # 筛选内容
        entry_search.insert(0, "请输入信息")
        entry_search.place(x=195,y=55)
    #     CreateToolTip(entry_search, "请注意大小写输入！")
        # 先清空表格，再插入数据，当字段选择为空、内容为空则显示全部
        def search_material():
            search_all = rep_result.copy()
            for i in search_all.columns:
                search_all[i] = search_all[i].apply(str)# 必须转字符，否则无法全局搜索

            # 清空
            for item in treeview.get_children():
                treeview.delete(item)
            # 查找并插入数据
            if entry_search.get() != "":
                search_content = str(entry_search.get())
                # 全局搜索
                if cbx.get() == "全局搜索":
                    search_df = pd.DataFrame(columns=search_all.columns)
                    for i in range(len(search_all.columns)):
                        search_df = search_df.append(search_all[search_all[
                            search_all.columns[i]].str.contains(search_content)])                
                    search_df.drop_duplicates(subset=["规格型号"], keep='first',inplace=True)
    #                 print(search_df)
                # 指定字段搜索
                else:
                    appoint = str(cbx.get())
                    search_df = search_all[search_all[appoint].str.contains(search_content)]
    #                 print(search_df)
                # 插入表格
                for i in range(len(search_df)):
                    if i % 2 == 0:
                        tag = "evenrow"
                    else:
                        tag = "oddrow"
                    treeview.insert('', i, values=list(search_df.iloc[i,:]),tags=tag)
            # 若输入值为空则显示全部内容
            else:
                # 插入
                for i in range(len(search_all)):
                    if i % 2 == 0:
                        tag = "evenrow"
                    else:
                        tag = "oddrow"
                    treeview.insert('', i, values=list(search_all.iloc[i,:]),tags=tag)

        btn_search_material = Button(frame,text="查找",font=("黑体",10,'bold'),bg='slategrey',
                                     fg='white',width=9,height=1,borderwidth=5,
                                     command=search_material)
        btn_search_material.place(x=320,y=50)
        
        # 修改补货计划
        def modify_plan(event):
            item_text = treeview.item(treeview.selection(), "values")
            plan = Tk()
            plan.title('补货计划修改')
            plan.geometry('1050x230')

            # Tips
            lb_tips = Label(plan,text="Tips:双击可以并且仅能修改每周的补货量及其备注",
                            font=("华文中宋",12),fg="brown")
            lb_tips.place(x=40,y=200)
            # 获取主数据列名
            columns = list(rep_result.columns)
            modify_treeview = ttk.Treeview(plan,height=1,show="headings",columns=columns)
            modify_treeview.place(x=20,y=20,relwidth=0.98)

            # 横向滚动条
            sb_x = ttk.Scrollbar(plan,command=modify_treeview.xview,orient="horizontal")
            sb_x.config(command=modify_treeview.xview)
            sb_x.place(in_=modify_treeview,relx=0, rely=1,relwidth=1)
            modify_treeview.config(xscrollcommand=sb_x.set)

            for i in range(len(columns)):
                modify_treeview.column(columns[i], width=100, anchor='center')

            # 显示列名
            for i in range(len(columns)):
                modify_treeview.heading(columns[i], text=columns[i])

            # 插入数据
            modify_treeview.insert('', 1, values=item_text)

            # 合并输入数据
            def set_value(event): 
                # 获取鼠标所选item
                for item in modify_treeview.selection():
                    item_text = modify_treeview.item(item, "values")

                column = modify_treeview.identify_column(event.x)# 所在列
                row = modify_treeview.identify_row(event.y)# 所在行，返回
                cn = int(str(column).replace('#',''))
                if cn  not in [5,6,7,8,9]:
                    cn = 100
                entryedit = Entry(plan,width=10)
                entryedit.insert(0,str(item_text[cn-1]))
                entryedit.place(x=150, y=150)
                Label_select = Label(plan,text=str(item_text[cn-1]),width=20)
                Label_select.place(x=150, y=100)
                # 将编辑好的信息更新到数据库中
                def save_edit():
#                     print(item)
                    if cn in [5,6,7,8]:
                        entry_value = int(entryedit.get().replace(",",""))
                        modify_treeview.set(item, column=column,value="{:,}".format(entry_value)) 
                        item_4 = int(item_text[4].replace(",",""))
                        item_5 = int(item_text[5].replace(",",""))
                        item_6 = int(item_text[6].replace(",",""))
                        item_7 = int(item_text[7].replace(",",""))
                        item_n = int(item_text[cn-1].replace(",",""))
                        modify_treeview.set(item, column=3,
                                            value="{:,}".format(item_4+item_5+item_6+item_7+
                                                      entry_value-item_n))
                    else:
                        modify_treeview.set(item, column=column,value=entryedit.get())
                    entryedit.destroy()
                    btn_input.destroy()
                    btn_cancal.destroy()
                    Label_select.destroy()

                btn_input = Button(plan, text='OK', width=7, command=save_edit)
                btn_input.place(x=260,y=150)

                # 取消输入
                def cancal_edit():
                    entryedit.destroy()
                    btn_input.destroy()
                    btn_cancal.destroy()
                    Label_select.destroy()

                btn_cancal = Button(plan, text='Cancel', width=7, command=cancal_edit)
                btn_cancal.place(x=350,y=150)


            # 触发双击事件
            modify_treeview.bind('<Double-1>', set_value)

            Label(plan,text="修改前：").place(x=100,y=100)
            Label(plan,text="修改后：").place(x=100,y=150)

            # 将编辑好的信息更新到数据库中
            def db_update():
                # 获取所有最新数据转为df
                t = modify_treeview.get_children()
                a = list()
                for i in t:
                    a.append(list(modify_treeview.item(i,'values')))
                df_now = pd.DataFrame(a,columns=columns)
                
                # 更新数据库
                next_month = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                          relativedelta(months=1)).strftime("%Y%m") # 下个月时间 
                for i in range(4):
                    SQL_update = "UPDATE AdjustRepPlan SET RepWeek_QTY = "+                     str(df_now.iloc[0,i+4].replace(",",""))+" WHERE JNJ_Date='"+next_month+                     "' AND Material='"+df_now["规格型号"].iloc[0]+"' AND week_No='"+                     df_now.columns[i+4]+"';"
                    PrismDatabaseOperation.Prism_update(SQL_update)
                # 更新备注
                SQL_update_remark = "UPDATE AdjustRepPlan SET Rep_Remark = '"+                 str(df_now.iloc[0,8])+"' WHERE JNJ_Date='"+next_month+"' AND Material='"+                 df_now["规格型号"].iloc[0]+"' AND week_No='"+df_now.columns[4]+"';"
                PrismDatabaseOperation.Prism_update(SQL_update_remark)
                
                # 刷新界面,直接将修改好的信息复制到主界面
                for i in range(6):
                    treeview.set(treeview.selection(), column=i+3,value=df_now.iloc[0,i+3])
                
                # 重新计算value
                # 读取补货计划
                SQL_AdjustRepPlan = "SELECT * From AdjustRepPlan WHERE JNJ_Date= "+next_month+";"
                AdjustRepPlan = PrismDatabaseOperation.Prism_select(SQL_AdjustRepPlan)
                if AdjustRepPlan.empty:
                    lack.append("补货计划")
                # Rep_Remark存储在W1中
                AdjustRepPlan_m = AdjustRepPlan[["Material","JNJ_Date","week_No","Rep_Remark"]]
                AdjustRepPlan = pd.pivot_table(AdjustRepPlan,index=["Material","JNJ_Date"],
                                               columns=["week_No"],
                                               values=["RepWeek_QTY"])
                AdjustRepPlan_col = []
                for i in range(len(AdjustRepPlan.columns)):
                    if type(AdjustRepPlan.columns.values[i][1]) == str :
                        AdjustRepPlan_col.append(AdjustRepPlan.columns.values[i][1])
                    else:
                        AdjustRepPlan_col.append(str(AdjustRepPlan.columns.values[i][1]))
                AdjustRepPlan.columns = AdjustRepPlan_col
                AdjustRepPlan = AdjustRepPlan.reset_index()# 重排索引

                # 合并备注数据
                AdjustRepPlan = pd.merge(AdjustRepPlan,AdjustRepPlan_m,
                                         on=["Material","JNJ_Date"],
                                         how="left")
                # 切片，并只保留含W1的数据
                AdjustRepPlan = AdjustRepPlan[AdjustRepPlan["week_No"]=="W1"]
                AdjustRepPlan = AdjustRepPlan[["Material","W1","W2","W3","W4","Rep_Remark"]]
                
                merge_A_P = pd.merge(ProductMaster,AdjustRepPlan,how="left",on="Material")
                merge_A_P['adjust_Rep_QTY'] = merge_A_P['W1']+merge_A_P['W2']+                 merge_A_P['W3']+merge_A_P['W4']
                merge_A_P["Rep_value"] = merge_A_P["GTS"] * merge_A_P["adjust_Rep_QTY"]
                
                
                # 判断颜色，大于指标红色，小于指标绿色
                if int(Target_amount) > int(sum(merge_A_P['Rep_value'])):
                    lb_amount = Label(frame,text="{:,}".format(
                        new_round(sum(merge_A_P['Rep_value']),2)),
                                      font=('黑体',15),width=15,height=2,fg='green',bg='WhiteSmoke')
                    lb_amount.place(x=820,y=46)
                else:
                    lb_amount = Label(frame,text="{:,}".format(
                        new_round(sum(merge_A_P['Rep_value']),2)),
                                      font=('黑体',15),width=15,height=2,fg='red',bg='WhiteSmoke')
                    lb_amount.place(x=820,y=46)
                tkinter.messagebox.showinfo("提示","成功！")
#                 read_adjust_db() # 最慢最全的更新办法                
                plan.destroy()

            Button(plan,text="确认修改",font=("黑体",12,'bold'),bg='slategrey',fg='white',width=9,
                   height=1,borderwidth=5,command=db_update).place(x=900,y=150)

            plan.mainloop()
        
        treeview.bind('<Double-1>', modify_plan)
        
        # 输入预测月指标
        def order_target(event):
            entry_input = Entry(frame,font=('黑体',15),width=15)
            entry_input.place(x=820,y=10)
            # 插入数据库中的目标金额
            SQL_OrderTarget = "SELECT * FROM OrderTarget WHERE JNJ_Date = '"+next_month+"';"
            OrderTarget_amount = PrismDatabaseOperation.Prism_select(SQL_OrderTarget)
            entry_input.insert(0,OrderTarget_amount["order_target"].iloc[0])
            
            # 确认输入函数，并将输入数据更新至数据库
            def input_target_amount():
                # 保存数据库并以最后的标准为主，方法：先插入下个月信息，再更新输入信息
                # 如果已存在于数据库，则直接更新，否则新增空的过度变量再更新
                if next_month in list(
                    PrismDatabaseOperation.Prism_select("SELECT JNJ_Date FROM OrderTarget")["JNJ_Date"]):
                    # 覆盖输入数据
                    SQL_delete = "DELETE FROM OrderTarget WHERE JNJ_Date = '"+next_month+"';"
                    PrismDatabaseOperation.Prism_delete(SQL_delete)
                    OrderTarget = pd.DataFrame(data={'JNJ_Date':[next_month],
                                                     'order_target':[float(entry_input.get())]})
                    PrismDatabaseOperation.Prism_insert('OrderTarget',OrderTarget)
                else:
                    OrderTarget = pd.DataFrame(data={'JNJ_Date':[next_month],
                                                     'order_target':[float(entry_input.get())]})
                    PrismDatabaseOperation.Prism_insert('OrderTarget',OrderTarget)


                SQL_select_target = "SELECT * FROM OrderTarget WHERE JNJ_Date = '"+next_month+"';"
                Target_amount = PrismDatabaseOperation.Prism_select(SQL_select_target)['order_target'].iloc[0]
                lb_target_amount = Label(frame,text=re_round(Target_amount),anchor="e",
                                         font=('黑体',15),width=15,height=2,bg='WhiteSmoke')
                lb_target_amount.place(x=820,y=3)
                # 双击提示
                CreateToolTip(lb_target_amount, "双击此处即可编辑")
                lb_target_amount.bind('<Double-1>',order_target)
                # 输入指标后变色
                if int(Target_amount) > int(sum(merge_all['Rep_value'])):
                    lb_amount = Label(frame,text=re_round(sum(merge_all['Rep_value'])),anchor="e",
                                      font=('黑体',15),width=15,height=2,fg='green',bg='WhiteSmoke')
                    lb_amount.place(x=820,y=46)
                else:
                    lb_amount = Label(frame,text=re_round(sum(merge_all['Rep_value'])),anchor="e",
                                      font=('黑体',15),width=15,height=2,fg='red',bg='WhiteSmoke')
                    lb_amount.place(x=820,y=46)

                btn_input_cancel.destroy()
                btn_input_target.destroy()
                entry_input.destroy()

            btn_input_target = Button(frame,text="OK",command=input_target_amount)
            btn_input_target.place(x=820,y=30)
            # 输入取消
            def input_cancel():
                btn_input_cancel.destroy()
                btn_input_target.destroy()
                entry_input.destroy()

            btn_input_cancel = Button(frame,text="No",command=input_cancel)
            btn_input_cancel.place(x=920,y=30)

        lb_target_amount.bind('<Double-1>',order_target)
        CreateToolTip(lb_target_amount, "双击此处即可编辑")
    
    read_adjust_db()
    
    
# 周补货追踪（若有W5数据，repw5数量则为0）
def Access_tracking():
    # 显示框
    frame = Frame(window,height=659,width=1015,bg='WhiteSmoke')
    frame.place(x=267,y=61)
    lb_title_f = Label(window,text="当前数据库最新至:"+JNJ_Month(1)[0],font=('黑体',12))
    lb_title_f.place(x=1000,y=25)
    # 标题
    lb_title = Label(window,text='进出追踪',font=('华文中宋',14),bg='WhiteSmoke',
                             fg='black',width=10,height=2)
    lb_title.place(x=280,y=10)
    
    # 金额标题
    title_value = Label(frame,text="补货金额",font=('黑体',12,"bold"),
                        bg='slategrey',relief=RIDGE,fg='white',anchor='center',width=76,height=2)
    title_value.place(x=300,y=10)
    
    # 明细标题
    title_detail = Label(frame,text="补货明细（BX）",font=('黑体',12,"bold"),
                        bg='slategrey',relief=RIDGE,fg='white',anchor='w',width=110,height=2)
    title_detail.place(x=5,y=319)
    
    # 下个月时间 
    next_month = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                  relativedelta(months=1)).strftime("%Y%m") 
    
    # 将调整后的补货计划赋值给rolling做初值
    SQL_adjrolling = "SELECT JNJ_Date FROM AdjustRollingRepPlan"
    if next_month not in list(PrismDatabaseOperation.Prism_select(SQL_adjrolling)['JNJ_Date']):
        SQL_AdjustRepPlan = "SELECT * FROM AdjustRepPlan WHERE JNJ_Date='"+         next_month+"';"
        AdjustRepPlan = PrismDatabaseOperation.Prism_select(SQL_AdjustRepPlan)
        PrismDatabaseOperation.Prism_insert('AdjustRollingRepPlan',AdjustRepPlan)
    # 到货数据上传
    def updata_WeeklyInbound():
        try:
            filename = tkinter.filedialog.askopenfilename().replace("/","\\")
            weekly_inbound = pd.read_excel(filename)
            weekly_inbound.rename(columns={"规格型号":"Material","数量":"Inboundweek_QTY"},
                                  inplace=True)
            jnj_date = filename[filename.rfind("_")+1:filename.rfind("_")+7]
            week_No = filename[filename.rfind("_")+7:filename.rfind("_")+9]
            weekly_inbound['JNJ_Date'] = jnj_date
            weekly_inbound['week_No'] = week_No
            if filename[-7:-5] not in ['W1','W2','W3','W4','W5'] or "到货" not in filename:
                tkinter.messagebox.showerror("错误","请检查文件命名是否正确！")
            else:
                # 提前聚合
                weekly_inbound = weekly_inbound.groupby(by=['JNJ_Date', 'Material','week_No'],
                                                        as_index=False).sum()
    #             print(weekly_inbound)
                weekly_inbound = weekly_inbound[["JNJ_Date","Material","Inboundweek_QTY",'week_No']]

                # 若已存在则删除后替换
                wk_no = list(PrismDatabaseOperation.Prism_select(
                    "SELECT week_No FROM WeeklyInbound WHERE JNJ_Date = '"+
                     filename[-13:-7]+"';")["week_No"].unique())
                # 读取主数据
                SQL_ProductMaster_all = "SELECT Material From ProductMaster;"
                ProductMaster_all = PrismDatabaseOperation.Prism_select(SQL_ProductMaster_all)
                # 检查是否都在ProductMaster里
                lack_material = weekly_inbound[~weekly_inbound["Material"].isin(
                    ProductMaster_all["Material"])]
                if lack_material.empty:
                    if (filename[-7:-5] in wk_no) and (len(wk_no) != 0):
                        SQL_delete = "DELETE FROM WeeklyInbound WHERE JNJ_Date = '"+                         jnj_date+"' AND week_No = '"+week_No+"';"
                        PrismDatabaseOperation.Prism_delete(SQL_delete)
                        PrismDatabaseOperation.Prism_insert("WeeklyInbound",weekly_inbound)
                        tkinter.messagebox.showinfo("成功","导入成功！")
                    else:
                        PrismDatabaseOperation.Prism_insert("WeeklyInbound",weekly_inbound)
                        tkinter.messagebox.showinfo("成功","导入成功！")
                    acl_track() # 同步更新计算
                else:
                    # 维护主数据
                    to_updata = pd.DataFrame(data={},columns=['规格型号','包装规格','分类Level3',
                                                              '分类Level4','ABC','分类',
                                                              '不含税单价','预测状态','MOQ',
                                                              '安全库存天数'])
                    to_updata["规格型号"] = lack_material['Material']
                    to_updata["预测状态"] = "MTS"
                    MasterData.master_update_batch(to_updata)
        except:
            tkinter.messagebox.showerror("错误","更新数据失败，请重新上传！")
    
    # 数据上传标题及按钮
    lb_update = Label(frame,text='到货数据上传：',font=('黑体',12),bg='WhiteSmoke',fg='DimGray',
                      anchor='w',width=30,height=2)
    lb_update.place(x=10,y=10)
    btn_updata = Button(frame,text='选择文件',command=updata_WeeklyInbound,fg='white',
                        font=("黑体",10,'bold'),bg='slategrey',width=9,height=1,borderwidth=5)
    btn_updata.place(x=180,y=16)
    
    # 订货数据上传
    def update_WeeklyOrder():
        try:
            filename = tkinter.filedialog.askopenfilename().replace("/","\\")
            weekly_order = pd.read_excel(filename)
            weekly_order.rename(columns={"规格型号":"Material","数量":"Orderweek_QTY"},
                                  inplace=True)
            jnj_date = filename[filename.rfind("_")+1:filename.rfind("_")+7]
            week_No = filename[filename.rfind("_")+7:filename.rfind("_")+9]
            weekly_order['JNJ_Date'] = jnj_date
            weekly_order['week_No'] = week_No
            weekly_order.drop(weekly_order[weekly_order["Material"].isnull()].index,
                                  inplace=True)#删除空行
            if week_No not in ['W1','W2','W3','W4','W5'] or "订货" not in filename:
                tkinter.messagebox.showerror("错误","请检查文件命名是否正确！")
            else:
                # 提前聚合
                weekly_order = weekly_order.groupby(by=['JNJ_Date', 'Material','week_No'],
                                                            as_index=False).sum()
                weekly_order = weekly_order[["JNJ_Date","Material","Orderweek_QTY",
                                                     'week_No']]
                # 若已存在则删除后替换
                sql_wk = "SELECT week_No FROM WeeklyOrder WHERE JNJ_Date = '"+jnj_date+"';"
                wk_no = list(PrismDatabaseOperation.Prism_select(sql_wk)["week_No"].unique())
                # 读取主数据
                SQL_ProductMaster_all = "SELECT Material From ProductMaster;"
                ProductMaster_all = PrismDatabaseOperation.Prism_select(SQL_ProductMaster_all)
                # 检查是否都在ProductMaster里
                lack_material = weekly_order[~weekly_order["Material"].isin(
                    ProductMaster_all["Material"])]
                if lack_material.empty:
                    if (week_No in wk_no) and (len(wk_no) != 0):
                        SQL_delete = "DELETE FROM WeeklyOrder WHERE JNJ_Date = '"+                         jnj_date+"' AND week_No = '"+week_No+"';"
                        PrismDatabaseOperation.Prism_delete(SQL_delete)
                        PrismDatabaseOperation.Prism_insert("WeeklyOrder",weekly_order)
                        tkinter.messagebox.showinfo("成功","导入成功！")
                    else:
                        PrismDatabaseOperation.Prism_insert("WeeklyOrder",weekly_order)
                        tkinter.messagebox.showinfo("成功","导入成功！")
                    acl_track() # 同步更新计算
                else:
                    # 维护主数据
                    to_updata = pd.DataFrame(data={},columns=['规格型号','包装规格','分类Level3',
                                                              '分类Level4','ABC','分类',
                                                              '不含税单价','预测状态','MOQ',
                                                              '安全库存天数'])
                    to_updata["规格型号"] = lack_material['Material']
                    to_updata["预测状态"] = "MTS"
                    MasterData.master_update_batch(to_updata)
        except:
            tkinter.messagebox.showerror("错误","更新数据失败，请重新上传！")
    
    lb_WeeklyOrder = Label(frame,text='订货数据上传：',font=('黑体',12),bg='WhiteSmoke',
                               fg='DimGray',anchor='w',width=30,height=2)
    lb_WeeklyOrder.place(x=10,y=50)
    btn_WeeklyOrder = Button(frame,text='选择文件',command=update_WeeklyOrder,fg='white',
                             font=("黑体",10,'bold'),bg='slategrey',width=9,height=1,borderwidth=5)
    btn_WeeklyOrder.place(x=180,y=56)
    
    # 读取数据，并将所需数据计算好
    def acl_track():
        # 缺失数据
        lack = []
        # 导入主数据
        SQL_ProductMaster = "SELECT Material,GTS,FCST_state From ProductMaster "+         "WHERE FCST_state = 'MTS';"
        ProductMaster = PrismDatabaseOperation.Prism_select(SQL_ProductMaster)

        # 获取到货信息
        next_month = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                      relativedelta(months=1)).strftime("%Y%m") # 下个月时间 
        SQL_WeeklyInbound = "SELECT * From WeeklyInbound WHERE JNJ_Date = '"+next_month+"';"
        WeeklyInbound = PrismDatabaseOperation.Prism_select(SQL_WeeklyInbound)
        # 提前对其进行聚合
        WeeklyInbound = WeeklyInbound.groupby(by=["JNJ_Date","Material","week_No"],
                                              as_index=False).sum()
        if WeeklyInbound.empty:
            lack.append("实际到货")

        # 获取订货信息
        SQL_WeeklyOrder = "SELECT * From WeeklyOrder WHERE JNJ_Date = '"+next_month+"';"
        WeeklyOrder = PrismDatabaseOperation.Prism_select(SQL_WeeklyOrder)
        if WeeklyOrder.empty:
            lack.append("实际订货")

        # 获取指标
        SQL_OrderTarget = "SELECT * FROM OrderTarget WHERE JNJ_Date = '"+next_month+"';"
        OrderTarget = PrismDatabaseOperation.Prism_select(SQL_OrderTarget)
        if OrderTarget.empty:
            lack.append("指标")

        # 读取补货计划
        SQL_Rep_plan = "SELECT JNJ_Date,Material,RepWeek_QTY,week_No FROM"+         " AdjustRepPlan WHERE JNJ_Date = '"+next_month+"';"
        Rep_plan = PrismDatabaseOperation.Prism_select(SQL_Rep_plan)
        if Rep_plan.empty:
            lack.append("补货计划")

        # 上个月在途
        SQL_Intransit = "SELECT * FROM Intransit WHERE JNJ_Date = '"+JNJ_Month(1)[0]+"';"
        Intransit = PrismDatabaseOperation.Prism_select(SQL_Intransit)
        if Intransit.empty:
            lack.append("上月在途")

        if lack != []:
            tkinter.messagebox.showerror("警告",str(lack)+"数据缺失！")

        # Week
        week_No = PrismDatabaseOperation.Prism_select("SELECT week_No FROM WeeklyInbound WHERE JNJ_Date = '"+
                                        next_month+"';")

        # 计算上个月在途金额
        SQL_ProductMaster_all = "SELECT Material,GTS,FCST_state From ProductMaster;"
        ProductMaster_all = PrismDatabaseOperation.Prism_select(SQL_ProductMaster_all)
        merge_Intransit = pd.merge(Intransit,ProductMaster_all,on='Material',how='left')
        merge_Intransit.fillna(0,inplace=True)
        Intransit_value = int(sum(merge_Intransit['GTS']*merge_Intransit['Intransit_QTY']))
        lb_Intransit = Label(frame,text="上月在途金额:",anchor='w',font=('黑体',12),
                             bg='WhiteSmoke',fg='DimGray',width=15,height=2)
        lb_Intransit.place(x=10,y=90)
        lb_Intransit_value = Label(frame,text="{:,}".format(Intransit_value),anchor='e',
                                   font=('黑体',12),bg='WhiteSmoke',fg='DimGray',
                                   width=15,height=2)
        lb_Intransit_value.place(x=140,y=90)

        # 合并数据--上个月计算的预测补货计划与本月实际到货
        merge_P_R = pd.merge(WeeklyInbound,Rep_plan,on=["Material","JNJ_Date","week_No"],
                             how='outer')
        merge_all = pd.merge(merge_P_R,ProductMaster,on=["Material"],how='left')
        merge_all.fillna(0,inplace=True)
        merge_all = merge_all[merge_all["FCST_state"]!=0] # 只显示MTS

        # 计算每周实际到货金额与计算金额（显示和计算依据）
        merge_all["InboundWeekly_value"] = merge_all["GTS"]*merge_all["Inboundweek_QTY"]
        merge_all["Rep_plan_value"] = merge_all["GTS"]*merge_all["RepWeek_QTY"]
        # 删除week_No为0的情况
        merge_all.drop(index=merge_all[merge_all["week_No"]==0].index,inplace=True) 
#         merge_all.to_excel(r'merge_all.xlsx',index=False)
#         # 去重复
#         merge_all.drop_duplicates(subset=['Material','week_No'], keep='first',inplace=True)
        merge_value = merge_all.copy()
        merge_value = pd.merge(merge_value,WeeklyOrder,on=["Material","JNJ_Date","week_No"],
                               how='outer')
        # merge_value.fillna(0,inplace=True)#存在空
        merge_value["Orderweek_value"] = merge_value["GTS"]*merge_value["Orderweek_QTY"]
        merge_value.fillna(0,inplace=True)
        merge_value = merge_value.groupby(by=['week_No'],as_index=False).sum()
        merge_value = merge_value[["week_No","InboundWeekly_value","Orderweek_value",
                                   "Rep_plan_value"]]
        merge_value["contrast_value"] = merge_value[
            "InboundWeekly_value"]/merge_value["Rep_plan_value"]
        merge_value = merge_value.T # 转置
        merge_value.columns = merge_value.iloc[0,:] # 将计算的第一行作为列名
        merge_value.drop(index='week_No',inplace=True) # 删除第一行
        #         print(merge_value)
        merge_value.reset_index(inplace=True)
        #         print(merge_value.columns)
        # 若无W5数据，则手动添加为0
        if len(merge_value.columns) <= 5:
            merge_value['W5'] = 0
            merge_value.columns = ['week_No','W1','W2','W3','W4','W5']
        else:
            merge_value.columns = ['week_No','W1','W2','W3','W4','W5']
        merge_value.fillna(0,inplace=True)
        #         print(merge_value)

        # 计算本周到货金额、本周订货金额
        if len(week_No['week_No'].unique()) == 1:
            inbound_value = int(merge_value.iloc[0,:]["W1"])
            order_value = int(merge_value.iloc[1,:]["W1"])
        elif len(week_No['week_No'].unique()) == 2:
            inbound_value = int(merge_value.iloc[0,:]["W2"])
            order_value = int(merge_value.iloc[1,:]["W2"])
        elif len(week_No['week_No'].unique()) == 3:
            inbound_value = int(merge_value.iloc[0,:]["W3"])
            order_value = int(merge_value.iloc[1,:]["W3"])
        elif len(week_No['week_No'].unique()) == 4:
            inbound_value = int(merge_value.iloc[0,:]["W4"])
            order_value = int(merge_value.iloc[1,:]["W4"])
        elif len(week_No['week_No'].unique()) == 5:
            inbound_value = int(merge_value.iloc[0,:]["W5"])
            order_value = int(merge_value.iloc[1,:]["W5"])
        else:
            inbound_value = 0
            order_value = 0

        lb_inbound = Label(frame,text="本周到货金额:",anchor='w',
                           font=('黑体',12),bg='WhiteSmoke',fg='DimGray',width=15,height=2)
        lb_inbound.place(x=10,y=130)
        lb_inbound_value = Label(frame,text="{:,}".format(inbound_value),anchor='e',
                                 font=('黑体',12),bg='WhiteSmoke',fg='DimGray',width=15,height=2)
        lb_inbound_value.place(x=140,y=130)

        lb_order = Label(frame,text="本周订货金额:",anchor='w',
                         font=('黑体',12),bg='WhiteSmoke',fg='DimGray',width=15,height=2)
        lb_order.place(x=10,y=170)
        lb_order_value = Label(frame,text="{:,}".format(order_value),anchor='e',
                              font=('黑体',12),bg='WhiteSmoke',fg='DimGray',width=15,height=2)
        lb_order_value.place(x=140,y=170)
        
        # 显示MTD到货金额=总到货金额/目标金额
        MTD_value = sum(merge_value.iloc[0][1:])/int(OrderTarget["order_target"].iloc[0])
        lb_MTD = Label(frame,text="MTD到货金额占比:",anchor='w',font=('黑体',12),
                       bg='WhiteSmoke',fg='DimGray',width=16,height=2)
        lb_MTD.place(x=10,y=210)
        lb_MTD_value = Label(frame,text="{:.1%}".format(MTD_value),anchor='e',
                             font=('黑体',12),
                             bg='WhiteSmoke',fg='brown',width=15,height=2)
        lb_MTD_value.place(x=140,y=210)
        
        # 计算本月剩余金额=目标金额-本月已到货金额
        InboundWeekly_value = int(OrderTarget["order_target"].iloc[0]-
                                  sum(merge_all["GTS"]*merge_all["Inboundweek_QTY"]))
        lb_InboundWeekly = Label(frame,text=next_month+"剩余额度:",anchor='w',
                                 font=('黑体',12),
                                 bg='WhiteSmoke',fg='DimGray',width=15,height=2)
        lb_InboundWeekly.place(x=10,y=250)
        lb_InboundWeekly_value = Label(frame,text="{:,}".format(InboundWeekly_value),anchor='e',
                                       font=('黑体',12),bg='WhiteSmoke',fg='brown',
                                       width=15,height=2)
        lb_InboundWeekly_value.place(x=140,y=250)
        
        # 将得到的计算结果展示在界面
        # ------------计算金额显示--------------#
        merge_value['week_No'] = ["实际到货金额","实际订货金额","计划补货金额","到货/计划"]
        columns = list(merge_value.columns)
        # 设置标题、字体、行高等样式
        style_head = ttk.Style()
        style_head.configure("MyStyle.Treeview.Heading",rowheight=56, font=("黑体",11))
        style_value = ttk.Style()
        style_value.configure("MyStyle.Treeview", rowheight=50, font=("黑体",11))
        treeview_value = ttk.Treeview(frame, height=4, show="headings",columns=columns,
                                      style='MyStyle.Treeview')
#         treeview_value = ttk.Treeview(frame, height=4, show="headings",columns=columns)
        treeview_value.place(x=300,y=60)

        # 表示列,不显示
        for i in range(0,len(merge_value.columns)):
            treeview_value.column(str(merge_value.columns[i]), width=115, anchor='center') 

        # 显示表头
        for i in range(len(merge_value.columns)):
            treeview_value.heading(str(merge_value.columns[i]), text=str(merge_value.columns[i]))

        # 插入数据，数字显示为千分位,最后一行显示百分比
        for i in range(len(merge_value)-1):
            treeview_value.insert('', i, 
                                  values=(merge_value[merge_value.columns[0]].iloc[i],
                                "{:,}".format(int(merge_value[merge_value.columns[1]].iloc[i])),
                                "{:,}".format(int(merge_value[merge_value.columns[2]].iloc[i])),
                                "{:,}".format(int(merge_value[merge_value.columns[3]].iloc[i])),
                                "{:,}".format(int(merge_value[merge_value.columns[4]].iloc[i])),
                                "{:,}".format(int(merge_value[merge_value.columns[5]].iloc[i]))
                                         )
                                 )
#         print('{:.0%}'.format(merge_value[merge_value.columns[1]].iloc[2]))
        treeview_value.insert('',len(merge_value)-1, 
                              values=(merge_value[merge_value.columns[0]].iloc[len(merge_value)-1],
                      '{:.0%}'.format(merge_value[merge_value.columns[1]].iloc[len(merge_value)-1]),
                      '{:.0%}'.format(merge_value[merge_value.columns[2]].iloc[len(merge_value)-1]),
                      '{:.0%}'.format(merge_value[merge_value.columns[3]].iloc[len(merge_value)-1]),
                      '{:.0%}'.format(merge_value[merge_value.columns[4]].iloc[len(merge_value)-1]),
                      '{:.0%}'.format(merge_value[merge_value.columns[5]].iloc[len(merge_value)-1])
                                        ))
        # ------------计算金额显示--------------#

        # 计算补货计划与实际到货内容
        merge_view = merge_all.pivot_table(values=['RepWeek_QTY','Inboundweek_QTY'],
                                           index=['Material','FCST_state'],columns='week_No')
        merge_view.reset_index(inplace=True) # 重设index，方便显示和计算数据
        merge_view = pd.merge(merge_view,Intransit[["Material","Intransit_QTY"]],on="Material",
                              how="outer")
        merge_view.fillna(0,inplace=True)
        del merge_view["Material"] # merge后会产生两列不同类型的material列，删除一列即可
        # 复核列名重置为一维列名
        merge_view_col = []
        for i in range(len(merge_view.columns)):
            if type(merge_view.columns[i]) == str:
                merge_view_col.append(str(merge_view.columns[i]))
            else:
                merge_view_col.append(str(merge_view.columns[i][0])+str(merge_view.columns[i][1]))
        merge_view.columns = merge_view_col
        # del merge_view["GTS"] # GTS 无需展示
        if "Inboundweek_QTYW5" not in merge_view.columns:
            merge_view["Inboundweek_QTYW5"] = 0
        elif "Inboundweek_QTYW4" not in merge_view.columns:
            merge_view["Inboundweek_QTYW4"] = 0
        elif "Inboundweek_QTYW3" not in merge_view.columns:
            merge_view["Inboundweek_QTYW3"] = 0
        elif "Inboundweek_QTYW2" not in merge_view.columns:
            merge_view["Inboundweek_QTYW2"] = 0
        elif "Inboundweek_QTYW1" not in merge_view.columns:
            merge_view["Inboundweek_QTYW1"] = 0
            
        merge_view["Inboundweek_QTY"] = merge_view["Inboundweek_QTYW1"]+merge_view["Inboundweek_QTYW2"]+        merge_view["Inboundweek_QTYW3"]+merge_view["Inboundweek_QTYW4"]+merge_view["Inboundweek_QTYW5"]
        merge_view["RepWeek_QTY"] = merge_view["RepWeek_QTYW1"]+merge_view["RepWeek_QTYW2"]+        merge_view["RepWeek_QTYW3"]+merge_view["RepWeek_QTYW4"]
        merge_view.drop(merge_view[merge_view["Material"]==0].index,inplace=True)
        merge_view = merge_view[["Material","FCST_state","Intransit_QTY","Inboundweek_QTY",
                                 "Inboundweek_QTYW1","Inboundweek_QTYW2","Inboundweek_QTYW3",
                                 "Inboundweek_QTYW4","Inboundweek_QTYW5","RepWeek_QTY",
                                 "RepWeek_QTYW1","RepWeek_QTYW2","RepWeek_QTYW3","RepWeek_QTYW4"]]
        # 重命名,"RepWeek_QTYW5":"W5"不显示
        merge_view.rename(columns={"Material":"规格型号","Intransit_QTY":"上月在途",
                                   "Inboundweek_QTY":"月到货量","Inboundweek_QTYW1":"W1 ",
                                   "Inboundweek_QTYW2":"W2 ","Inboundweek_QTYW3":"W3 ",
                                   "Inboundweek_QTYW4":"W4 ","Inboundweek_QTYW5":"W5 ",
                                   "RepWeek_QTY":"月补货量","RepWeek_QTYW1":"W1",
                                   "RepWeek_QTYW2":"W2","RepWeek_QTYW3":"W3",
                                   "RepWeek_QTYW4":"W4","FCST_state":"预测状态"},inplace=True)

        
#         print(merge_view)
        # ------------补货计划与实际到货显示--------------#
        columns_view = list(merge_view.columns)
        # 设置标题、字体、行高等样式
        treeview_view = ttk.Treeview(frame, height=12,show="headings", columns=columns_view)
        treeview_view.place(x=5,y=370,relwidth=0.97)
        # 竖向滚动条
        sb_y = ttk.Scrollbar(frame,command=treeview_view.yview)
        sb_y.config(command=treeview_view.yview)
        sb_y.place(in_=treeview_view,relx=1, rely=0,relheight=1)
        treeview_view.config(yscrollcommand=sb_y.set)
        # 横向滚动条
        sb_x = ttk.Scrollbar(frame,command=treeview_view.xview,orient="horizontal")
        sb_x.config(command=treeview_view.xview)
        sb_x.place(in_=treeview_view,relx=0, rely=1,relwidth=1)
        treeview_view.config(xscrollcommand=sb_x.set)

        # 行交替颜色
        def fixed_map(option):# 重要！无此步骤则无法显示
            return [elm for elm in style.map("Treeview", query_opt=option)
                    if elm[:2] != ("!disabled", "!selected")]
        style = ttk.Style()
        style.map("Treeview",foreground=fixed_map("foreground"),background=fixed_map("background"))

        treeview_view.tag_configure('oddrow', background='LightGrey')
        treeview_view.tag_configure('evenrow', background='white')

        # 表示列,不显示
        treeview_view.column(str(merge_view.columns[0]), width=90, anchor='center')
        treeview_view.column(str(merge_view.columns[1]), width=60, anchor='center')
        for i in range(2,len(merge_view.columns)):
            treeview_view.column(str(merge_view.columns[i]), width=68, anchor='center') 

        # 显示表头
        for i in range(len(merge_view.columns)):
            treeview_view.heading(str(merge_view.columns[i]), text=str(merge_view.columns[i]))
        
        # 坐标重排
        def odd_even_color():
            for index,row in enumerate(treeview_view.get_children()):
                if index % 2 == 0:
                    treeview_view.item(row,tags="evenrow")
                else:
                    treeview_view.item(row,tags="oddrow")
                    
        # 插入数据，数字显示为千分位,最后一行显示百分比
        for i in range(len(merge_view)):
            if i % 2 == 0:
                tag = "evenrow"
            else:
                tag = "oddrow"
#             print("{:,}".format(int(merge_view["W1"].iloc[i])))
            treeview_view.insert('', i,
                                 values=(merge_view[merge_view.columns[0]].iloc[i],
                                         merge_view[merge_view.columns[1]].iloc[i],
                       "{:,}".format(int(merge_view[merge_view.columns[2]].iloc[i])),
                       "{:,}".format(int(merge_view[merge_view.columns[3]].iloc[i])),
                       "{:,}".format(int(merge_view[merge_view.columns[4]].iloc[i])),
                       "{:,}".format(int(merge_view[merge_view.columns[5]].iloc[i])),
                       "{:,}".format(int(merge_view[merge_view.columns[6]].iloc[i])),
                       "{:,}".format(int(merge_view[merge_view.columns[7]].iloc[i])),
                       "{:,}".format(int(merge_view[merge_view.columns[8]].iloc[i])),
                       "{:,}".format(int(merge_view[merge_view.columns[9]].iloc[i])),
                       "{:,}".format(int(merge_view[merge_view.columns[10]].iloc[i])),
                       "{:,}".format(int(merge_view[merge_view.columns[11]].iloc[i])),
                       "{:,}".format(int(merge_view[merge_view.columns[12]].iloc[i])),
                       "{:,}".format(int(merge_view[merge_view.columns[13]].iloc[i]))
                                        ),
                                 tags=tag) 
        
        # Treeview、列名、排列方式
        def treeview_sort_column(tv, col, reverse):  
            L = [(tv.set(k, col), k) for k in tv.get_children('')]
            try:
                for i in range(len(L)):
                    L[i] = (float(L[i][0].replace(',', '')),L[i][1])
            except:
                pass
            L.sort(reverse=reverse)  # 排序方式
            # 根据排序后索引移动
            for index, (val, k) in enumerate(L):
                tv.move(k, '', index)
            # 重写标题，使之成为再点倒序的标题
            tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))
            odd_even_color()
            
        # 绑定函数，使表头可排序
        for col in columns_view:
            treeview_view.heading(col, text=col, command=
                             lambda _col=col: treeview_sort_column(treeview_view, _col, False))
        
        # 添加筛选功能
        Label(frame,text="筛选字段：",bg='slategrey',font=("黑体",12,"bold"),fg="white"
             ).place(x=590,y=330)
        cbx = ttk.Combobox(frame,font=("黑体",12),width=10) #筛选字段
        comvalue = tkinter.StringVar()
        cbx["values"] = ["全局搜索"] + columns_view
        cbx.current(1)
        cbx.place(x=680,y=330)
        entry_search = Entry(frame,font=("黑体",12),width=12) # 筛选内容
        entry_search.insert(0, "请输入信息")
        entry_search.place(x=790,y=330)
    #     CreateToolTip(entry_search, "请注意大小写输入！")
        # 先清空表格，再插入数据，当字段选择为空、内容为空则显示全部
        def search_material():
            search_all = merge_view.copy()
            for i in search_all.columns:
                search_all[i] = search_all[i].apply(str)# 必须转字符，否则无法全局搜索

            # 清空
            for item in treeview_view.get_children():
                treeview_view.delete(item)
            # 查找并插入数据
            if entry_search.get() != "":
                search_content = str(entry_search.get())
                # 全局搜索
                if cbx.get() == "全局搜索":
                    search_df = pd.DataFrame(columns=search_all.columns)
                    for i in range(len(search_all.columns)):
                        search_df = search_df.append(search_all[search_all[
                            search_all.columns[i]].str.contains(search_content)])                
                    search_df.drop_duplicates(subset=["规格型号"], keep='first',inplace=True)
    #                 print(search_df)
                # 指定字段搜索
                else:
                    appoint = str(cbx.get())
                    search_df = search_all[search_all[appoint].str.contains(search_content)]
    #                 print(search_df)
                # 插入表格
                for i in range(len(search_df)):
                    if i % 2 == 0:
                        tag = "evenrow"
                    else:
                        tag = "oddrow"
                    treeview_view.insert('', i, values=list(search_df.iloc[i,:]),tags=tag)
            # 若输入值为空则显示全部内容
            else:
                # 插入
                for i in range(len(search_all)):
                    if i % 2 == 0:
                        tag = "evenrow"
                    else:
                        tag = "oddrow"
                    treeview_view.insert('', i, values=list(search_all.iloc[i,:]),tags=tag)

        btn_search_material = Button(frame,text="查找",font=("黑体",10,'bold'),bg='slategrey',
                                     fg='white',width=9,height=1,borderwidth=5,
                                     command=search_material)
        btn_search_material.place(x=900,y=325)
    
    acl_track() # 第一次点击运行显示


# 周补货更新
def rolling_rep():
    # 主显示框
    frame = Frame(window,height=655,width=1015,bg='WhiteSmoke')
    frame.place(x=267,y=61)
    # 下个月时间 
    next_month = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                  relativedelta(months=1)).strftime("%Y%m")
    # wk_No，以WeeklyOutbound的周数为准
    SQL_week = "SELECT week_No FROM WeeklyOutbound WHERE JNJ_Date = '"+next_month+"';"
    try:
        wk_No = sorted(list(PrismDatabaseOperation.Prism_select(SQL_week)["week_No"].unique()))[-1]
    except:
        wk_No = "无"
        
    # 提示
    lb_title_f = Label(window,text="当前数据库最新至:"+JNJ_Month(1)[0],font=('黑体',12))
    lb_title_f.place(x=1000,y=25)
    
    # 标题
    lb_title = Label(window,text='补货更新',font=('华文中宋',14),bg='WhiteSmoke',
                             fg='black',width=10,height=2)
    lb_title.place(x=280,y=10)
    
    # 初始值
    SQL_adjrolling = "SELECT JNJ_Date FROM AdjustRollingRepPlan"
    if next_month not in list(PrismDatabaseOperation.Prism_select(SQL_adjrolling)['JNJ_Date']):
        SQL_AdjustRepPlan = "SELECT * FROM AdjustRepPlan WHERE JNJ_Date='"+         next_month+"';"
        AdjustRepPlan = PrismDatabaseOperation.Prism_select(SQL_AdjustRepPlan)
        PrismDatabaseOperation.Prism_insert('AdjustRollingRepPlan',AdjustRepPlan)
        
    #出货数据上传
    def update_WeeklyOutbound():
        try:
            filename = tkinter.filedialog.askopenfilename().replace("/","\\")
            weekly_outbound = pd.read_excel(filename)
            weekly_outbound.rename(columns={"规格型号":"Material","数量":"Outboundweek_QTY"},
                                  inplace=True)
            jnj_date = filename[filename.rfind("_")+1:filename.rfind("_")+7]
            week_No = filename[filename.rfind("_")+7:filename.rfind("_")+9]
            weekly_outbound['JNJ_Date'] = jnj_date
            weekly_outbound['week_No'] = week_No
            weekly_outbound.drop(weekly_outbound[weekly_outbound["Material"].isnull()].index,
                                  inplace=True)#删除空行
            if week_No not in ['W1','W2','W3','W4','W5'] or "出库" not in filename:
                tkinter.messagebox.showerror("错误","请检查文件命名是否正确！")
            else:
                # 提前聚合
                weekly_outbound = weekly_outbound.groupby(by=['JNJ_Date', 'Material','week_No'],
                                                            as_index=False).sum()
                weekly_outbound = weekly_outbound[["JNJ_Date","Material","Outboundweek_QTY",
                                                     'week_No']]
                # 若已存在则删除后替换
                SQL_wk =  "SELECT week_No FROM WeeklyOutbound WHERE JNJ_Date = '"+jnj_date+"';"
                wk_no = list(PrismDatabaseOperation.Prism_select(SQL_wk)["week_No"].unique())
                # 读取主数据
                SQL_ProductMaster_all = "SELECT Material From ProductMaster;"
                ProductMaster_all = PrismDatabaseOperation.Prism_select(SQL_ProductMaster_all)
                # 检查是否都在ProductMaster里
                lack_material = weekly_outbound[~weekly_outbound["Material"].isin(
                    ProductMaster_all["Material"])]
                if lack_material.empty:
                    if (week_No in wk_no) and (len(wk_no) != 0):
                        SQL_delete = "DELETE FROM WeeklyOutbound WHERE JNJ_Date = '"+                         jnj_date+"' AND week_No = '"+week_No+"';"
                        PrismDatabaseOperation.Prism_delete(SQL_delete)
                        PrismDatabaseOperation.Prism_insert("WeeklyOutbound",weekly_outbound)
                        tkinter.messagebox.showinfo("成功","导入成功！")
                    else:
                        PrismDatabaseOperation.Prism_insert("WeeklyOutbound",weekly_outbound)
                        tkinter.messagebox.showinfo("成功","导入成功！")
                    # wk_No，以WeeklyOutbound的周数为准
                    SQL_week = "SELECT week_No FROM WeeklyOutbound WHERE JNJ_Date = '"+next_month+"';"
                    wk_No = sorted(list(PrismDatabaseOperation.Prism_select(SQL_week)["week_No"].unique()))[-1]
                    lb_week_Outbound = Label(frame,text="更新至 "+wk_No,font=('黑体',10),bg='WhiteSmoke',
                              fg='DimGray',anchor='w',height=2)
                    lb_week_Outbound.place(x=280,y=15)
                else:
                    # 维护主数据
                    to_updata = pd.DataFrame(data={},columns=['规格型号','包装规格','分类Level3',
                                                              '分类Level4','ABC','分类',
                                                              '不含税单价','预测状态','MOQ',
                                                              '安全库存天数'])
                    to_updata["规格型号"] = lack_material['Material']
                    to_updata["预测状态"] = "MTS"
                    MasterData.master_update_batch(to_updata)
        except:
            tkinter.messagebox.showerror("错误","更新数据失败，请重新上传！")
    

    # 数据上传标题及按钮
    lb_update = Label(frame,text='出库数据上传：',font=('黑体',12),bg='WhiteSmoke',fg='DimGray',
                      anchor='w',width=30,height=2)
    lb_update.place(x=10,y=10)
    lb_week_Outbound = Label(frame,text="更新至 "+wk_No,font=('黑体',10),bg='WhiteSmoke',
                              fg='DimGray',anchor='w',height=2)
    lb_week_Outbound.place(x=280,y=15)
    btn_WeeklyOutbound = Button(frame,text='选择文件',command=update_WeeklyOutbound,fg='white',
                             font=("黑体",10,'bold'),bg='slategrey',width=9,height=1,borderwidth=5)
    btn_WeeklyOutbound.place(x=180,y=16)
    
    # 缺或数据上传
    def update_WeeklyBackorder():
        try:
            filename = tkinter.filedialog.askopenfilename().replace("/","\\")
            weekly_backorder = pd.read_excel(filename)
            weekly_backorder.rename(columns={"规格型号":"Material","数量":"Backorderweek_QTY"},
                                  inplace=True)
            jnj_date = filename[filename.rfind("_")+1:filename.rfind("_")+7]
            week_No = filename[filename.rfind("_")+7:filename.rfind("_")+9]
            weekly_backorder['JNJ_Date'] = jnj_date
            weekly_backorder['week_No'] = week_No
            weekly_backorder.drop(weekly_backorder[weekly_backorder["Material"].isnull()].index,
                                  inplace=True)#删除空行
            if week_No not in ['W1','W2','W3','W4','W5'] or "缺货" not in filename:
                tkinter.messagebox.showerror("错误","请检查文件命名是否正确！")
            else:
                # 提前聚合
                weekly_backorder = weekly_backorder.groupby(by=['JNJ_Date', 'Material','week_No'],
                                                            as_index=False).sum()
                weekly_backorder = weekly_backorder[["JNJ_Date","Material","Backorderweek_QTY",
                                                     'week_No']]
                # 若已存在则删除后替换
                sql_wk = "SELECT week_No FROM WeeklyBackorder WHERE JNJ_Date = '"+jnj_date+"';"
                wk_no = list(PrismDatabaseOperation.Prism_select(sql_wk)["week_No"].unique())
                # 读取主数据
                SQL_ProductMaster_all = "SELECT Material From ProductMaster;"
                ProductMaster_all = PrismDatabaseOperation.Prism_select(SQL_ProductMaster_all)
                # 检查是否都在ProductMaster里
                lack_material = weekly_backorder[~weekly_backorder["Material"].isin(
                    ProductMaster_all["Material"])]
                if lack_material.empty:
                    if (week_No in wk_no) and (len(wk_no) != 0):
                        SQL_delete = "DELETE FROM WeeklyBackorder WHERE JNJ_Date = '"+                         jnj_date+"' AND week_No = '"+week_No+"';"
                        PrismDatabaseOperation.Prism_delete(SQL_delete)
                        PrismDatabaseOperation.Prism_insert("WeeklyBackorder",weekly_backorder)
                        tkinter.messagebox.showinfo("成功","导入成功！")
                    else:
                        PrismDatabaseOperation.Prism_insert("WeeklyBackorder",weekly_backorder)
                        tkinter.messagebox.showinfo("成功","导入成功！")
                    # wk_No，WeeklyBackorder
                    SQL_Backorder = "SELECT week_No FROM WeeklyBackorder WHERE JNJ_Date = '"+                     next_month+"';"
                    wk_No_Backorder = sorted(list(
                        PrismDatabaseOperation.Prism_select(SQL_Backorder)["week_No"].unique()))[-1]
                    lb_week_Backorder = Label(frame,text="更新至 "+wk_No_Backorder,font=('黑体',10),
                                              bg='WhiteSmoke',fg='DimGray',anchor='w',height=2)
                    lb_week_Backorder.place(x=280,y=55)
#                     acl_rolling() # 同步更新计算
                else:
                    # 维护主数据
                    to_updata = pd.DataFrame(data={},columns=['规格型号','包装规格','分类Level3',
                                                              '分类Level4','ABC','分类',
                                                              '不含税单价','预测状态','MOQ',
                                                              '安全库存天数'])
                    to_updata["规格型号"] = lack_material['Material']
                    to_updata["预测状态"] = "MTS"
                    MasterData.master_update_batch(to_updata)
        except:
            tkinter.messagebox.showerror("错误","更新数据失败，请重新上传！")
    

    # 数据上传标题及按钮
    lb_WeeklyOrder = Label(frame,text='缺货数据上传：',font=('黑体',12),bg='WhiteSmoke',
                               fg='DimGray',anchor='w',width=30,height=2)
    lb_WeeklyOrder.place(x=10,y=50)
    # wk_No，WeeklyBackorder
    SQL_Backorder = "SELECT week_No FROM WeeklyBackorder WHERE JNJ_Date = '"+next_month+"';"
    try:
        wk_No_Backorder = sorted(list(PrismDatabaseOperation.Prism_select(SQL_Backorder)["week_No"].unique()))[-1]
    except:
        wk_No_Backorder = "无"
    lb_week_Backorder = Label(frame,text="更新至 "+wk_No_Backorder,font=('黑体',10),bg='WhiteSmoke',
                              fg='DimGray',anchor='w',height=2)
    lb_week_Backorder.place(x=280,y=55)
    btn_WeeklyBackorder = Button(frame,text='选择文件',command=update_WeeklyBackorder,fg='white',
                             font=("黑体",10,'bold'),bg='slategrey',width=9,height=1,borderwidth=5)
    btn_WeeklyBackorder.place(x=180,y=56)
    
    # 到货数据上传
    def updata_WeeklyInbound():
        try:
            filename = tkinter.filedialog.askopenfilename().replace("/","\\")
            weekly_inbound = pd.read_excel(filename)
            weekly_inbound.rename(columns={"规格型号":"Material","数量":"Inboundweek_QTY"},
                                  inplace=True)
            jnj_date = filename[filename.rfind("_")+1:filename.rfind("_")+7]
            week_No = filename[filename.rfind("_")+7:filename.rfind("_")+9]
            weekly_inbound['JNJ_Date'] = jnj_date
            weekly_inbound['week_No'] = week_No
            if filename[-7:-5] not in ['W1','W2','W3','W4','W5'] or "到货" not in filename:
                tkinter.messagebox.showerror("错误","请检查文件命名是否正确！")
            else:
                # 提前聚合
                weekly_inbound = weekly_inbound.groupby(by=['JNJ_Date', 'Material','week_No'],
                                                        as_index=False).sum()
    #             print(weekly_inbound)
                weekly_inbound = weekly_inbound[["JNJ_Date","Material","Inboundweek_QTY",'week_No']]

                # 若已存在则删除后替换
                wk_no = list(PrismDatabaseOperation.Prism_select(
                    "SELECT week_No FROM WeeklyInbound WHERE JNJ_Date = '"+
                     filename[-13:-7]+"';")["week_No"].unique())
                # 读取主数据
                SQL_ProductMaster_all = "SELECT Material From ProductMaster;"
                ProductMaster_all = PrismDatabaseOperation.Prism_select(SQL_ProductMaster_all)
                # 检查是否都在ProductMaster里
                lack_material = weekly_inbound[~weekly_inbound["Material"].isin(
                    ProductMaster_all["Material"])]
                if lack_material.empty:
                    if (filename[-7:-5] in wk_no) and (len(wk_no) != 0):
                        SQL_delete = "DELETE FROM WeeklyInbound WHERE JNJ_Date = '"+                         jnj_date+"' AND week_No = '"+week_No+"';"
                        PrismDatabaseOperation.Prism_delete(SQL_delete)
                        PrismDatabaseOperation.Prism_insert("WeeklyInbound",weekly_inbound)
                        tkinter.messagebox.showinfo("成功","导入成功！")
                    else:
                        PrismDatabaseOperation.Prism_insert("WeeklyInbound",weekly_inbound)
                        tkinter.messagebox.showinfo("成功","导入成功！")
                    # wk_No，WeeklyInbound
                    SQL_Inbound = "SELECT week_No FROM WeeklyInbound WHERE JNJ_Date = '"+                     next_month+"';"
                    wk_No_Inbound = sorted(list(
                        PrismDatabaseOperation.Prism_select(SQL_Inbound)["week_No"].unique()))[-1]
                    lb_week_Inbound = Label(frame,text="更新至 "+wk_No_Inbound,font=('黑体',10),
                                            bg='WhiteSmoke',fg='DimGray',anchor='w',height=2)
                    lb_week_Inbound.place(x=280,y=95)
#                     acl_rolling()
                else:
                    # 维护主数据
                    to_updata = pd.DataFrame(data={},columns=['规格型号','包装规格','分类Level3',
                                                              '分类Level4','ABC','分类',
                                                              '不含税单价','预测状态','MOQ',
                                                              '安全库存天数'])
                    to_updata["规格型号"] = lack_material['Material']
                    to_updata["预测状态"] = "MTS"
                    MasterData.master_update_batch(to_updata)
        except:
            tkinter.messagebox.showerror("错误","更新数据失败，请重新上传！")
        
    # 数据上传标题及按钮
    lb_update = Label(frame,text='到货数据上传：',font=('黑体',12),bg='WhiteSmoke',fg='DimGray',
                      anchor='w',width=30,height=2)
    lb_update.place(x=10,y=90)
    # wk_No，WeeklyInbound
    SQL_Inbound = "SELECT week_No FROM WeeklyInbound WHERE JNJ_Date = '"+next_month+"';"
    try:
        wk_No_Inbound = sorted(list(PrismDatabaseOperation.Prism_select(SQL_Inbound)["week_No"].unique()))[-1]
    except:
        wk_No_Inbound = "无"
    lb_week_Inbound = Label(frame,text="更新至 "+wk_No_Inbound,font=('黑体',10),bg='WhiteSmoke',
                              fg='DimGray',anchor='w',height=2)
    lb_week_Inbound.place(x=280,y=95)
    btn_updata = Button(frame,text='选择文件',command=updata_WeeklyInbound,fg='white',
                        font=("黑体",10,'bold'),bg='slategrey',width=9,height=1,borderwidth=5)
    btn_updata.place(x=180,y=96)
    
    # 订货数据上传
    def update_WeeklyOrder():
        try:
            filename = tkinter.filedialog.askopenfilename().replace("/","\\")
            weekly_order = pd.read_excel(filename)
            weekly_order.rename(columns={"规格型号":"Material","数量":"Orderweek_QTY"},
                                  inplace=True)
            jnj_date = filename[filename.rfind("_")+1:filename.rfind("_")+7]
            week_No = filename[filename.rfind("_")+7:filename.rfind("_")+9]
            weekly_order['JNJ_Date'] = jnj_date
            weekly_order['week_No'] = week_No
            weekly_order.drop(weekly_order[weekly_order["Material"].isnull()].index,
                                  inplace=True)#删除空行
            if week_No not in ['W1','W2','W3','W4','W5'] or "订货" not in filename:
                tkinter.messagebox.showerror("错误","请检查文件命名是否正确！")
            else:
                # 提前聚合
                weekly_order = weekly_order.groupby(by=['JNJ_Date', 'Material','week_No'],
                                                            as_index=False).sum()
                weekly_order = weekly_order[["JNJ_Date","Material","Orderweek_QTY",
                                                     'week_No']]
                # 若已存在则删除后替换
                sql_wk = "SELECT week_No FROM WeeklyOrder WHERE JNJ_Date = '"+jnj_date+"';"
                wk_no = list(PrismDatabaseOperation.Prism_select(sql_wk)["week_No"].unique())
                # 读取主数据
                SQL_ProductMaster_all = "SELECT Material From ProductMaster;"
                ProductMaster_all = PrismDatabaseOperation.Prism_select(SQL_ProductMaster_all)
                # 检查是否都在ProductMaster里
                lack_material = weekly_order[~weekly_order["Material"].isin(
                    ProductMaster_all["Material"])]
                if lack_material.empty:
                    if (week_No in wk_no) and (len(wk_no) != 0):
                        SQL_delete = "DELETE FROM WeeklyOrder WHERE JNJ_Date = '"+                         jnj_date+"' AND week_No = '"+week_No+"';"
                        PrismDatabaseOperation.Prism_delete(SQL_delete)
                        PrismDatabaseOperation.Prism_insert("WeeklyOrder",weekly_order)
                        tkinter.messagebox.showinfo("成功","导入成功！")
                    else:
                        PrismDatabaseOperation.Prism_insert("WeeklyOrder",weekly_order)
                        tkinter.messagebox.showinfo("成功","导入成功！")
                    # wk_No，WeeklyOrder
                    SQL_Order = "SELECT week_No FROM WeeklyOrder WHERE JNJ_Date = '"+                     next_month+"';"
                    wk_No_Order = sorted(list(
                        PrismDatabaseOperation.Prism_select(SQL_Order)["week_No"].unique()))[-1]
                    lb_week_Order = Label(frame,text="更新至 "+wk_No_Order,font=('黑体',10),
                                          bg='WhiteSmoke',fg='DimGray',anchor='w',height=2)
                    lb_week_Order.place(x=280,y=135)
#                     acl_rolling()
                else:
                    # 维护主数据
                    to_updata = pd.DataFrame(data={},columns=['规格型号','包装规格','分类Level3',
                                                              '分类Level4','ABC','分类',
                                                              '不含税单价','预测状态','MOQ',
                                                              '安全库存天数'])
                    to_updata["规格型号"] = lack_material['Material']
                    to_updata["预测状态"] = "MTS"
                    MasterData.master_update_batch(to_updata)
        except:
            tkinter.messagebox.showerror("错误","更新数据失败，请重新上传！")
    
    lb_WeeklyOrder = Label(frame,text='订货数据上传：',font=('黑体',12),bg='WhiteSmoke',
                               fg='DimGray',anchor='w',width=30,height=2)
    lb_WeeklyOrder.place(x=10,y=130)
    # wk_No，WeeklyOrder
    SQL_Order = "SELECT week_No FROM WeeklyOrder WHERE JNJ_Date = '"+next_month+"';"
    try:
        wk_No_Order = sorted(list(PrismDatabaseOperation.Prism_select(SQL_Order)["week_No"].unique()))[-1]
    except:
        wk_No_Order = "无"
    lb_week_Order = Label(frame,text="更新至 "+wk_No_Order,font=('黑体',10),bg='WhiteSmoke',
                              fg='DimGray',anchor='w',height=2)
    lb_week_Order.place(x=280,y=135)
    btn_WeeklyOrder = Button(frame,text='选择文件',command=update_WeeklyOrder,fg='white',
                             font=("黑体",10,'bold'),bg='slategrey',width=9,height=1,borderwidth=5)
    btn_WeeklyOrder.place(x=180,y=136)
    
    # 直接读取计划数据
    def read_rolling():
        # 缺失数据
        lack = []
        # 读取并合并数据
        # 导入主数据
        SQL_ProductMaster = "SELECT Material,GTS,FCST_state,ABC,MOQ From ProductMaster "
        ProductMaster = PrismDatabaseOperation.Prism_select(SQL_ProductMaster)
        ProductMaster.rename(columns={"ABC":"Class"},inplace=True)

        # 下个月时间 
        next_month = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                      relativedelta(months=1)).strftime("%Y%m")
        # 获取补货权值weekly_pattern
        WeeklyPattern = PrismDatabaseOperation.Prism_select("SELECT * FROM WeeklyPattern")
        WK1 = WeeklyPattern[WeeklyPattern['week']=='WK1']['pattern'].iloc[0]
        WK2 = WeeklyPattern[WeeklyPattern['week']=='WK2']['pattern'].iloc[0]
        WK3 = WeeklyPattern[WeeklyPattern['week']=='WK3']['pattern'].iloc[0]
        WK4 = WeeklyPattern[WeeklyPattern['week']=='WK4']['pattern'].iloc[0]
        if WeeklyPattern.empty:
            lack.append("补货权值")

        # wk_No，以WeeklyOutbound的周数为准
        SQL_week = "SELECT week_No FROM WeeklyOutbound WHERE JNJ_Date = '"+next_month+"';"
        wk_No = sorted(list(PrismDatabaseOperation.Prism_select(SQL_week)["week_No"].unique()))[-1]

        # ⭐每周获取出货数据
        SQL_WeeklyOutbound = "SELECT Material,Outboundweek_QTY,week_No From "+         "WeeklyOutbound WHERE JNJ_Date = '"+next_month+"';"
        WeeklyOutbound = PrismDatabaseOperation.Prism_select(SQL_WeeklyOutbound)
        if WeeklyOutbound.empty:
            lack.append("出货")

        # ⭐每周下单量数据
        SQL_WeeklyOrder = "SELECT Material,Orderweek_QTY,week_No FROM "+         "WeeklyOrder WHERE JNJ_Date = '"+next_month+"';"
        WeeklyOrder = PrismDatabaseOperation.Prism_select(SQL_WeeklyOrder)
        if WeeklyOrder.empty:
            lack.append("每周下单数据")

        # ⭐每周取最新版的缺货数据
        SQL_Backorder = "SELECT Material,Backorderweek_QTY,week_No FROM "+         "WeeklyBackorder WHERE JNJ_Date = '"+next_month+"';"
        WeeklyBackorder = PrismDatabaseOperation.Prism_select(SQL_Backorder)
        if WeeklyBackorder.empty:
            lack.append("每周缺货数据")

        # 调整后的需求数据
        SQL_AdjustFCSTDemand = "SELECT Material,FCST_Demand1 From AdjustFCSTDemand"+         " WHERE JNJ_Date = '"+JNJ_Month(1)[0]+"';"
        AdjustFCSTDemand = PrismDatabaseOperation.Prism_select(SQL_AdjustFCSTDemand)

        # 安全库存
        # 出库数据3个月
        SQL_Outbound = "SELECT * From Outbound WHERE JNJ_Date = '"+JNJ_Month(3)[0]+         "' OR JNJ_Date = '"+JNJ_Month(3)[1]+"' OR JNJ_Date = '"+JNJ_Month(3)[2]+"';"
        Outbound = PrismDatabaseOperation.Prism_select(SQL_Outbound)

        Outbound = Outbound.pivot_table(index='Material',columns='JNJ_Date')
        Outbound.reset_index(inplace=True)
        Outbound_QTY_col = [] # 换列名，方便直接取出相应月份的数据
        for i in range(4):
            Outbound_QTY_col.append(Outbound.columns.values[i][1])
        Outbound_QTY_col[0] = 'Material'
        Outbound.columns = Outbound_QTY_col

        # 安全库存天数
        SQL_SafetyStockDay = "SELECT [Class],[Safetystock_Day] From SafetyStockDay;"
        SafetyStockDay = PrismDatabaseOperation.Prism_select(SQL_SafetyStockDay)
        if SafetyStockDay.empty:
            lack.append("安全库存天数")

        # 上月Intransit
        SQL_Intransit = "SELECT Material,Intransit_QTY From Intransit WHERE JNJ_Date = '"+         JNJ_Month(1)[0]+"';"
        Intransit = PrismDatabaseOperation.Prism_select(SQL_Intransit)

        # 上月Onhand_QTY
        SQL_Onhand = "SELECT Material,Onhand_QTY From Onhand WHERE JNJ_Date = '"+         JNJ_Month(1)[0]+"';"
        Onhand = PrismDatabaseOperation.Prism_select(SQL_Onhand)

        # 上月Putaway
        SQL_Putaway = "SELECT Material,Putaway_QTY From Putaway WHERE JNJ_Date = '"+         JNJ_Month(1)[0]+"';"
        Putaway = PrismDatabaseOperation.Prism_select(SQL_Putaway)

        # 读取补货计划
        SQL_Rep_plan = "SELECT Material,RepWeek_QTY,week_No FROM"+         " AdjustRollingRepPlan WHERE JNJ_Date = '"+next_month+"';"
        Rep_plan = PrismDatabaseOperation.Prism_select(SQL_Rep_plan)
        if Rep_plan.empty:
            lack.append("补货计划")

        # 如有数缺失则显示
        if lack != []:
            tkinter.messagebox.showerror("警告",str(lack)+"数据缺失！")

        # 合并数据--上个月的数据合并，并计算安全库存
        merge_0 = pd.merge(AdjustFCSTDemand,Onhand,on=["Material"],how='outer')
        merge_1 = pd.merge(merge_0,Intransit,on=["Material"],how='outer')
        merge_2 = pd.merge(merge_1,Putaway,on=["Material"],how='outer')
        merge_3 = pd.merge(merge_2,Outbound,on=["Material"],how='outer')
        merge_4 = pd.merge(merge_3,ProductMaster,on=["Material"],how='outer')
        merge_last = pd.merge(merge_4,SafetyStockDay,on=["Class"],how='outer')
        # 计算安全库存量(四舍五入取整)
        merge_last.fillna(0,inplace=True)
        merge_last['Safetystock_QTY'] = 0
        month_1 = JNJ_Month(3)[0]
        month_2 = JNJ_Month(3)[1]
        month_3 = JNJ_Month(3)[2]
        for i in range(len(merge_last)):
            merge_last['Safetystock_QTY'].iloc[i] = new_round(
                merge_last['Safetystock_Day'].iloc[i]*(merge_last[month_1].iloc[i]+
                                                       merge_last[month_2].iloc[i]+
                                                       merge_last[month_3].iloc[i])/90,0)

        # 每周的数据更新，并更换列名
        merge_5 = pd.merge(WeeklyOutbound,Rep_plan,
                           on=["Material","week_No"],
                           how='outer')
        merge_6 = pd.merge(merge_5,WeeklyBackorder,
                           on=["Material","week_No"],
                           how='outer')
        merge_7 = pd.merge(merge_6,WeeklyOrder,
                           on=["Material","week_No"],
                           how='outer')
        merge_8 = merge_7.pivot_table(index='Material',columns='week_No').reset_index()

        merge_colunms = []
        for i in range(len(merge_8.columns)):
            merge_colunms.append(merge_8.columns[i][0]+merge_8.columns[i][1])
        merge_8.columns = merge_colunms

        merge_all = pd.merge(merge_8,merge_last,on=["Material"],how='outer')
        # 只选取预测状态为MTS的code
        merge_all = merge_all[merge_all["FCST_state"]=="MTS"]
        merge_all.fillna(0,inplace=True)

        # 当Wk不同时输出不同的补货计划
        merge_all["Rolling_RepW1_QTY"] = merge_all["RepWeek_QTYW1"]
        merge_all["Rolling_RepW2_QTY"] = merge_all["RepWeek_QTYW2"]
        merge_all["Rolling_RepW3_QTY"] = merge_all["RepWeek_QTYW3"]
        merge_all["Rolling_RepW4_QTY"] = merge_all["RepWeek_QTYW4"]
        # 当前缺货数
        
        merge_all["Backorderweek_QTY"] = merge_all["Backorderweek_QTYW1"]
        # 已出货量
        merge_all["Outboundweek_QTY"] = merge_all["Outboundweek_QTYW1"]

        # 计算补货计划值
        merge_all["Rolling_Rep_QTY"] = (merge_all["Rolling_RepW1_QTY"]+
                                          merge_all["Rolling_RepW2_QTY"]+
                                          merge_all["Rolling_RepW3_QTY"]+
                                          merge_all["Rolling_RepW4_QTY"])
        merge_all["Rolling_Rep_value"] = merge_all["Rolling_Rep_QTY"]*merge_all["GTS"]
        merge_all.fillna(0,inplace=True)
        
        acl_rolling(merge_all)
    
    # 读取并计算rolling
    def read_acl_rolling():
        # 缺失数据
        lack = []
        # 读取并合并数据
        # 导入主数据
        SQL_ProductMaster = "SELECT Material,GTS,FCST_state,ABC,MOQ From ProductMaster "
        ProductMaster = PrismDatabaseOperation.Prism_select(SQL_ProductMaster)
        ProductMaster.rename(columns={"ABC":"Class"},inplace=True)

        # 下个月时间 
        next_month = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                      relativedelta(months=1)).strftime("%Y%m")
        # 获取补货权值weekly_pattern
        WeeklyPattern = PrismDatabaseOperation.Prism_select("SELECT * FROM WeeklyPattern")
        WK1 = WeeklyPattern[WeeklyPattern['week']=='WK1']['pattern'].iloc[0]
        WK2 = WeeklyPattern[WeeklyPattern['week']=='WK2']['pattern'].iloc[0]
        WK3 = WeeklyPattern[WeeklyPattern['week']=='WK3']['pattern'].iloc[0]
        WK4 = WeeklyPattern[WeeklyPattern['week']=='WK4']['pattern'].iloc[0]
        if WeeklyPattern.empty:
            lack.append("补货权值")

        # wk_No，以WeeklyOutbound的周数为准
        SQL_week = "SELECT week_No FROM WeeklyOutbound WHERE JNJ_Date = '"+next_month+"';"
        try:
            wk_No = sorted(list(PrismDatabaseOperation.Prism_select(SQL_week)["week_No"].unique()))[-1]
        except:
            wk_No = "W1"

        # ⭐每周获取出货数据
        SQL_WeeklyOutbound = "SELECT Material,Outboundweek_QTY,week_No From "+         "WeeklyOutbound WHERE JNJ_Date = '"+next_month+"';"
        WeeklyOutbound = PrismDatabaseOperation.Prism_select(SQL_WeeklyOutbound)
        if WeeklyOutbound.empty:
            lack.append("出货")

        # ⭐每周下单量数据
        SQL_WeeklyOrder = "SELECT Material,Orderweek_QTY,week_No FROM "+         "WeeklyOrder WHERE JNJ_Date = '"+next_month+"';"
        WeeklyOrder = PrismDatabaseOperation.Prism_select(SQL_WeeklyOrder)
        if WeeklyOrder.empty:
            lack.append("每周下单数据")

        # ⭐每周取最新版的缺货数据
        SQL_Backorder = "SELECT Material,Backorderweek_QTY,week_No FROM "+         "WeeklyBackorder WHERE JNJ_Date = '"+next_month+"';"
        WeeklyBackorder = PrismDatabaseOperation.Prism_select(SQL_Backorder)
        if WeeklyBackorder.empty:
            lack.append("每周缺货数据")

        # 调整后的需求数据
        SQL_AdjustFCSTDemand = "SELECT Material,FCST_Demand1 From AdjustFCSTDemand"+         " WHERE JNJ_Date = '"+JNJ_Month(1)[0]+"';"
        AdjustFCSTDemand = PrismDatabaseOperation.Prism_select(SQL_AdjustFCSTDemand)
#         if AdjustFCSTDemand.empty:
#             lack.append("二级需求")

        # 安全库存
        # 出库数据3个月
        SQL_Outbound = "SELECT * From Outbound WHERE JNJ_Date = '"+JNJ_Month(3)[0]+         "' OR JNJ_Date = '"+JNJ_Month(3)[1]+"' OR JNJ_Date = '"+JNJ_Month(3)[2]+"';"
        Outbound = PrismDatabaseOperation.Prism_select(SQL_Outbound)
#         if Outbound.empty:
#             lack.append("出库数据")
        Outbound = Outbound.pivot_table(index='Material',columns='JNJ_Date')
        Outbound.reset_index(inplace=True)
        Outbound_QTY_col = [] # 换列名，方便直接取出相应月份的数据
        for i in range(4):
            Outbound_QTY_col.append(Outbound.columns.values[i][1])
        Outbound_QTY_col[0] = 'Material'
        Outbound.columns = Outbound_QTY_col

        # 安全库存天数
        SQL_SafetyStockDay = "SELECT [Class],[Safetystock_Day] From SafetyStockDay;"
        SafetyStockDay = PrismDatabaseOperation.Prism_select(SQL_SafetyStockDay)
        if SafetyStockDay.empty:
            lack.append("安全库存天数")

        # 上月Intransit
        SQL_Intransit = "SELECT Material,Intransit_QTY From Intransit WHERE JNJ_Date = '"+         JNJ_Month(1)[0]+"';"
        Intransit = PrismDatabaseOperation.Prism_select(SQL_Intransit)
#         if Intransit.empty:
#             lack.append("上个月在途")

        # 上月Onhand_QTY
        SQL_Onhand = "SELECT Material,Onhand_QTY From Onhand WHERE JNJ_Date = '"+         JNJ_Month(1)[0]+"';"
        Onhand = PrismDatabaseOperation.Prism_select(SQL_Onhand)
#         if Onhand.empty:
#             lack.append("上个月可发")

        # 上月Putaway
        SQL_Putaway = "SELECT Material,Putaway_QTY From Putaway WHERE JNJ_Date = '"+         JNJ_Month(1)[0]+"';"
        Putaway = PrismDatabaseOperation.Prism_select(SQL_Putaway)
#         if Putaway.empty:
#             lack.append("上个月预入库")

        # 读取补货计划
        SQL_Rep_plan = "SELECT Material,RepWeek_QTY,week_No FROM"+         " AdjustRollingRepPlan WHERE JNJ_Date = '"+next_month+"';"
        Rep_plan = PrismDatabaseOperation.Prism_select(SQL_Rep_plan)
        if Rep_plan.empty:
            lack.append("补货计划")

        # 如有数缺失则显示
        if lack != []:
            tkinter.messagebox.showerror("警告",str(lack)+"数据缺失！")

        # 合并数据--上个月的数据合并，并计算安全库存
        merge_0 = pd.merge(AdjustFCSTDemand,Onhand,on=["Material"],how='outer')
        merge_1 = pd.merge(merge_0,Intransit,on=["Material"],how='outer')
        merge_2 = pd.merge(merge_1,Putaway,on=["Material"],how='outer')
        merge_3 = pd.merge(merge_2,Outbound,on=["Material"],how='outer')
        merge_4 = pd.merge(merge_3,ProductMaster,on=["Material"],how='outer')
        merge_last = pd.merge(merge_4,SafetyStockDay,on=["Class"],how='outer')
        # 计算安全库存量(四舍五入取整)
        merge_last.fillna(0,inplace=True)
        merge_last['Safetystock_QTY'] = 0
        month_1 = JNJ_Month(3)[0]
        month_2 = JNJ_Month(3)[1]
        month_3 = JNJ_Month(3)[2]
        for i in range(len(merge_last)):
            merge_last['Safetystock_QTY'].iloc[i] = new_round(
                merge_last['Safetystock_Day'].iloc[i]*(merge_last[month_1].iloc[i]+
                                                       merge_last[month_2].iloc[i]+
                                                       merge_last[month_3].iloc[i])/90,0)

        # 每周的数据更新，并更换列名
        merge_5 = pd.merge(WeeklyOutbound,Rep_plan,
                           on=["Material","week_No"],
                           how='outer')
        merge_6 = pd.merge(merge_5,WeeklyBackorder,
                           on=["Material","week_No"],
                           how='outer')
        merge_7 = pd.merge(merge_6,WeeklyOrder,
                           on=["Material","week_No"],
                           how='outer')
        merge_8 = merge_7.pivot_table(index='Material',columns='week_No').reset_index()

        merge_colunms = []
        for i in range(len(merge_8.columns)):
            merge_colunms.append(merge_8.columns[i][0]+merge_8.columns[i][1])
        merge_8.columns = merge_colunms

        merge_all = pd.merge(merge_8,merge_last,on=["Material"],how='outer')
        # 只选取预测状态为MTS的code
        merge_all = merge_all[merge_all["FCST_state"]=="MTS"]
        merge_all.fillna(0,inplace=True)

        # 待确认数据梳理部分 
        # 当Wk不同时输出不同的补货计划
        if wk_No == "W1":
            merge_all["Rolling_RepW1_QTY"] = merge_all["Orderweek_QTYW1"]
            merge_all["Rolling_RepW2_QTY"] = new_round((WK2*merge_all["FCST_Demand1"]+
                                                       merge_all["Backorderweek_QTYW1"]+
                                                       merge_all["Safetystock_QTY"]-
                                          (merge_all["Onhand_QTY"]+merge_all["Intransit_QTY"]+
                                            merge_all["Putaway_QTY"]+merge_all['Orderweek_QTYW1']
                                            -merge_all['Outboundweek_QTYW1']))/
                                            merge_all['MOQ'],0)*merge_all['MOQ']
            # 小于0则替换成0
            merge_all.loc[merge_all["Rolling_RepW2_QTY"]<=0,"Rolling_RepW2_QTY"] = 0
            merge_all["Rolling_RepW3_QTY"] = new_round(WK3*merge_all["FCST_Demand1"]/
                                                       merge_all['MOQ'],0)*merge_all['MOQ']
            merge_all["Rolling_RepW4_QTY"] = new_round(WK4*merge_all["FCST_Demand1"]/
                                                       merge_all['MOQ'],0)*merge_all['MOQ']
            # 当前缺货数
            merge_all["Backorderweek_QTY"] = merge_all["Backorderweek_QTYW1"]
            # 已出货量
            merge_all["Outboundweek_QTY"] = merge_all["Outboundweek_QTYW1"]

        elif wk_No == "W2":
            merge_all["Rolling_RepW1_QTY"] = merge_all["Orderweek_QTYW1"]
            merge_all["Rolling_RepW2_QTY"] = merge_all["Orderweek_QTYW2"]
            merge_all["Rolling_RepW3_QTY"] = new_round((WK3*merge_all["FCST_Demand1"]+
                                                       merge_all["Backorderweek_QTYW2"]+
                                                       merge_all["Safetystock_QTY"]-
                                          (merge_all["Onhand_QTY"]+merge_all["Intransit_QTY"]+
                                            merge_all["Putaway_QTY"]+merge_all['Orderweek_QTYW1']+
                                            merge_all['Orderweek_QTYW2']-
                                            merge_all['Outboundweek_QTYW1']-
                                            merge_all['Outboundweek_QTYW2']))/
                                            merge_all['MOQ'],0)*merge_all['MOQ']
            # 小于0则替换成0
            merge_all.loc[merge_all["Rolling_RepW3_QTY"]<=0,"Rolling_RepW3_QTY"] = 0
            merge_all["Rolling_RepW4_QTY"] = new_round(WK4*merge_all["FCST_Demand1"]
                                                       /merge_all['MOQ'],0)*merge_all['MOQ']
            # 当前缺货数
            merge_all["Backorderweek_QTY"] = merge_all["Backorderweek_QTYW2"]
            # 已出货量
            merge_all["Outboundweek_QTY"] = merge_all["Outboundweek_QTYW1"]+             merge_all['Outboundweek_QTYW2']

        elif wk_No == "W3":
            merge_all["Rolling_RepW1_QTY"] = merge_all["Orderweek_QTYW1"]
            merge_all["Rolling_RepW2_QTY"] = merge_all["Orderweek_QTYW2"]
            merge_all["Rolling_RepW3_QTY"] = merge_all["Orderweek_QTYW3"]
            merge_all["Rolling_RepW4_QTY"] = new_round((WK4*merge_all["FCST_Demand1"]+
                                                       merge_all["Backorderweek_QTYW3"]+
                                                       merge_all["Safetystock_QTY"]-
                                              (merge_all["Onhand_QTY"]+
                                               merge_all["Intransit_QTY"]+
                                                merge_all["Putaway_QTY"]+
                                               merge_all['Orderweek_QTYW1']+
                                                merge_all['Orderweek_QTYW2']+
                                                merge_all['Orderweek_QTYW3']-
                                                merge_all['Outboundweek_QTYW1']-
                                                merge_all['Outboundweek_QTYW2']-
                                                merge_all['Outboundweek_QTYW3']))/
                                                merge_all['MOQ'],0)*merge_all['MOQ']
            # 小于0则替换成0
            merge_all.loc[merge_all["Rolling_RepW4_QTY"]<=0,"Rolling_RepW4_QTY"] = 0
            # 当前缺货数
            merge_all["Backorderweek_QTY"] = merge_all["Backorderweek_QTYW3"]
            # 已出货量
            merge_all["Outboundweek_QTY"] = merge_all["Outboundweek_QTYW1"]+             merge_all['Outboundweek_QTYW2']+merge_all['Outboundweek_QTYW3']
        
        elif wk_No == "W4":
            merge_all["Rolling_RepW1_QTY"] = merge_all["Orderweek_QTYW1"]
            merge_all["Rolling_RepW2_QTY"] = merge_all["Orderweek_QTYW2"]
            merge_all["Rolling_RepW3_QTY"] = merge_all["Orderweek_QTYW3"]
            merge_all["Rolling_RepW4_QTY"] = merge_all["Orderweek_QTYW4"]
            # 当前缺货数
            merge_all["Backorderweek_QTY"] = merge_all["Backorderweek_QTYW4"]
            # 已出货量
            merge_all["Outboundweek_QTY"] = merge_all["Outboundweek_QTYW1"]+             merge_all['Outboundweek_QTYW2']+merge_all['Outboundweek_QTYW3']+             merge_all['Outboundweek_QTYW4']
            
        else:
            tkinter.messagebox.showinfo("提示","数据缺失")
        # 计算补货计划值
        merge_all["Rolling_Rep_QTY"] = (merge_all["Rolling_RepW1_QTY"]+
                                          merge_all["Rolling_RepW2_QTY"]+
                                          merge_all["Rolling_RepW3_QTY"]+
                                          merge_all["Rolling_RepW4_QTY"])
        merge_all["Rolling_Rep_value"] = merge_all["Rolling_Rep_QTY"]*merge_all["GTS"]
        merge_all.fillna(0,inplace=True)
        # ⭐存储，需先逆透视，再加上JNJ_Date和remark才行,取消计算功能，当存在数据时，先读取数据
        # 点击更新按钮，覆盖数据，再做保存
        RollingRep = merge_all[["Material","Rolling_RepW1_QTY","Rolling_RepW2_QTY",
                                "Rolling_RepW3_QTY","Rolling_RepW4_QTY"]]
        RollingRep.rename(columns={"Rolling_RepW1_QTY":"W1","Rolling_RepW2_QTY":"W2",
                                "Rolling_RepW3_QTY":"W3","Rolling_RepW4_QTY":"W4"},
                          inplace=True)
        RollingRep = RollingRep.melt(id_vars =['Material'], var_name = 'week_No', 
                                 value_name = 'RepWeek_QTY')
        RollingRep["JNJ_Date"] = next_month
        RollingRep["Rep_Remark"] = ""
        # 判断当前数据库中的月份
        sql_jnj_date = "SELECT JNJ_Date FROM AdjustRollingRepPlan"
        jnj_date = list(PrismDatabaseOperation.Prism_select(sql_jnj_date)['JNJ_Date'])
        # 若已有数据则直接覆盖
        if next_month in jnj_date:
            SQL_delete = "DELETE FROM AdjustRollingRepPlan WHERE JNJ_Date = '"+next_month+"';"
            PrismDatabaseOperation.Prism_delete(SQL_delete)
            PrismDatabaseOperation.Prism_insert('AdjustRollingRepPlan',RollingRep)
        else:
            PrismDatabaseOperation.Prism_insert('AdjustRollingRepPlan',RollingRep)
        
        acl_rolling(merge_all)
    
    btn_acl_rolling = Button(frame,text="更新补货计划",command=read_acl_rolling,font=('黑体',12,'bold'),
                             width=15,height=1,bg='slategrey',fg='white',borderwidth=5)
    btn_acl_rolling.place(x=820,y=130)

    # 计算rolling数据，通过周办法计算
    def acl_rolling(merge_data):
        # 获取数据
        merge_all = merge_data
        
        # 获取指标
        SQL_OrderTarget = "SELECT * FROM OrderTarget WHERE JNJ_Date = '"+next_month+"';"
        OrderTarget = PrismDatabaseOperation.Prism_select(SQL_OrderTarget)
            
        #**********************界面显示**********************#
        
        # 截取所需显示的内容
        rep_result = merge_all[["Material","Rolling_Rep_QTY","Rolling_RepW1_QTY",
                               "Rolling_RepW2_QTY","Rolling_RepW3_QTY",
                               "Rolling_RepW4_QTY","Backorderweek_QTY",
                               "Outboundweek_QTY","FCST_Demand1"]]
        
        # 重命名
        rep_result.rename(columns={"Material":"规格型号","Rolling_Rep_QTY":"月补货量",
                                   "Rolling_RepW1_QTY":"W1 ","Rolling_RepW2_QTY":"W2",
                                   "Rolling_RepW3_QTY":"W3 ","Rolling_RepW4_QTY":"W4",
                                   "Backorderweek_QTY":"当前缺货","Outboundweek_QTY":"已出货量",
                                   "FCST_Demand1":"月需求"},
                          inplace=True)
        # 当前补货金额计算、月度补货金额
        Label(frame,text="月度指标金额：",font=('华文中宋',14),width=12,height=1,bg='WhiteSmoke'
             ).place(x=675,y=10)
        Label(frame,text="当前补货金额：",font=('华文中宋',14),width=12,height=1,bg='WhiteSmoke'
             ).place(x=675,y=55)
        # 指标
        if next_month in list(OrderTarget['JNJ_Date']):
            Target_amount = OrderTarget['order_target'].iloc[0]
            lb_target_amount = Label(frame,text=re_round(Target_amount),font=('黑体',15),
                                     width=15,height=2,anchor="e",bg='WhiteSmoke')
            lb_target_amount.place(x=820,y=3)
        else:
            Target_amount = 0
            lb_target_amount = Label(frame,text=re_round(0),font=('黑体',15),width=15,height=2,
                                     anchor="e",bg='WhiteSmoke')
            lb_target_amount.place(x=820,y=3)
        
        # 双击提示
        CreateToolTip(lb_target_amount, "双击此处即可编辑")
        
        # 判断颜色，大于指标红色，小于指标绿色
        merge_Rep_value = sum(merge_all["Rolling_Rep_value"])
        if Target_amount > merge_Rep_value:
            lb_amount = Label(frame,text=re_round(merge_Rep_value),anchor="e",
                              font=('黑体',15),width=15,height=2,fg='green',bg='WhiteSmoke')
            lb_amount.place(x=820,y=46)
        else:
            lb_amount = Label(frame,text=re_round(merge_Rep_value),anchor="e",
                              font=('黑体',15),width=15,height=2,fg='red',bg='WhiteSmoke')
            lb_amount.place(x=820,y=46)
        
        # 将得到的计算结果展示在界面
        columns = list(rep_result.columns)

        # 设置样式
        style_head = ttk.Style()
        style_head.configure("MyStyle.Treeview.Heading",rowheight=50,font=("华文中宋",12))
        style_value = ttk.Style()
        style_value.configure("MyStyle.Treeview", rowheight=24)
        treeview = ttk.Treeview(frame, height=16, show="headings",selectmode="extended",
                                columns=columns,style='MyStyle.Treeview')

        # 添加滚动条
        # 竖向滚动条
        sb_y = ttk.Scrollbar(frame,command=treeview.yview)
        sb_y.config(command=treeview.yview)
        sb_y.place(in_=treeview,relx=1, rely=0,relheight=1)
        treeview.config(yscrollcommand=sb_y.set)
        # 横向滚动条
        sb_x = ttk.Scrollbar(frame,command=treeview.xview,orient="horizontal")
        sb_x.config(command=treeview.xview)
        sb_x.place(in_=treeview,relx=0, rely=1,relwidth=1)
        treeview.config(xscrollcommand=sb_x.set)
        treeview.place(x=0,y=210,relwidth=0.98)

        # Tips
        Label(frame,text="* 操作提示:双击相应数据可以且仅能编辑未上传下单量的周补货量",
              font=("黑体",10),bg='WhiteSmoke').place(in_=treeview,x=0,y=425)
        # 表示列,不显示
        for i in range(0,len(rep_result.columns)):
            treeview.column(str(rep_result.columns[i]), width=100, anchor='center') 

        # 显示表头
        for i in range(len(rep_result.columns)):
            treeview.heading(str(rep_result.columns[i]), text=str(rep_result.columns[i]))

        # 行交替颜色
        def fixed_map(option):# 重要！无此步骤则无法显示
            return [elm for elm in style.map("Treeview", query_opt=option)
                    if elm[:2] != ("!disabled", "!selected")]
        style = ttk.Style()
        style.map("Treeview",foreground=fixed_map("foreground"),background=fixed_map("background"))

        treeview.tag_configure('oddrow', background='LightGrey')
        treeview.tag_configure('evenrow', background='white')

        # 行坐标重排
        def odd_even_color():
            for index,row in enumerate(treeview.get_children()):
                if index % 2 == 0:
                    treeview.item(row,tags="evenrow")
                else:
                    treeview.item(row,tags="oddrow")
            
        # 插入数据，数字显示为千分位
        for i in range(len(rep_result)):
            if i % 2 == 0:
                tag = "evenrow"
            else:
                tag = "oddrow"
            treeview.insert('', i, 
                              values=(rep_result[rep_result.columns[0]].iloc[i],
                    "{:,}".format(int(rep_result[rep_result.columns[1]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[2]].iloc[i])),   
                    "{:,}".format(int(rep_result[rep_result.columns[3]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[4]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[5]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[6]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[7]].iloc[i])),
                    "{:,}".format(int(rep_result[rep_result.columns[8]].iloc[i])))
                           ,tags=tag)

        # Treeview、列名、排列方式
        def treeview_sort_column(tv, col, reverse):  
            L = [(tv.set(k, col), k) for k in tv.get_children('')]
            try:
                for i in range(len(L)):
                    L[i] = (float(L[i][0].replace(',', '')),L[i][1])
            except:
                pass
            L.sort(reverse=reverse)  # 排序方式
            # 根据排序后索引移动
            for index, (val, k) in enumerate(L):
                tv.move(k, '', index)
            # 重写标题，使之成为再点倒序的标题
            tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))
            odd_even_color()

        # 绑定函数，使表头可排序
        for col in columns:
            treeview.heading(col, text=col, command=
                             lambda _col=col: treeview_sort_column(treeview, _col, False))
        
        # 选择路径，输出保存
        def output_plan():
            filename = tkinter.filedialog.asksaveasfilename()
            # 遍历获取所有数据，并生成df
            # 改变文本存储的数字
            t = treeview.get_children()
            a = list()
            for i in t:
                a.append(list(treeview.item(i,'values')))
            df_now = pd.DataFrame(a,columns=columns)
            # 按列名输出
            df_now = df_now[["规格型号","月补货量","W1","W2","W3","W4"]]
            # 指定列修改千分位为数字
            for i in range(0,len(df_now.columns)):
                try:
                    df_now[df_now.columns[i]] = df_now.loc[:,df_now.columns[i]].apply(
                        lambda x: float(x.replace(",", "")))
                except:
                    pass
            df_now.to_excel(filename+'.xls',index=False)
        
        btn_output = Button(frame,text="下载补货计划",font=('黑体',12,'bold'),width=15,height=1,
                            bg='slategrey',fg='white',borderwidth=5,command=output_plan)
        btn_output.place(x=820,y=170)
        
        # 搜索功能
        Label(frame,text="筛选字段：",bg='WhiteSmoke',font=("黑体",12)).place(x=5,y=180)
        cbx = ttk.Combobox(frame,font=("黑体",11),width=10) #筛选字段
        comvalue = tkinter.StringVar()
        cbx["values"] = ["全局搜索"] + columns
        cbx.current(1)
        cbx.place(x=85,y=180)
        entry_search = Entry(frame,font=("黑体",11),width=12) # 筛选内容
        entry_search.insert(0, "请输入信息")
        entry_search.place(x=195,y=180)

        # 先清空表格，再插入数据，当字段选择为空、内容为空则显示全部
        def search_material():
            search_all = rep_result.copy()
            for i in search_all.columns:
                # 必须转字符，否则无法全局搜索
                try:
                    search_all[i] = search_all[i].map(lambda x:format(int(x),','))
                except:
                    search_all[i] = search_all[i].apply(str)

            # 清空
            for item in treeview.get_children():
                treeview.delete(item)
            # 查找并插入数据
            if entry_search.get() != "":
                search_content = str(entry_search.get())
                # 全局搜索
                if cbx.get() == "全局搜索":
                    search_df = pd.DataFrame(columns=search_all.columns)
                    for i in range(len(search_all.columns)):
                        search_df = search_df.append(search_all[search_all[
                            search_all.columns[i]].str.contains(search_content)])                
                    search_df.drop_duplicates(subset=["规格型号"], keep='first',inplace=True)
    #                 print(search_df)
                # 指定字段搜索
                else:
                    appoint = str(cbx.get())
                    search_df = search_all[search_all[appoint].str.contains(search_content)]
    #                 print(search_df)
                # 插入表格
                for i in range(len(search_df)):
                    if i % 2 == 0:
                        tag = "evenrow"
                    else:
                        tag = "oddrow"
                    treeview.insert('', i, values=list(search_df.iloc[i,:]),tags=tag)
            # 若输入值为空则显示全部内容
            else:
                # 插入
                for i in range(len(search_all)):
                    if i % 2 == 0:
                        tag = "evenrow"
                    else:
                        tag = "oddrow"
                    treeview.insert('', i, values=list(search_all.iloc[i,:]),tags=tag)

        btn_search_material = Button(frame,text="查找",font=("黑体",10,'bold'),bg='slategrey',
                                     fg='white',width=9,height=1,borderwidth=5,
                                     command=search_material)
        btn_search_material.place(x=315,y=175)
        
        # 输入预测月指标
        def order_target(event):
            entry_input = Entry(frame,font=('黑体',15),width=15)
            entry_input.place(x=820,y=10)
            # 插入数据库中的目标金额
            SQL_OrderTarget = "SELECT * FROM OrderTarget WHERE JNJ_Date = '"+next_month+"';"
            OrderTarget_amount = PrismDatabaseOperation.Prism_select(SQL_OrderTarget)
            entry_input.insert(0,OrderTarget_amount["order_target"].iloc[0])
            # 确认输入函数，并将输入数据更新至数据库
            def input_target_amount():
                # 保存数据库并以最后的标准为主，方法：先插入下个月信息，再更新输入信息
                # 如果已存在于数据库，则直接更新，否则新增空的过度变量再更新
                if next_month in list(
                    PrismDatabaseOperation.Prism_select("SELECT JNJ_Date FROM OrderTarget")["JNJ_Date"]):
                    # 覆盖输入数据
                    SQL_delete = "DELETE FROM OrderTarget WHERE JNJ_Date = '"+next_month+"';"
                    PrismDatabaseOperation.Prism_delete(SQL_delete)
                    OrderTarget = pd.DataFrame(data={'JNJ_Date':[next_month],
                                                     'order_target':[float(entry_input.get())]})
                    PrismDatabaseOperation.Prism_insert('OrderTarget',OrderTarget)
                else:
                    OrderTarget = pd.DataFrame(data={'JNJ_Date':[next_month],
                                                     'order_target':[float(entry_input.get())]})
                    PrismDatabaseOperation.Prism_insert('OrderTarget',OrderTarget)

                Target_amount = float(entry_input.get())
                lb_target_amount = Label(frame,text=re_round(Target_amount),anchor="e",
                                         font=('黑体',15),width=15,height=2,bg='WhiteSmoke')
                lb_target_amount.place(x=820,y=3)
                # 双击提示
                CreateToolTip(lb_target_amount, "双击此处即可编辑")
                lb_target_amount.bind('<Double-1>',order_target)
                # 输入指标后变色
                if Target_amount > merge_Rep_value:
                    lb_amount = Label(frame,text=re_round(merge_Rep_value),anchor="e",
                                      font=('黑体',15),width=15,height=2,fg='green',bg='WhiteSmoke')
                    lb_amount.place(x=820,y=46)
                else:
                    lb_amount = Label(frame,text=re_round(merge_Rep_value),anchor="e",
                                      font=('黑体',15),width=15,height=2,fg='red',bg='WhiteSmoke')
                    lb_amount.place(x=820,y=46)

                btn_input_cancel.destroy()
                btn_input_target.destroy()
                entry_input.destroy()

            btn_input_target = Button(frame,text="OK",command=input_target_amount)
            btn_input_target.place(x=820,y=37)
            # 输入取消
            def input_cancel():
                btn_input_cancel.destroy()
                btn_input_target.destroy()
                entry_input.destroy()

            btn_input_cancel = Button(frame,text="No",command=input_cancel)
            btn_input_cancel.place(x=942,y=37)
                
        # 双击即可弹出编辑信息
        lb_target_amount.bind('<Double-1>',order_target)
        # 修改补货计划
        def modify_rolling_plan(event):
            item_text = treeview.item(treeview.selection(), "values")
            plan = Tk()
            plan.title('修改')
            plan.geometry('1060x230')

            # Tips
            lb_tips = Label(plan,text="Tips:双击可以并且仅能修改每周的补货量",
                            font=("华文中宋",12),fg="brown")
            lb_tips.place(x=40,y=200)
            # 获取主数据列名
            columns = list(rep_result.columns)
            modify_treeview = ttk.Treeview(plan,height=1,show="headings",columns=columns)
            modify_treeview.place(x=20,y=20,relwidth=0.98)

            # 横向滚动条
            sb_x = ttk.Scrollbar(plan,command=modify_treeview.xview,orient="horizontal")
            sb_x.config(command=modify_treeview.xview)
            sb_x.place(in_=modify_treeview,relx=0, rely=1,relwidth=1)
            modify_treeview.config(xscrollcommand=sb_x.set)

            for i in range(len(columns)):
                modify_treeview.column(columns[i], width=95, anchor='center')

            # 显示列名
            for i in range(len(columns)):
                modify_treeview.heading(columns[i], text=columns[i])

            # 插入数据
            modify_treeview.insert('', 1, values=item_text)

            # 合并输入数据
            def set_value(event): 
                # 获取鼠标所选item
                for item in modify_treeview.selection():
                    item_text = modify_treeview.item(item, "values")

                column = modify_treeview.identify_column(event.x)# 所在列
                row = modify_treeview.identify_row(event.y)# 所在行，返回
                cn = int(str(column).replace('#',''))
                # wk_No，以WeeklyOutbound的周数为准
                SQL_week = "SELECT week_No FROM WeeklyOutbound WHERE JNJ_Date = '"+next_month+"';"
                wk_No = sorted(list(PrismDatabaseOperation.Prism_select(SQL_week)["week_No"].unique()))[-1]
                if wk_No == "W1":
                    if cn  not in [5,6,4]:
                        cn = 100
                elif wk_No == "W2":
                    if cn  not in [5,6]:
                        cn = 100
                elif wk_No == "W3":
                    if cn  not in [5]:
                        cn = 100
                entryedit = Entry(plan,width=10)
                entryedit.insert(0,str(item_text[cn-1]))
                entryedit.place(x=150, y=150)
                Label_select = Label(plan,text=str(item_text[cn-1]),width=20)
                Label_select.place(x=150, y=100)
                # 将编辑好的信息更新到数据库中
                def save_edit():
                    if cn in [3,4,5,6]:
                        entry_value = int(entryedit.get().replace(",",""))
                        modify_treeview.set(item, column=column,value="{:,}".format(entry_value)) 
                        item_4 = int(item_text[2].replace(",",""))
                        item_5 = int(item_text[3].replace(",",""))
                        item_6 = int(item_text[4].replace(",",""))
                        item_7 = int(item_text[5].replace(",",""))
                        item_n = int(item_text[cn-1].replace(",",""))
                        modify_treeview.set(item,column=1,
                                            value="{:,}".format(item_4+item_5+item_6+item_7+
                                                      entry_value-item_n))
                    else:
                        modify_treeview.set(item, column=column,value=entryedit.get())
                    entryedit.destroy()
                    btn_input.destroy()
                    btn_cancal.destroy()
                    Label_select.destroy()

                btn_input = Button(plan, text='OK', width=7, command=save_edit)
                btn_input.place(x=260,y=150)

                # 取消输入
                def cancal_edit():
                    entryedit.destroy()
                    btn_input.destroy()
                    btn_cancal.destroy()
                    Label_select.destroy()

                btn_cancal = Button(plan, text='Cancel', width=7, command=cancal_edit)
                btn_cancal.place(x=350,y=150)

            # 触发双击事件
            modify_treeview.bind('<Double-1>', set_value)

            Label(plan,text="修改前：").place(x=100,y=100)
            Label(plan,text="修改后：").place(x=100,y=150)

            # 将编辑好的信息更新到数据库中
            def db_update():
                # 获取所有最新数据转为df
                t = modify_treeview.get_children()
                a = list()
                for i in t:
                    a.append(list(modify_treeview.item(i,'values')))
                df_now = pd.DataFrame(a,columns=columns)
                # 更新数据库
                next_month = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                          relativedelta(months=1)).strftime("%Y%m") # 下个月时间 
                for i in range(4):
                    SQL_update = "UPDATE AdjustRollingRepPlan SET RepWeek_QTY = "+                     str(df_now.iloc[0,i+2].replace(",",""))+" WHERE JNJ_Date='"+next_month+                     "' AND Material='"+df_now["规格型号"].iloc[0]+"' AND week_No='"+                     df_now.columns[i+2]+"';"
#                     print(SQL_update)
                    PrismDatabaseOperation.Prism_update(SQL_update)
#                 # 更新备注
#                 SQL_update_remark = "UPDATE AdjustRollingRepPlan SET Rep_Remark = '"+ \
#                 str(df_now.iloc[0,8])+"' WHERE JNJ_Date='"+next_month+"' AND Material='"+ \
#                 df_now["规格型号"].iloc[0]+"' AND week_No='"+df_now.columns[4]+"';"
#                 Prism_db.Prism_update(SQL_update_remark)
                
                # 刷新界面,直接将修改好的信息复制到主界面
                for i in range(5):
                    treeview.set(treeview.selection(), column=i+1,value=df_now.iloc[0,i+1])
                
                # 重新计算value
                # 读取主数据
                SQL_ProductMaster = "SELECT Material,GTS,FCST_state,ABC,MOQ From ProductMaster "
                ProductMaster = PrismDatabaseOperation.Prism_select(SQL_ProductMaster)
                ProductMaster.rename(columns={"ABC":"Class"},inplace=True)
                # 读取补货计划
                next_month = (datetime.datetime.strptime(JNJ_Month(1)[0],"%Y%m")+
                                          relativedelta(months=1)).strftime("%Y%m") # 下个月时间
                SQL_AdjustRepPlan = "SELECT * From AdjustRollingRepPlan WHERE JNJ_Date= "+                 next_month+";"
                AdjustRepPlan = PrismDatabaseOperation.Prism_select(SQL_AdjustRepPlan)

                # 数透
                AdjustRepPlan = pd.pivot_table(AdjustRepPlan,index=["Material","JNJ_Date"],
                                               columns=["week_No"],
                                               values=["RepWeek_QTY"])
                AdjustRepPlan_col = []
                for i in range(len(AdjustRepPlan.columns)):
                    if type(AdjustRepPlan.columns.values[i][1]) == str :
                        AdjustRepPlan_col.append(AdjustRepPlan.columns.values[i][1])
                    else:
                        AdjustRepPlan_col.append(str(AdjustRepPlan.columns.values[i][1]))
                AdjustRepPlan.columns = AdjustRepPlan_col
                AdjustRepPlan = AdjustRepPlan.reset_index()# 重排索引

                merge_A_P = pd.merge(ProductMaster,AdjustRepPlan,how="left",on="Material")
                merge_A_P = merge_A_P[merge_A_P["FCST_state"]=="MTS"]
                merge_A_P.fillna(0,inplace=True)
                merge_A_P['adjust_Rep_QTY'] = merge_A_P['W1']+merge_A_P['W2']+                 merge_A_P['W3']+merge_A_P['W4']
                merge_A_P["Rep_value"] = merge_A_P["GTS"] * merge_A_P["adjust_Rep_QTY"]
                # 判断颜色，大于指标红色，小于指标绿色
                merge_Rep_value = sum(merge_A_P["Rep_value"])
                if Target_amount > merge_Rep_value:
                    lb_amount = Label(frame,text=re_round(merge_Rep_value),anchor="e",
                                      font=('黑体',15),width=15,height=2,fg='green',bg='WhiteSmoke')
                    lb_amount.place(x=520,y=46)
                else:
                    lb_amount = Label(frame,text=re_round(merge_Rep_value),anchor="e",
                                      font=('黑体',15),width=15,height=2,fg='red',bg='WhiteSmoke')
                    lb_amount.place(x=520,y=46)

                tkinter.messagebox.showinfo("提示","成功！")              
                plan.destroy()

            Button(plan,text="确认修改",font=("黑体",12,'bold'),bg='slategrey',fg='white',width=9,
                   height=1,borderwidth=5,command=db_update).place(x=900,y=150)

            plan.mainloop()
        
        treeview.bind('<Double-1>', modify_rolling_plan)
        
    read_rolling()# 第一次点击则自动运行
    
    
# 点击目录事件触发颜色变换和frame覆盖
# 二级目录按钮样式风格设置
def s_1():
    place_x = 0 
    place_y = 10
    font = ('华文中宋',14)
    bg='WhiteSmoke'
    fg='DimGray'
    borderwidth=0
    width=15
    height=2
    return place_x,place_y,font,bg,fg,borderwidth,width,height
class content:
    # 一级目录按钮-1及其相应事件与控件：数据更新
    def One_content():
        # 显示读取数据文件按钮、主数据维护按钮
        def One_content_1():
            frame_one_1 = Frame(window,height=655,width=1015,bg="WhiteSmoke")
            frame_one_1.place(x=267,y=61)
            lb_title = Label(window,text='更新当月数据 ',font=('华文中宋',14),bg='WhiteSmoke',
                             fg='black',width=10,height=2)
            lb_title.place(x=280,y=10)
            lb_title_f = Label(window,text="当前数据库最新至:"+JNJ_Month(1)[0],font=('黑体',12))
            lb_title_f.place(x=1000,y=25)
            btn_read_dir = Button(frame_one_1,text='选择指定文件夹上传',command=read_dir,
                                  font=('黑体',12,'bold'),bg='slategrey',fg='white',width=30,
                                  borderwidth=5,
                                  compound=CENTER)
            btn_read_dir.place(x=550,y=100)
            # 上传提示
            lb_tips=Label(window,text="仅支持上传一次，请按模板格式上传。\n"+             "如需更改数据，请联系管理员。",font=("黑体",12),fg='brown')
            lb_tips.place(in_=btn_read_dir,x=0,y=50)
            btn_create = Button(frame_one_1,text='下载数据更新模板',command=create_model,
                                font=('黑体',12,'bold'),bg='slategrey',fg='white',width=30,
                                borderwidth=5)
            btn_create.place(x=150,y=100)
        
        # 二级目录按钮
        frame = Frame(window, height=200, width=185,bg='Gainsboro')# bg=""背景色透明      
        btn_read = Button(frame,text='更新当月数据',command=One_content_1,font=s_1()[2],
                          fg=s_1()[4],width=s_1()[6],height=s_1()[7],borderwidth=s_1()[5],
                          anchor="center")
        btn_read.place(x=s_1()[0],y=s_1()[1])

        btn_master_maintain = Button(frame,text='主数据维护',command=master_maintain,
                                 font=s_1()[2],fg=s_1()[4],width=s_1()[6],height=s_1()[7],
                                 borderwidth=s_1()[5],compound=CENTER)
        btn_master_maintain.place(x=s_1()[0],y=s_1()[1]+80)
        frame.place(x=80,y=100)
        
        # 覆盖一级按钮，实现替换颜色
        btn_One_2 = Button(window,image=img_btn_update_png_2,borderwidth=0,height=45,width=45,
                           command=content.One_content)
        btn_One_2.place(x=17,y=100)
        btn_Two = Button(window,image=img_btn_FCST_png_1,borderwidth=0,height=45,width=45,
                         command=content.Two_content)
        btn_Two.place(x=17,y=180)
        btn_Three = Button(window,image=img_btn_Rep_png_1,borderwidth=0,height=45,width=45,
                           command=content.Three_content)
        btn_Three.place(x=17,y=260)
        btn_Four = Button(window,image=img_btn_Track_png_1,borderwidth=0,height=45,width=45,
                          command=content.Four_content)
        btn_Four.place(x=17,y=340)
        btn_Set = Button(window,image=img_btn_Set_png_1,borderwidth=0,height=45,width=45,
                         command=content.Five_content)
        btn_Set.place(x=17,y=640)
#         btn_More = Button(window,image=img_btn_More_png_1,borderwidth=0,height=42,width=42,
#                          command=content.Six_content)
#         btn_More.place(x=17,y=640)
        
        # 显示提示文本
        CreateToolTip(btn_One_2, "更新数据")
        CreateToolTip(btn_Two, "预测需求")
        CreateToolTip(btn_Three, "补货拆周")
        CreateToolTip(btn_Four, "订单追踪")
        CreateToolTip(btn_Set, "设置")
#         CreateToolTip(btn_More, "更多")
        
    # 一级目录按钮-2及其相应事件：需求预测
    def Two_content():
        def Two_content_1():
            # 标题
            lb_title_f = Label(window,text="当前数据库最新至:"+JNJ_Month(1)[0],font=('黑体',12))
            lb_title_f.place(x=1000,y=25)
            lb_title = Label(window,text='二级需求预测 ',font=('华文中宋',14),bg='WhiteSmoke',
                                  fg='black',width=10,height=2)
            lb_title.place(x=280,y=10)
            # 内容
            frame_1 = Frame(window,height=655,width=1015,bg='WhiteSmoke')
            frame_1.place(x=267,y=61)
#             btn_modify_history = Button(frame_1,text='修改历史数据',font=('黑体',12),bg='grey',
#                       fg='DimGray',width=18,height=2,borderwidth=2,compound=CENTER)
#             btn_modify_history.place(x=100,y=5)
#             btn_modify_model = Button(frame_1,text='修改模型数据',font=('黑体',12),bg='grey',
#                                   fg='DimGray',width=18,height=2,borderwidth=2,compound=CENTER)
#             btn_modify_model.place(x=300,y=5)
            btn_forecast = Button(frame_1,text='开始预测',command=forecast,font=('黑体',12,'bold'),
                                  bg='slategrey',fg='white',width=15,borderwidth=5,compound=CENTER)
            btn_forecast.place(x=650,y=15)
        
        frame = Frame(window, height=200, width=185,bg='Gainsboro')     
#         btn_clean = Button(frame,text='需求数据清洗',font=('黑体',12),bg='grey',
#                           fg='DimGray',width=18,height=2,borderwidth=0,compound=CENTER)
#         btn_clean.place(x=0,y=100)
        btn_Mape = Button(frame,text='Mape&Bias',command=MapeBias,
                                 font=s_1()[2],fg=s_1()[4],width=s_1()[6],height=s_1()[7],
                                 borderwidth=s_1()[5],compound=CENTER)
        btn_Mape.place(x=s_1()[0],y=s_1()[1])
        btn_FCSTDemand = Button(frame,text='二级需求预测',command=Two_content_1,
                                 font=s_1()[2],fg=s_1()[4],width=s_1()[6],height=s_1()[7],
                                 borderwidth=s_1()[5],compound=CENTER)
        btn_FCSTDemand.place(x=s_1()[0],y=s_1()[1]+80)

        frame.place(x=80,y=100)
        
        # 覆盖按钮，替换颜色
        btn_One = Button(window,image=img_btn_update_png_1,borderwidth=0,height=45,width=45,
                         command=content.One_content)
        btn_One.place(x=17,y=100)
        btn_Two_2 = Button(window,image=img_btn_FCST_png_2,borderwidth=0,height=45,width=45,
                           command=content.Two_content)
        btn_Two_2.place(x=17,y=180)
        btn_Three = Button(window,image=img_btn_Rep_png_1,borderwidth=0,height=45,width=45,
                           command=content.Three_content)
        btn_Three.place(x=17,y=260)
        btn_Four = Button(window,image=img_btn_Track_png_1,borderwidth=0,height=45,width=45,
                          command=content.Four_content)
        btn_Four.place(x=17,y=340)
        btn_Set = Button(window,image=img_btn_Set_png_1,borderwidth=0,height=45,width=45,
                         command=content.Five_content)
        btn_Set.place(x=17,y=640)
#         btn_More = Button(window,image=img_btn_More_png_1,borderwidth=0,height=42,width=42,
#                          command=content.Six_content)
#         btn_More.place(x=17,y=640)
                
        # 显示提示文本
        CreateToolTip(btn_One, "更新数据")
        CreateToolTip(btn_Two_2, "预测需求")
        CreateToolTip(btn_Three, "补货拆周")
        CreateToolTip(btn_Four, "订单追踪")
        CreateToolTip(btn_Set, "设置")
#         CreateToolTip(btn_More, "更多")
    
    # 一级目录按钮-3及其相应事件：补货计划
    def Three_content():
        # 二级目录按钮
        frame = Frame(window, height=200, width=185,bg='Gainsboro')# bg=""背景色透明
        btn_replenishment = Button(frame,text='补货计划',font=s_1()[2],command=Replenishment,
                                   fg=s_1()[4],width=s_1()[6],height=s_1()[7],borderwidth=s_1()[5],
                                   compound=CENTER)
        btn_replenishment.place(x=s_1()[0],y=s_1()[1])
        btn_modify_rep = Button(frame,text='手动修改',font=s_1()[2],command=modify_Replenishment,
                                fg=s_1()[4],width=s_1()[6],height=s_1()[7],borderwidth=s_1()[5],
                                compound=CENTER)
        btn_modify_rep.place(x=s_1()[0],y=s_1()[1]+80)
        
        frame.place(x=80,y=100)
        
        # 覆盖按钮，替换颜色
        btn_One = Button(window,image=img_btn_update_png_1,borderwidth=0,height=45,width=45,
                         command=content.One_content)
        btn_One.place(x=17,y=100)
        btn_Two = Button(window,image=img_btn_FCST_png_1,borderwidth=0,height=45,width=45,
                         command=content.Two_content)
        btn_Two.place(x=17,y=180)
        btn_Three_2 = Button(window,image=img_btn_Rep_png_2,borderwidth=0,height=45,width=45,
                             command=content.Three_content)
        btn_Three_2.place(x=17,y=260)
        btn_Four = Button(window,image=img_btn_Track_png_1,borderwidth=0,height=45,width=45,
                          command=content.Four_content)
        btn_Four.place(x=17,y=340)
        btn_Set = Button(window,image=img_btn_Set_png_1,borderwidth=0,height=45,width=45,
                         command=content.Five_content)
        btn_Set.place(x=17,y=640)
#         btn_More = Button(window,image=img_btn_More_png_1,borderwidth=0,height=42,width=42,
#                          command=content.Six_content)
#         btn_More.place(x=17,y=640)
        
        # 显示提示文本
        CreateToolTip(btn_One, "更新数据")
        CreateToolTip(btn_Two, "预测需求")
        CreateToolTip(btn_Three_2, "补货拆周")
        CreateToolTip(btn_Four, "订单追踪")
        CreateToolTip(btn_Set, "设置")
#         CreateToolTip(btn_More, "更多")
    
    # 一级目录按钮-4及其相应事件：进出追踪
    def Four_content():
        # 二级目录
        frame = Frame(window, height=200, width=185,bg='Gainsboro')
        btn_clean = Button(frame,text='进出追踪',font=s_1()[2],command=Access_tracking,
                           fg=s_1()[4],width=s_1()[6],height=s_1()[7],borderwidth=s_1()[5],
                           compound=CENTER)
        btn_clean.place(x=s_1()[0],y=s_1()[1])
        btn_rolling_rep = Button(frame,text='补货更新',font=s_1()[2],command=rolling_rep,
                                fg=s_1()[4],width=s_1()[6],height=s_1()[7],borderwidth=s_1()[5],
                                compound=CENTER)
        btn_rolling_rep.place(x=s_1()[0],y=s_1()[1]+80)
        frame.place(x=80,y=100)
        
        # 覆盖按钮，替换颜色
        btn_One = Button(window,image=img_btn_update_png_1,borderwidth=0,height=45,width=45,
                         command=content.One_content)
        btn_One.place(x=17,y=100)
        btn_Two = Button(window,image=img_btn_FCST_png_1,borderwidth=0,height=45,width=45,
                         command=content.Two_content)
        btn_Two.place(x=17,y=180)
        btn_Three = Button(window,image=img_btn_Rep_png_1,borderwidth=0,height=45,width=45,
                           command=content.Three_content)
        btn_Three.place(x=17,y=260)
        btn_Four_2 = Button(window,image=img_btn_Track_png_2,borderwidth=0,height=45,width=45,
                            command=content.Four_content)
        btn_Four_2.place(x=17,y=340)
        btn_Set = Button(window,image=img_btn_Set_png_1,borderwidth=0,height=45,width=45,
                         command=content.Five_content)
        btn_Set.place(x=17,y=640)
#         btn_More = Button(window,image=img_btn_More_png_1,borderwidth=0,height=42,width=42,
#                          command=content.Six_content)
#         btn_More.place(x=17,y=640)
        
        # 显示提示文本
        CreateToolTip(btn_One, "更新数据")
        CreateToolTip(btn_Two, "预测需求")
        CreateToolTip(btn_Three, "补货拆周")
        CreateToolTip(btn_Four_2, "订单追踪")
        CreateToolTip(btn_Set, "设置")
#         CreateToolTip(btn_More, "更多")
        
    # 一级目录按钮-5及其相应事件：设置
    def Five_content():
        frame = Frame(window,height=655,width=1015,bg='WhiteSmoke')
        frame.place(x=267,y=61)
        lb_continue = Label(frame,text="当前版本：Prism V2.1",font=('黑体',15),bg='WhiteSmoke',
                       fg=s_1()[4],width=20,height=2)
        lb_continue.place(x=300,y=300)
        
        # 标题
        lb_title = Label(window,text='设置',font=('华文中宋',14),bg='WhiteSmoke',
                         fg='black',width=10,height=2)
        lb_title.place(x=280,y=10)
        
        # 二级目录
        frame_2 = Frame(window, height=200, width=185,bg='Gainsboro')     
        frame_2.place(x=80,y=100)
        
        # 覆盖按钮，替换颜色
        btn_One = Button(window,image=img_btn_update_png_1,borderwidth=0,height=45,width=45,
                         command=content.One_content)
        btn_One.place(x=17,y=100)
        btn_Two = Button(window,image=img_btn_FCST_png_1,borderwidth=0,height=45,width=45,
                         command=content.Two_content)
        btn_Two.place(x=17,y=180)
        btn_Three = Button(window,image=img_btn_Rep_png_1,borderwidth=0,height=45,width=45,
                           command=content.Three_content)
        btn_Three.place(x=17,y=260)
        btn_Four = Button(window,image=img_btn_Track_png_1,borderwidth=0,height=45,width=45,
                          command=content.Four_content)
        btn_Four.place(x=17,y=340)
        btn_Set_2 = Button(window,image=img_btn_Set_png_2,borderwidth=0,height=45,width=45,
                           command=content.Five_content)
        btn_Set_2.place(x=17,y=640)
#         btn_More = Button(window,image=img_btn_More_png_1,borderwidth=0,height=42,width=42,
#                          command=content.Six_content)
#         btn_More.place(x=17,y=640)
        
        # 显示提示文本
        CreateToolTip(btn_One, "更新数据")
        CreateToolTip(btn_Two, "预测需求")
        CreateToolTip(btn_Three, "补货拆周")
        CreateToolTip(btn_Four, "订单追踪")
        CreateToolTip(btn_Set_2, "设置")
#         CreateToolTip(btn_More, "更多")


# 一级目录按钮显示
btn_One = Button(window,image=img_btn_update_png_1,borderwidth=0,height=45,width=45,
                 command=content.One_content)
btn_One.place(x=17,y=100)

btn_Two = Button(window,image=img_btn_FCST_png_1,borderwidth=0,height=45,width=45,
                 command=content.Two_content)
btn_Two.place(x=17,y=180)

btn_Three = Button(window,image=img_btn_Rep_png_1,borderwidth=0,height=45,width=45,
                 command=content.Three_content)
btn_Three.place(x=17,y=260)

btn_Four = Button(window,image=img_btn_Track_png_1,borderwidth=0,height=45,width=45,
                  command=content.Four_content)
btn_Four.place(x=17,y=340)

btn_Set = Button(window,image=img_btn_Set_png_1,borderwidth=0,height=45,width=45,
                 command=content.Five_content)
btn_Set.place(x=17,y=640)

# btn_More = Button(window,image=img_btn_More_png_1,borderwidth=0,height=42,width=42,
#                  command=content.Six_content)
# btn_More.place(x=17,y=640)

# 鼠标移动至位置显示
CreateToolTip(btn_One, "更新数据")
CreateToolTip(btn_Two, "预测需求")
CreateToolTip(btn_Three, "补货拆周")
CreateToolTip(btn_Four, "订单追踪")
CreateToolTip(btn_Set, "设置")
# CreateToolTip(btn_More, "更多")

# 退出按钮
btn_quit = Button(window,text="×",command=window.destroy,bg='WhiteSmoke',
                  borderwidth=0,font=('黑体',18))
btn_quit.place(x=1240,y=5)
CreateToolTip(btn_quit, "关闭")

window.mainloop()

