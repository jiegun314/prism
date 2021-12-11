from prism_database_operation import PrismDatabaseOperation
import pandas as pd
from tkinter import Tk, messagebox, Text
from tkinter.ttk import *


# Master表信息维护
class MasterData:
    def __init__(self) -> None:
        pass
    
    # 查询code数据
    # Jeffrey - 无用
    def master_search(self) -> pd.DataFrame:
        sql_cmd = "SELECT Material FROM ProductMaster"
        df_material = self.prism_db_operation.Prism_select(sql_cmd).values
        return df_material
    
    # 主数据校验,校验是否存在，若否，则返回相应code
    # Jeffrey - 无用
    def master_check(self, df_input) -> pd.Series:
        master_code = pd.DataFrame(columns=['Material'], data=self.master_search())
        lack_code = df_input[~df_input['Material'].isin(master_code['Material'])]['Material']
        if lack_code.empty:
            return 0
        else:
            return lack_code

    # 合并输入数据
    def set_value(self, update_type, modify_treeview, modify_master, event) -> None: 
        # 获取鼠标所选item
        for item in modify_treeview.selection():
            item_text = modify_treeview.item(item, "values")

        column = modify_treeview.identify_column(event.x)# 所在列
        cn = int(str(column).replace('#',''))
        
        if update_type == 'single':
            Label_select = Label(modify_master,text=str(item_text[cn-1]),width=20,anchor="w")
            Label_select.place(x=150, y=100)

            entryedit = Entry(modify_master,width=13)
            entryedit.insert(0,str(item_text[cn-1]))
            entryedit.place(x=150, y=150)
        else:
            Label_select = Label(modify_master,text=str(item_text[cn-1]),width=20)
            Label_select.place(x=150, y=300)
            
            entryedit = Text(modify_master,width=15,height = 1)
            entryedit.place(x=150, y=350)
        
        treeedit_value = entryedit.get() if update_type == 'single' else entryedit.get(0.0, "end")[:-1]
        axis_y = 150 if update_type == 'single' else 350
        
        # 将编辑好的信息更新到数据库中
        def save_edit():
            # 获取
            modify_treeview.set(item, column=column, value=treeedit_value)
            entryedit.destroy()
            btn_input.destroy()
            btn_cancal.destroy()
            Label_select.destroy()
            
        btn_input = Button(modify_master, text='OK', width=7, command=save_edit)
        btn_input.place(x=260,y=axis_y)
        
        # 取消输入
        def cancal_edit():
            entryedit.destroy()
            btn_input.destroy()
            btn_cancal.destroy()
            Label_select.destroy()
        
        btn_cancal = Button(modify_master, text='Cancel', width=7, command=cancal_edit)
        btn_cancal.place(x=350,y=axis_y)

    def db_update(self, modify_treeview, column_input, modify_master, type):
        # 获取所有最新数据,直接更新所有数据
        # 先删除，再直接附加~更简单~
        # 遍历获取所有数据，并生成df
        try:
            list_value = [modify_treeview.item(item,'values') for item in modify_treeview.get_children()]
            if type=='single':
                column_now=column_input
            else:
                with PrismDatabaseOperation() as db_op:
                    column_now=db_op.Prism_select('SELECT TOP 1 * FROM ProductMaster').columns
            df_now = pd.DataFrame(list_value, columns=column_now)
            if type=='single':
                df_now.rename(columns={"规格型号":"Material", 
                                       "不含税单价":"GTS",
                                       "预测状态":"FCST_state"},
                              inplace=True)
            for i in ["GTS","包装规格","MOQ","安全库存天数"]:
                df_now[i] = df_now[i].astype(float)
            # 删除已有material
            for i in range(len(df_now)):
                with PrismDatabaseOperation() as db_op:
                    sql_cmd = "DELETE FROM ProductMaster WHERE Material =\'%s\'" % df_now['Material'].iloc[i]
                    db_op.Prism_delete(sql_cmd)
            with PrismDatabaseOperation() as db_op:
                if db_op.Prism_insert('ProductMaster',df_now):
                    messagebox.showinfo("提示","成功！")
        except:
            messagebox.showerror("错误","修改失败！请检查数据格式")
        master_maintain()
        modify_master.destroy()

    # 修改数据（数据校验修改前后是否存在重复）
    def master_update(self, item_text):
        modify_master = Tk()
        modify_master.title('主数据修改')
        modify_master.geometry('1050x250')
        columns = ['规格型号','包装规格','分类Level3','分类Level4','ABC','分类','不含税单价','预测状态','MOQ','安全库存天数']
        modify_treeview = Treeview(modify_master, height=1, show="headings", columns=columns)
        modify_treeview.place(x=20,y=20)

        for i in range(len(columns)):
            modify_treeview.column(columns[i], width=100, anchor='center')

        # 显示列名
        for i in range(len(columns)):
            modify_treeview.heading(columns[i], text=columns[i])
        modify_treeview.insert('', 1, values=item_text)
        
        # 合并输入数据
        self.set_value(update_type='single', modify_treeview=modify_treeview, modify_master=modify_master)

        # 触发双击事件
        modify_treeview.bind('<Double-1>', self.set_value)

        # 显示文本数据
        Label(modify_master,
              text="Tips：包装规格、不含税单价、MOQ、安全库存天数必须为数字！",
              fg='red').place(x=100,y=200)
        Label(modify_master,text="修改前：").place(x=100,y=100)
        Label(modify_master,text="修改后：").place(x=100,y=150)

        Button(modify_master,
               text="确认修改", 
               font=("黑体",12,'bold'), 
               bg='slategrey', 
               fg='white',
               width=9, 
               height=1, 
               orderwidth=5, 
               command=self.db_update(modify_treeview=modify_treeview, 
                                      column_input=columns, 
                                      modify_master=modify_master, 
                                      type='single')).place(x=920,y=200)

        modify_master.mainloop()

    # 批量修改数据
    def master_update_batch(self, df):
        # 数据修改窗口
        modify_master = Tk()
        modify_master.title('主数据修改')
        # 窗体大小随df变化而变化
        modify_master.geometry('1050x500')
        columns = list(df.columns)
        modify_treeview = Treeview(modify_master, height=10, show="headings", 
                                       columns=columns)
        modify_treeview.place(x=20,y=20)
        sb = Scrollbar(modify_master,command=modify_treeview.yview)
        sb.config(command=modify_treeview.yview)
        sb.place(x=975,y=0,in_ = modify_treeview,height=230)
        modify_treeview.config(yscrollcommand=sb.set)
        
        #  表示列,不显示,文本靠左，数字靠右
        for i in columns:
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
                                    command=lambda _col=col: treeview_sort_column(modify_treeview, _col, False))
        
        # 合并输入数据
        self.set_value(update_type='batch', modify_treeview=modify_treeview, modify_master=modify_master)

        # 触发双击事件
        modify_treeview.bind('<Double-1>', self.set_value)

        # 显示文本数据
        Label(modify_master,text="修改前：").place(x=100,y=300)
        Label(modify_master,text="修改后：").place(x=100,y=350)
        
        Button(modify_master, 
               text="确认修改", 
               font=("黑体",12,'bold'), 
               bg='slategrey', 
               fg='white',
               width=9, 
               height=1, 
               borderwidth=5, 
               command=self.db_update(modify_treeview=modify_treeview, 
                                      column_input='', 
                                      modify_master=modify_master, 
                                      type='batch')).place(x=920,y=400)
        
        # 提示文本
        Label(modify_master,text="Tips：包装规格、不含税单价、MOQ、安全库存天数必须为数字！",
              fg='red').place(x=100,y=400)

        modify_master.mainloop()