import tkinter
import pandas as pd
import pyodbc
from sqlalchemy import create_engine

# 链接数据库 python 对sql数据库操作进行封装
#-- coding: utf-8 --
# Prism 数据库操作，输入sql语句，返回查询df、修改数据、删除数据等
class PrismDatabaseOperation:
    """
    执行Prism数据库sql语句的函数，可进行增、删、改、查
    注：增加数据时不直接通过读取excel文件输入，而是先read处理后再insert
    """
    def __init__(self) -> None:
        # define database link parameter
        self.dict_db_connection = {
            'server': 'localhost',
            'user': 'ryan',
            'password': 'prism123+',
            'database': 'Prism',
            'charset': 'cp936'
        }
    
    def __enter__(self):
        self.connect_to_database()
    
    def __exit__(self):
        self.cursor.close()
    
    def connect_to_database(self):
        try:
            # 连接数据库
            self.conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server}; SERVER= %s; DATABASE= %s; UID=%s; PWD=%s' % (
                self.dict_db_connection['server'], self.dict_db_connection['database'], self.dict_db_connection['user'], self.dict_db_connection['password']))
            self.cursor = self.conn.cursor()
        except:
            tkinter.messagebox.showerror("错误","连接数据库失败！")
        pass
    
    # 查询
    def Prism_select(self, sql_cmd) -> pd.DataFrame:
        try:
            self.cursor.execute(sql_cmd) # 执行语句
            data = self.cursor.fetchall()
            columnDes = self.cursor.description #获取连接对象的描述信息
            columnNames = [columnDes[i][0] for i in range(len(columnDes))]
            df = pd.DataFrame([list(i) for i in data],columns=columnNames)
            return df
        except:
            tkinter.messagebox.showerror("错误","数据查询失败！")

    # 修改
    def Prism_update(self, sql_cmd):
        try:
            self.cursor.execute(sql_cmd)
            self.cursor.commit()
        except:
            tkinter.messagebox.showerror("错误","数据更新失败！")

    # 删除
    def Prism_delete(self, sql_cmd):
        try:
            self.cursor.execute(sql_cmd)
            self.cursor.commit()
        except:
            tkinter.messagebox.showerror("错误","数据删除失败！")

    # 增加,特殊方法，直接附加df,如有null，替换为0
    def Prism_insert(self, db_sheetname: str, df_input: pd.DataFrame) -> bool:
        try:
            df_input.to_sql(db_sheetname, 
                            con=self.conn, 
                            if_exists='append', 
                            index=False)
            return True
        except:
            tkinter.messagebox.showerror("错误",db_sheetname+"数据插入失败！")
            return False