import tkinter
import pandas as pd
import sqlite3
from sqlalchemy import create_engine


# 链接数据库 python 对sql数据库操作进行封装
# -- coding: utf-8 --
# Prism 数据库操作，输入sql语句，返回查询df、修改数据、删除数据等
class PrismDatabaseOperation:

    def __init__(self) -> None:
        self.connect_to_database()

    # def __enter__(self):
    #     self.connect_to_database()

    def __exit__(self):
        self.cursor.close()

    def connect_to_database(self):
        try:
            # 连接数据库
            self.conn = sqlite3.connect('prism_data.db')
            self.cursor = self.conn.cursor()
        except:
            print("连接数据库失败！")
            # tkinter.messagebox.showerror("错误", "连接数据库失败！")

    # 查询
    def Prism_select(self, sql_cmd) -> pd.DataFrame:
        try:
            df_data = pd.read_sql(con=self.conn, sql=sql_cmd, index_col=None)
            return df_data
        except:
            print("数据查询失败！")
            # tkinter.messagebox.showerror("错误", "数据查询失败！")

    # 修改
    def Prism_update(self, sql_cmd):
        try:
            self.cursor.execute(sql_cmd)
            self.cursor.commit()
            # self.conn.close()
        except:
            print("数据更新失败！")
            # tkinter.messagebox.showerror("错误", "数据更新失败！")

    # 删除
    def Prism_delete(self, sql_cmd):
        try:
            self.cursor.execute(sql_cmd)
            self.cursor.commit()
            # self.conn.close()
        except:
            print("数据删除失败！")
            # tkinter.messagebox.showerror("错误", "数据删除失败！")

    # 增加,特殊方法，直接附加df,如有null，替换为0
    def Prism_insert(self, db_sheetname: str, df_input: pd.DataFrame) -> bool:
        try:
            # value = str(tuple("?" * len(df_input.columns)))  # 占位符
            # # 将需插入的值转换为[()]形式
            # df_value = df_input.apply(lambda x: tuple(x), axis=1).values.tolist()
            # sql_insert = "insert into " + db_sheetname + str(tuple(list(df_input.columns))) + " values " + value + ";"
            # # print(sql_insert)
            # # 执行语句,executemany效率较高
            # self.cursor.executemany(sql_insert, df_value)
            # self.cursor.commit()
            # # df_input.to_sql(db_sheetname,con=self.conn, if_exists='append', index=False)
            # self.conn.close()
            # return True
            df_input.to_sql(db_sheetname,
                            con=self.conn,
                            if_exists='append',
                            index=False)
            # 替换sqlite链接能使用的额语句
            self.conn.commit()
            # self.conn.close()
            return True
        except:
            print(db_sheetname + "数据插入失败！")
            # self.conn.close()
            # tkinter.messagebox.showerror("错误", db_sheetname + "数据插入失败！")
            return False


if __name__ == '__main__':
    module_test = PrismDatabaseOperation()
    print(module_test.Prism_select("select Material from ProductMaster"))
