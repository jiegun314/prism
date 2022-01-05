# from tkinter import Label, Frame, ttk, filedialog, Button, Entry, messagebox
# from matplotlib.pyplot import axis
import pandas as pd
import numpy as np
from prism_database_operation import PrismDatabaseOperation
import datetime
from dateutil.relativedelta import relativedelta
import math


# import decimal
#
# # 默认截取，更改为四舍五入
# decimal.getcontext().rounding = "ROUND_HALF_UP"


class PrismCalculation:
    def __init__(self) -> None:
        # 数据库操作
        self.prime_db_ops = PrismDatabaseOperation()
        # 缺失数据提醒
        self._lst_missing = []
        # 返回数据
        self._dictionary = {}
        # # 操作用户名记录
        # self._user_name = ""
        # # 上传时间
        # self._jnj_time = ""
        pass

    # 获取日期与强生年月对应关系表,判断是否为关账日
    def get_close_date(self, _today: str) -> bool:
        # 获取对照表数据
        sql_select = "SELECT * FROM JNJCalendar"
        df_jnj_calendar = self.prime_db_ops.Prism_select(sql_select)
        # df_close_date
        # 转换为日期格式，并与数据源进行匹配
        pass

    # jeffrey - 没必要
    def new_round(self, _float, _len):
        if isinstance(_float, float):
            if str(_float)[::-1].find('.') <= _len:
                return _float
            if str(_float)[-1] == '5':
                return round(float(str(_float)[:-1] + '6'), _len)
            else:
                return round(_float, _len)
        else:
            return round(_float, _len)

    # 四舍五入替换为decimal
    # def decimal_round(self, _float, n):
    #     if n < 0:
    #         n = 0
    #         print("默认替换为0")
    #     _len = "0." + "0" * n
    #     result = decimal.Decimal(str(_float)).quantize(decimal.Decimal(_len))
    #     return result

    # Jeffrey: 在函数中规定输入输出类型是一个好习惯
    def get_jnj_month(self, n: int) -> list:
        # Jeffrey: 所有能够在数据库里面完成的运算不要拿到外面来算,包括去重，排序等
        sql_select = "SELECT DISTINCT(JNJ_Date) AS jnj_month From Outbound ORDER BY jnj_month"
        df_month_output = self.prime_db_ops.Prism_select(sql_select)
        # Jeffrey: 多利用pandas的标准函数，包括转换list等
        lst_month = df_month_output['jnj_month'].values.tolist()
        # Outbound_month = list(self.prime_db_ops.Prism_select(sql_select)['JNJ_Date'].unique())
        # last_month = sorted(Outbound_month)[-n:]
        return lst_month[-n:]

    def get_next_jnj_month(self) -> str:
        next_month = (datetime.datetime.strptime(self.get_jnj_month(1)[0], "%Y%m") +
                      relativedelta(months=1)).strftime("%Y%m")  # 下个月时间
        return next_month

    # 获取指定月份的季节因子
    def get_season_factor(self, input_month: str) -> float:
        sql_cmd = 'SELECT season_factor FROM SeasonFactor WHERE JNJ_Date=\"%s\"' % input_month
        df_season_factor = self.prime_db_ops.Prism_select(sql_cmd)
        lst_season_factor = df_season_factor.values.tolist()
        if len(lst_season_factor[0]) == 0:
            return 1
        else:
            return lst_season_factor[0][0]

    # 读取周补货权值 -字典形式直接读取所有
    def get_weekly_pattern(self) -> list:
        sql_cmd = "SELECT * FROM WeeklyPattern"
        df_weekly_pattern = self.prime_db_ops.Prism_select(sql_cmd)
        # 获取补货权值
        WK1 = df_weekly_pattern[df_weekly_pattern['week'] == 'WK1']['pattern'].iloc[0]
        WK2 = df_weekly_pattern[df_weekly_pattern['week'] == 'WK2']['pattern'].iloc[0]
        WK3 = df_weekly_pattern[df_weekly_pattern['week'] == 'WK3']['pattern'].iloc[0]
        WK4 = df_weekly_pattern[df_weekly_pattern['week'] == 'WK4']['pattern'].iloc[0]
        return [WK1, WK2, WK3, WK4]

    # # 命名对应数据表（？）
    # def df_rename(self) -> dict:
    #     dict_rename = {"Intransit": "在途",
    #                    "Intransit_QTY": "在途量"}
    #     return dict_rename

    # 主数据批量维护插入数据功能
    def master_update_batch(self, df_new_material: pd.DataFrame) -> bool:
        # 数据预处理
        df_new_material.rename(columns={"规格型号": "Material", "不含税单价": "GTS", "预测状态": "FCST_state"},
                               inplace=True)
        # 查询主数据
        sql_select = "SELECT * FROM ProductMaster;"
        ProductMaster = self.prime_db_ops.Prism_select(sql_select)
        # 判断主数据本身是否为空
        try:
            if ProductMaster.empty:
                self.prime_db_ops.Prism_insert("ProductMaster", df_new_material)
                print("提示，主数据为空！")
            else:
                # 判断是否重复，如果重复，则以最新的为准；即删除已存在的code，再把之前的一起添加入主数据
                df_exist_material = ProductMaster[ProductMaster["Material"].isin(df_new_material["Material"].tolist())]
                if not df_exist_material.empty:
                    for i in range(len(df_exist_material)):
                        SQL_delete = "DELETE FROM ProductMaster WHERE Material ='%s'" % \
                                     df_exist_material['Material'].iloc[i]
                        self.prime_db_ops.Prism_delete(SQL_delete)
                # 写入数据库
                self.prime_db_ops.Prism_insert("ProductMaster", df_new_material)
            return True
        except:
            return False

    # 根据表格内容、字段查询，返回查询数据
    def search_material(self, df_search: pd.DataFrame, search_field: str, search_content: str = "") -> pd.DataFrame:
        for i in df_search.columns:
            df_search[i] = df_search[i].apply(str)  # 必须转字符，否则无法全局搜索

        # 查找并插入数据
        if search_content != "":
            # 全局搜索
            if search_field == "全局搜索":
                search_df = pd.DataFrame(columns=df_search.columns)
                for i in range(len(df_search.columns)):
                    search_df = search_df.append(df_search[df_search[df_search.columns[i]].str.contains(
                        search_content)])
                search_df.drop_duplicates(subset=["规格型号"], keep='first', inplace=True)
            # 指定字段搜索
            else:
                appoint = search_field
                search_df = df_search[df_search[appoint].str.contains(search_content)]
            # 插入表格
            return search_df
        # 若输入值为空则显示全部内容
        else:
            return df_search

    # 根据时间判断数据是否已存在于数据库
    def jnj_date_exist(self, table_name: str, new_date: str) -> bool:
        sql_select = "SELECT JNJ_Date FROM " + table_name + " GROUP BY JNJ_Date;"
        # print(sql_select)
        JNJ_Date_exist = self.prime_db_ops.Prism_select(sql_select)
        if JNJ_Date_exist.empty:
            return True
        else:
            if new_date in JNJ_Date_exist["JNJ_Date"].to_list():
                return True
            else:
                return False

    # 检查主数据是否有新增code
    def check_product_master(self, df_merge: pd.DataFrame) -> pd.DataFrame:
        # 获取当前Product Master
        ProductMaster = self.prime_db_ops.Prism_select('SELECT * FROM ProductMaster')
        # 合并
        df_merge_lack = pd.merge(ProductMaster, df_merge, on='Material', how='outer')
        # to_open = df_merge_lack[df_merge_lack['预测状态']=='MTS'][df_merge_lack['FCST_state']=='关']
        # to_add = df_merge_lack[df_merge_lack['预测状态'] == 'MTS'][df_merge_lack['FCST_state'].isnull()]
        # to_update = to_open.append(to_add)
        to_add = df_merge_lack[df_merge_lack['FCST_state'].isnull()]
        # 只截取主数据所需数据
        to_add = to_add[list(ProductMaster.columns)]
        to_add.rename(columns={'Material': '规格型号', 'GTS': '不含税单价', 'FCST_state': '预测状态'}, inplace=True)

        # 检测是code是否缺失
        if not to_add.empty:
            return to_add
        else:
            return pd.DataFrame()

    # 导入基础数据文件--在途  *中文文件名报错
    def input_intransit(self, file_name: str, df: pd.DataFrame) -> dict:
        # 检查输入数据的类型是否正确
        # 匹配强生年月
        jnj_date = file_name[file_name.find("_") + 1:file_name.find("_") + 7]
        # 导入销售库存信息
        if '在途' in file_name:
            try:
                # Intransit = pd.read_excel(file_path, dtype={'规格型号': str})
                df['JNJ_Date'] = jnj_date
                # 统一命名
                df.rename(columns={'规格型号': 'Material', '数量': 'Intransit_QTY'}, inplace=True)
                df_intransit = df[['Material', 'Intransit_QTY', 'JNJ_Date']]
                # 预处理：删除Material为空的空行
                df_intransit.drop(df_intransit[df_intransit["Material"].isnull()].index, inplace=True)
                # 导入之前判断数据库的强生年月（文件名中含有）中是否已经存在，若有，则删除再覆盖；若无，则直接添加
                if self.jnj_date_exist("Intransit", jnj_date):
                    SQL_delete = "DELETE FROM Intransit WHERE JNJ_Date = '%s'" % jnj_date
                    self.prime_db_ops.Prism_delete(SQL_delete)
                # 导入数据库
                input_state = self.prime_db_ops.Prism_insert('Intransit', df_intransit)
                if input_state:
                    print(file_name + "导入成功")
                else:
                    # 返回异常报错信息 risk error
                    print(file_name + "数据导入失败！")
            except:
                print(file_name + "数据导入失败！")
        else:
            print("请检查文件名是否正确！")

        # 对导入文件做一次检查,若有新增code，则返回新增的code， 并中断数据上传，提醒用户再次上传数据
        self._dictionary["new_code"] = self.check_product_master(df_intransit)
        self._dictionary["input_state"] = input_state

        return self._dictionary

    # 导入基础数据文件--出库
    def input_outbound(self, file_name: str, df: pd.DataFrame) -> dict:
        # 检查输入数据的类型是否正确
        # 匹配强生年月
        jnj_date = file_name[file_name.find("_") + 1:file_name.find("_") + 7]
        # 导入销售库存信息
        if '出库' in file_name:
            try:
                df['JNJ_Date'] = jnj_date
                # 统一命名
                df.rename(columns={'规格型号': 'Material', '数量': 'Outbound_QTY'}, inplace=True)
                df_outbound = df[['Material', 'Outbound_QTY', 'JNJ_Date']]
                # 删除Material为空的空行
                df_outbound.drop(df_outbound[df_outbound["Material"].isnull()].index, inplace=True)
                # 导入之前判断数据库的强生年月（文件名中含有）中是否已经存在，若有，则删除再覆盖；若无，则直接添加
                # 导入之前判断数据库的强生年月（文件名中含有）中是否已经存在，若有，则删除再覆盖；若无，则直接添加
                if self.jnj_date_exist("Outbound", jnj_date):
                    SQL_delete = "DELETE FROM Outbound WHERE JNJ_Date = '%s'" % jnj_date
                    self.prime_db_ops.Prism_delete(SQL_delete)
                # 导入数据库
                input_state = self.prime_db_ops.Prism_insert('Outbound', df_outbound)
                if input_state:
                    print(file_name + "导入成功")
                else:
                    print(file_name + "数据导入失败！")
            except:
                print(file_name + "数据导入失败！")
        else:
            print("请检查文件名是否正确！")

        # 对导入文件做一次检查,若有新增code，则返回新增的code， 并中断数据上传，提醒用户再次上传数据
        self._dictionary["new_code"] = self.check_product_master(df_outbound)
        self._dictionary["input_state"] = input_state

        return self._dictionary

    # 导入基础数据文件--可发
    def input_onhand(self, file_name: str, df: pd.DataFrame) -> dict:
        # 检查输入数据的类型是否正确
        # 匹配强生年月
        jnj_date = file_name[file_name.find("_") + 1:file_name.find("_") + 7]
        # 导入销售库存信息
        if '可发' in file_name:
            try:
                df['JNJ_Date'] = jnj_date
                # 统一命名
                df.rename(columns={'规格型号': 'Material', '数量': 'Onhand_QTY'}, inplace=True)
                df_onhand = df[['Material', 'Onhand_QTY', 'JNJ_Date']]
                # 删除Material为空的空行
                df_onhand.drop(df_onhand[df_onhand["Material"].isnull()].index, inplace=True)
                # 导入之前判断数据库的强生年月（文件名中含有）中是否已经存在，若有，则删除再覆盖；若无，则直接添加
                if self.jnj_date_exist("Onhand", jnj_date):
                    SQL_delete = "DELETE FROM Onhand WHERE JNJ_Date = '%s'" % jnj_date
                    self.prime_db_ops.Prism_delete(SQL_delete)
                # 导入数据库
                input_state = self.prime_db_ops.Prism_insert('Onhand', df_onhand)
                if input_state:
                    print(file_name + "导入成功")
                else:
                    print(file_name + "数据导入失败！")
            except:
                print(file_name + "数据导入失败！")
        else:
            print("请检查文件名是否正确！")

        # 对导入文件做一次检查,若有新增code，则返回新增的code， 并中断数据上传，提醒用户再次上传数据
        self._dictionary["new_code"] = self.check_product_master(df_onhand)
        self._dictionary["input_state"] = input_state

        return self._dictionary

    # 导入基础数据文件--缺货
    def input_backorder(self, file_name: str, df: pd.DataFrame) -> dict:
        # 检查输入数据的类型是否正确
        # 匹配强生年月
        jnj_date = file_name[file_name.find("_") + 1:file_name.find("_") + 7]
        # 导入销售库存信息
        if '缺货' in file_name:
            try:
                df['JNJ_Date'] = jnj_date
                # 统一命名
                df.rename(columns={'规格型号': 'Material', '数量': 'Backorder_QTY'}, inplace=True)
                df_backorder = df[['Material', 'Backorder_QTY', 'JNJ_Date']]
                # 删除Material为空的空行
                df_backorder.drop(df_backorder[df_backorder["Material"].isnull()].index, inplace=True)
                # 导入之前判断数据库的强生年月（文件名中含有）中是否已经存在，若有，则删除再覆盖；若无，则直接添加
                if self.jnj_date_exist("Backorder", jnj_date):
                    SQL_delete = "DELETE FROM Backorder WHERE JNJ_Date = '%s'" % jnj_date
                    self.prime_db_ops.Prism_delete(SQL_delete)
                # 导入数据库
                input_state = self.prime_db_ops.Prism_insert('Backorder', df_backorder)
                if input_state:
                    print(file_name + "导入成功")
                else:
                    print(file_name + "数据导入失败！")
            except:
                print(file_name + "数据导入失败！")
        else:
            print("请检查文件名是否正确！")

        # 对导入文件做一次检查,若有新增code，则返回新增的code， 并中断数据上传，提醒用户再次上传数据
        self._dictionary["new_code"] = self.check_product_master(df_backorder)
        self._dictionary["input_state"] = input_state

        return self._dictionary

    # 导入基础数据文件--预入库
    def input_putaway(self, file_name: str, df: pd.DataFrame) -> dict:
        # 检查输入数据的类型是否正确
        # 匹配强生年月
        jnj_date = file_name[file_name.find("_") + 1:file_name.find("_") + 7]
        # 导入销售库存信息
        if '预入库' in file_name:
            try:
                df['JNJ_Date'] = jnj_date
                # 统一命名
                df.rename(columns={'规格型号': 'Material', '数量': 'Putaway_QTY'}, inplace=True)
                df_putaway = df[['Material', 'Putaway_QTY', 'JNJ_Date']]
                # 删除Material为空的空行
                df_putaway.drop(df_putaway[df_putaway["Material"].isnull()].index, inplace=True)
                # 导入之前判断数据库的强生年月（文件名中含有）中是否已经存在，若有，则删除再覆盖；若无，则直接添加
                if self.jnj_date_exist("Putaway", jnj_date):
                    SQL_delete = "DELETE FROM Putaway WHERE JNJ_Date = '%s'" % jnj_date
                    self.prime_db_ops.Prism_delete(SQL_delete)
                # 导入数据库
                input_state = self.prime_db_ops.Prism_insert('Putaway', df_putaway)
                if input_state:
                    print(file_name + "导入成功")
                else:
                    print(file_name + "数据导入失败！")
            except:
                print(file_name + "数据导入失败！")
        else:
            print("请检查文件名是否正确！")

        # 对导入文件做一次检查,若有新增code，则返回新增的code， 并中断数据上传，提醒用户再次上传数据
        self._dictionary["new_code"] = self.check_product_master(df_putaway)
        self._dictionary["input_state"] = input_state

        return self._dictionary

    # 读取出库库存
    def get_outbound_record(self, month_qty: int) -> pd.DataFrame:
        lst_month = self.get_jnj_month(month_qty)
        str_month_list = '(\"' + '\",\"'.join(lst_month) + '\")'
        sql_cmd = "SELECT * FROM Outbound WHERE JNJ_Date in %s" % str_month_list
        Outbound_QTY = self.prime_db_ops.Prism_select(sql_cmd)
        Outbound_QTY = Outbound_QTY.pivot_table(index='Material', columns="JNJ_Date", values='Outbound_QTY')
        Outbound_QTY = Outbound_QTY.reset_index()
        return Outbound_QTY

    # # 获取主数据*
    # def get_product_master(self, condition: str) -> pd.DataFrame:
    #     # 主数据
    #     SQL_ProductMaster = "SELECT * From ProductMaster %s;" % condition
    #     ProductMaster = self.prime_db_ops.Prism_select(SQL_ProductMaster)
    #     # ProductMaster.rename(columns={'ABC': "Class"}, inplace=True)
    #     return ProductMaster

    # # 表格合并
    # def df_merge(self, df_input:pd.DataFrame) -> pd.DataFrame:
    #     # 获取数据

    # # 读取缺货库存
    # def get_backorder_record(self, month_qty: int) -> pd.DataFrame:
    #     lst_month = self.get_jnj_month(month_qty)
    #     str_month_list = '(\"' + '\",\"'.join(lst_month) + '\")'
    #     sql_cmd = "SELECT * FROM Backorder WHERE JNJ_Date in %s" % str_month_list
    #     Backorder_QTY = self.prime_db_ops.Prism_select(sql_cmd)
    #     Backorder_QTY = Backorder_QTY.pivot_table(index='Material', columns="JNJ_Date", values='Backorder_QTY')
    #     Backorder_QTY = Backorder_QTY.reset_index()
    #     return Backorder_QTY

    # Jeffrey 独立模块拆成独立函数 - 更新需求数据操作
    def update_actual_demand(self) -> None:
        last_month = self.get_jnj_month(1)[0]
        # Jeffrey: 在变量命名中最好能表明它的数据类型，尤其对于dataframe这种特殊数据类型
        # Jeffrey：使用标准的字符串拼接
        df_outbound = self.prime_db_ops.Prism_select("SELECT * FROM Outbound WHERE JNJ_Date=\"%s\"" % last_month)
        df_backorder = self.prime_db_ops.Prism_select("SELECT * FROM Backorder WHERE JNJ_Date=\"%s\"" % last_month)
        df_backorder_ttl = pd.merge(df_outbound, df_backorder, how="outer", on=['Material', 'JNJ_Date'])
        df_season_factor = self.prime_db_ops.Prism_select("SELECT * FROM SeasonFactor;")
        df_act_demand = pd.merge(df_backorder_ttl, df_season_factor, how="left", on=['JNJ_Date'])
        # Jeffrey: 尽可能缩减重复语句
        df_act_demand.fillna({'Backorder_QTY': 0, 'Outbound_QTY': 0}, inplace=True)
        # df_act_demand['Backorder_QTY'].fillna(0,inplace=True)
        # df_act_demand['Outbound_QTY'].fillna(0,inplace=True)
        df_act_demand['ActDemand_QTY'] = (df_act_demand['Outbound_QTY'] + df_act_demand['Backorder_QTY']
                                          ) / df_act_demand['season_factor']
        # Jeffrey: 同名覆盖会引起歧义，尽量避免
        df_act_demand.drop(columns=['Outbound_QTY', 'Backorder_QTY', 'season_factor'], inplace=True)
        # df_act_demand = df_act_demand[['JNJ_Date','Material','ActDemand_QTY']]
        # 插入之前判断是否已有数据
        df_current_demand = self.prime_db_ops.Prism_select("SELECT * FROM ActDemand WHERE JNJ_Date=\"%s\"" % last_month)
        if df_current_demand.empty:
            self.prime_db_ops.Prism_insert('ActDemand', df_act_demand)
        else:
            df_act_demand = df_current_demand.copy()
            self._lst_missing.append("模型数据")
        pass

    # Jeffrey - 将计算forecast模块单列，将来可复用
    def generate_forecast(self, df_input: pd.DataFrame, lst_weight: list, fcst_mth: int) -> pd.DataFrame:
        # get following month list
        lst_month_name = []
        for i in range(fcst_mth):
            following_mth = datetime.datetime.strptime(df_input.columns[-1], '%Y%m') + relativedelta(months=+ (i + 1))
            lst_month_name.append(following_mth.strftime('%Y%m'))
        lst_cycle_month = [3, 6, 12]
        for i in range(fcst_mth):
            for j in range(len(lst_cycle_month)):
                df_input['sum_%s' % lst_cycle_month[j]] = (df_input.iloc[:, 0 - lst_cycle_month[j]:].mean(axis=1)) * \
                                                          lst_weight[i]
            df_input[lst_month_name[i]] = df_input.iloc[:, (0 - len(lst_cycle_month)):].sum(
                axis=1) * self.get_season_factor(lst_month_name[i])
            # remove sum (-4:-2)
            df_input.drop(columns=df_input.columns[(-1 - len(lst_cycle_month)):-1], axis=1, inplace=True)
        return df_input

    def forecast_generation(self) -> pd.DataFrame:
        self.update_actual_demand()

        # Jeffrey: 多次查询数据库会显著降低性能，采用合理的sql语句一次查询完成
        lst_12_months = self.get_jnj_month(12)
        str_month_list = '(\"' + '\",\"'.join(lst_12_months) + '\")'
        sql_cmd = 'SELECT * FROM ActDemand WHERE JNJ_Date in %s ORDER BY JNJ_Date' % str_month_list
        df_fcst_model = self.prime_db_ops.Prism_select(sql_cmd)
        # df_fcst_model = pd.DataFrame()
        # for i in range(12):
        #     SQL_select = "SELECT * FROM ActDemand WHERE JNJ_Date='"+lst_12_months[i]+"';"
        #     df_fcst_model = df_fcst_model.append(self.prime_db_ops.Prism_select(SQL_select))

        # 获取所有FCST_state为开的ProductMaster数据
        SQL_state = "SELECT Material,分类Level4 FROM ProductMaster WHERE FCST_state = 'MTS'"
        df_mst = self.prime_db_ops.Prism_select(SQL_state)

        # 链接数据，得到所需的计算12个月的数据
        df_acl_demand = pd.merge(df_mst, df_fcst_model, how="outer", on=['Material'])
        df_acl_demand.fillna(0, inplace=True)

        # 数透，material为行，JNJ_Date为列
        df_demand_pivot = df_acl_demand.pivot_table(index='Material', columns='JNJ_Date', values='ActDemand_QTY')
        df_demand_pivot.fillna(0, inplace=True)

        # Jeffrey - 制定value创建透视表，不需要以下代码
        # 换列名，方便直接取出相应月份的数据
        # acl_FCSTDemand_12_col = []
        # for i in range(len(df_demand_pivot.columns)):
        #     if type(df_demand_pivot.columns.values[i][1]) == str :
        #         acl_FCSTDemand_12_col.append(df_demand_pivot.columns.values[i][1])
        #     else:
        #         acl_FCSTDemand_12_col.append(str(df_demand_pivot.columns.values[i][1]))
        # #         print(acl_FCSTDemand_12_col)
        # df_demand_pivot.columns = acl_FCSTDemand_12_col
        try:
            del df_demand_pivot["0"]  # 当master里有而需求中没有就会出现不必要的0列，删除即可
        except:
            pass
        # 获取权值
        df_fcst_weight = self.prime_db_ops.Prism_select("SELECT * FROM FCSTWeight")
        # Jeffrey - 能精简则精简
        lst_weight = df_fcst_weight['值'].values.tolist()
        # [w1, w2, w3] = df_fcst_weight['值'].values.tolist()
        # w1 = FCSTWeight['值'].iloc[0]
        # w2 = FCSTWeight['值'].iloc[1]
        # w3 = FCSTWeight['值'].iloc[2]
        # 循环计算3次，以预测值预测后三月内容，nice！
        # Jeffrey - 预测主函数单列
        df_forecast_result = self.generate_forecast(df_demand_pivot, lst_weight, fcst_mth=3)

        # for i in range(3):
        #     # 获取相应月数据，并计算相应结果列
        #     model_3 = df_demand_pivot[df_demand_pivot.columns[-3:]]
        #     acl_model_3 = model_3[list(model_3.columns)[0]] # 求和3月值
        #     for i in range(1,len(model_3.columns)):
        #         acl_model_3 = acl_model_3 + model_3[list(model_3.columns)[i]]
        #     acl_model_3 = acl_model_3/3*w1

        #     model_6 = df_demand_pivot[df_demand_pivot.columns[-6:]]
        #     acl_model_6 = model_6[list(model_6.columns)[0]] # 求和6月值
        #     for i in range(1,len(model_6.columns)):
        #         acl_model_6 = acl_model_6 + model_6[list(model_6.columns)[i]]
        #     acl_model_6 = acl_model_6/6*w2

        #     model_12 = df_demand_pivot[df_demand_pivot.columns[-12:]]
        #     acl_model_12 = model_12[list(model_12.columns)[0]] # 求和12月值
        #     for i in range(1,len(model_12.columns)):
        #         acl_model_12 = acl_model_12 + model_12[list(model_12.columns)[i]]
        #     acl_model_12 = acl_model_12/12*w3
        #     FCSTModel_QTY = acl_model_3+acl_model_6+acl_model_12

        #     # 获取计算月的后一个月
        #     month_after_1 = datetime.datetime.strptime(df_demand_pivot.columns[-1],
        #                                                '%Y%m')+relativedelta(months=+1)
        #     month_after_1 = month_after_1.strftime('%Y%m')
        #     # 将计算出来的最新一个月数据赋值给12月的计算数据
        #     df_demand_pivot[month_after_1] = FCSTModel_QTY

        # 取出预测的3个值，并排好列序，以防插入数据库时失败
        df_fcst_demand = df_forecast_result.iloc[:, -3:].reset_index()  # 将数透格式转为df
        df_fcst_demand.columns = ['Material', 'FCST_Demand1', 'FCST_Demand2', 'FCST_Demand3']
        # jeffrey - 直接将月份列放在第一列，以免以后调整
        df_fcst_demand.insert(0, 'JNJ_Date', len(df_fcst_demand) * self.get_jnj_month(1))

        # Jeffrey - 季节因子的引入直接采用子函数，而且在计算中直接导入

        # # 前面除过的季节因子，现在乘回来，需获取当月的JNJ_Date并获取后三个月的季节因子
        # try:
        #     df_season_factor = self.prime_db_ops.Prism_select("SELECT * FROM SeasonFactor")
        #     months_1 = (datetime.datetime.strptime(self.get_jnj_month(1)[0],"%Y%m")+
        #                 relativedelta(months=1)).strftime("%Y%m")
        #     months_2 = (datetime.datetime.strptime(self.get_jnj_month(1)[0],"%Y%m")+
        #                 relativedelta(months=2)).strftime("%Y%m")
        #     months_3 = (datetime.datetime.strptime(self.get_jnj_month(1)[0],"%Y%m")+
        #                 relativedelta(months=3)).strftime("%Y%m")
        #     SeasonFactor_1 = df_season_factor[df_season_factor['JNJ_Date']==months_1
        #                                  ]['season_factor'].iloc[0]
        #     SeasonFactor_2 = df_season_factor[df_season_factor['JNJ_Date']==months_2
        #                                  ]['season_factor'].iloc[0]
        #     SeasonFactor_3 = df_season_factor[df_season_factor['JNJ_Date']==months_3
        #                                  ]['season_factor'].iloc[0]
        # except:
        #     # messagebox.showerror("错误","季节因子不存在，请维护季节因子数据")
        #     print('季节因子不存在，请维护季节因子数据')

        # df_fcst_demand['FCST_Demand1'] = df_fcst_demand['FCST_Demand1']*SeasonFactor_1
        # df_fcst_demand['FCST_Demand2'] = df_fcst_demand['FCST_Demand2']*SeasonFactor_2
        # df_fcst_demand['FCST_Demand3'] = df_fcst_demand['FCST_Demand3']*SeasonFactor_3

        # FCSTDemand = FCST_Demand[['JNJ_Date','Material','FCST_Demand1','FCST_Demand2',
        #                           'FCST_Demand3']]
        df_fcst_demand.fillna(0, inplace=True)
        # 全部取整
        # df_fcst_demand['FCST_Demand1'] = self.new_round(df_fcst_demand['FCST_Demand1'],0)
        # df_fcst_demand['FCST_Demand2'] = self.new_round(df_fcst_demand['FCST_Demand2'],0)
        # df_fcst_demand['FCST_Demand3'] = self.new_round(df_fcst_demand['FCST_Demand3'],0)

        # Jeffrey - 没必要做精确取整，为确保demand，可以考虑上取整，否则所有小于0.5的都会被清零
        df_fcst_demand[['FCST_Demand1', 'FCST_Demand2', 'FCST_Demand3']] = df_fcst_demand[
            ['FCST_Demand1', 'FCST_Demand2', 'FCST_Demand3']].applymap(np.int64)

        # 判断数据库中是否已存在，将计算出来的数值插入到数据库中
        # Jeffrey - 已经存在不覆盖？？逻辑需要调整，如果存在就不覆盖，应该先计算是不是已经有了，可以直接不计算
        SQL_select = "SELECT DISTINCT(JNJ_Date) AS jnj_month FROM FCSTDemand"
        FCST_Demand_Date = self.prime_db_ops.Prism_select(SQL_select)['jnj_month'].values.tolist()
        if self.get_jnj_month(1) not in FCST_Demand_Date:
            self.prime_db_ops.Prism_insert('FCSTDemand', df_fcst_demand)
        else:
            SQL = 'SELECT * FROM FCSTDemand WHERE JNJ_Date =\"$s\"' % self.get_jnj_month(1)[0]
            df_fcst_demand = self.prime_db_ops.Prism_select(SQL)
            self._lst_missing.append("预测数据")
            # tkinter.messagebox.showinfo("提示",JNJ_Month(1)[0]+"预测数据已存在！")

        # 断数据库中是否已存在，将FCST_Demand1存入预测需求调整表格中
        AdjustFCSTDemand = df_fcst_demand[['JNJ_Date', 'Material', 'FCST_Demand1']]
        AdjustFCSTDemand.loc[:, "Remark"] = ""
        SQL_select = "SELECT DISTINCT(JNJ_Date) AS jnj_month FROM AdjustFCSTDemand"
        Adjust_FCST_Demand_Date = self.prime_db_ops.Prism_select(SQL_select)['jnj_month'].values.tolist()
        if self.get_jnj_month(1) not in Adjust_FCST_Demand_Date:
            self.prime_db_ops.Prism_insert('AdjustFCSTDemand', AdjustFCSTDemand)
        else:
            SQL = 'SELECT * FROM AdjustFCSTDemand WHERE JNJ_Date =\"$s\"' % self.get_jnj_month(1)[0]
            AdjustFCSTDemand = self.prime_db_ops.Prism_select(SQL)
            self._lst_missing.append("已调整预测需求")

            # 如果missing已存在，则提示哪些已存在
        if self._lst_missing:
            # messagebox.showinfo("提示",str(missing)+"已存在!")
            print('%s已存在' % str(self._lst_missing))

        # 将预测结果显示到主窗口，规格型号、产品家族、出库记录（6个月）、置信度（6个月数据，千分位）
        SQL_select = "SELECT [Material],[分类Level4] FROM ProductMaster WHERE FCST_state = 'MTS'"
        ProductMaster = self.prime_db_ops.Prism_select(SQL_select)

        # 获取出库记录（6个月）
        # Jeffrey - 独立功能，单独成函数
        Outbound_QTY = self.get_outbound_record(month_qty=6)

        # 合并并计算置信度
        # Jeffrey - 避免无意义命名
        df_merge_temp_1 = pd.merge(ProductMaster, Outbound_QTY, how='left', on="Material")
        df_merge_temp_2 = pd.merge(df_merge_temp_1, df_fcst_demand, how='left', on="Material")
        df_fcst_final = pd.merge(df_merge_temp_2, AdjustFCSTDemand, how='left', on="Material")
        # Jeffrey - 删除列尽量用标准函数
        df_fcst_final.drop(columns=['JNJ_Date_x', 'JNJ_Date_y'], inplace=True)
        df_fcst_final['miu'] = np.mean(df_fcst_final[self.get_jnj_month(6)].iloc[:], axis=1)
        df_fcst_final['sigma'] = np.std(df_fcst_final[self.get_jnj_month(6)].iloc[:], axis=1)
        df_fcst_final.rename(columns={'Material': '规格型号', 'FCST_Demand1_x': '模型预测值',
                                      'FCST_Demand1_y': '最终预测值', 'Remark': '修改原因'},
                             inplace=True)
        # # 历史数据为0，置信度应该为高
        # # 根据出库记录，计算置信度，按可信程度区分，切片替换,|x-μ|的距离判断即可
        # Jeffrey - 用简单的逻辑计算sigma level
        df_fcst_final.loc[:, 'sigma_level'] = abs(df_fcst_final['最终预测值'] - df_fcst_final['miu']) / df_fcst_final[
            'sigma']
        df_fcst_final.loc[:, '置信度'] = ""
        df_fcst_final.loc[df_fcst_final['sigma_level'] > 3, '置信度'] = '低'
        df_fcst_final.loc[(df_fcst_final['sigma_level'] > 2) & (df_fcst_final['sigma_level'] <= 3), '置信度'] = '较低'
        df_fcst_final.loc[(df_fcst_final['sigma_level'] > 1) & (df_fcst_final['sigma_level'] <= 2), '置信度'] = '较高'
        df_fcst_final.loc[df_fcst_final['sigma_level'] <= 1, '置信度'] = '高'
        df_fcst_final.fillna(0, inplace=True)
        # Jeffrey 尽可能利用完整数组
        month_list = self.get_jnj_month(6)
        month_list.reverse()
        # 重新截取数据并排序
        df_fcst_final = df_fcst_final[["规格型号", "分类Level4", "模型预测值", "最终预测值", "修改原因", "置信度"] + month_list]

        return df_fcst_final

    # 计算mape和bias
    def mape_bias(self) -> dict:
        last_month = self.get_jnj_month(1)[0]
        # 出库数据
        SQL_Outbound = "SELECT * FROM Outbound WHERE JNJ_Date='%s';" % last_month
        Outbound = self.prime_db_ops.Prism_select(SQL_Outbound)
        Outbound = Outbound[["Material", "Outbound_QTY"]]
        if Outbound.empty:
            self._lst_missing.append("出库")

        # 调整后的需求数据
        SQL_AdjustFCSTDemand = "SELECT * FROM AdjustFCSTDemand WHERE JNJ_Date='%s'" % self.get_jnj_month(2)[0]
        AdjustFCSTDemand = self.prime_db_ops.Prism_select(SQL_AdjustFCSTDemand)
        AdjustFCSTDemand = AdjustFCSTDemand[["Material", "FCST_Demand1", "Remark"]]
        if AdjustFCSTDemand.empty:
            self._lst_missing.append("调整后的需求数据")

        # 缺货数据
        SQL_Backorder = "SELECT * FROM Backorder WHERE JNJ_Date='%s'" % last_month
        Backorder = self.prime_db_ops.Prism_select(SQL_Backorder)
        Backorder = Backorder[["Material", "Backorder_QTY"]]
        if Backorder.empty:
            self._lst_missing.append("缺货")

        # 主数据
        SQL_ProductMaster = "SELECT Material,ABC,FCST_state From ProductMaster;"
        ProductMaster = self.prime_db_ops.Prism_select(SQL_ProductMaster)

        # 合并
        df_merge_temp_1 = pd.merge(AdjustFCSTDemand, Outbound, on="Material", how="outer")
        df_merge_temp_2 = pd.merge(df_merge_temp_1, Backorder, on="Material", how="outer")
        merge_all = pd.merge(df_merge_temp_2, ProductMaster, on="Material", how="outer")
        merge_all = merge_all[merge_all["FCST_state"] == "MTS"]
        merge_all.fillna(0, inplace=True)

        # 计算Gap、Mape、Bias
        merge_all["Gap"] = merge_all["Outbound_QTY"] - merge_all["FCST_Demand1"]
        total_Mape = sum(abs(merge_all["Gap"])) / sum(merge_all["Outbound_QTY"])
        total_Bias = sum(merge_all["Gap"]) / sum(merge_all["Outbound_QTY"])
        merge_all["Mape"] = abs(merge_all["Gap"]) / merge_all["Outbound_QTY"]

        # mape剔除Outbound_QTY=0的情况
        mape_df = merge_all[merge_all["Outbound_QTY"] != 0].sort_values(by=["Mape"], ascending=False)

        # 替换空值重命名、排序、截取所需信息、切换格式等
        merge_all.fillna(0, inplace=True)
        merge_all.rename(columns={"Material": "规格型号", "Outbound_QTY": "实际出库",
                                  "FCST_Demand1": "需求预测值", "Remark": "调整原因",
                                  "Backorder_QTY": "缺货"}, inplace=True)
        # drop替换
        merge_all = merge_all[["规格型号", "ABC", "实际出库", "需求预测值", "Gap", "Mape", "调整原因"]]
        merge_all.loc[np.isinf(merge_all["Mape"]), "Mape"] = "无实际出库"
        self._dictionary['mape_df'] = mape_df
        self._dictionary['merge_all'] = mape_df
        self._dictionary['total_Mape'] = total_Mape
        self._dictionary['total_Bias'] = total_Bias

        return self._dictionary

    # 计算历史数据
    def history_data(self) -> dict:
        # 计算历史数据
        JNJ_12Month = str(tuple(self.get_jnj_month(12)))

        # 主数据
        SQL_ProductMaster = "SELECT Material,ABC,FCST_state From ProductMaster;"
        ProductMaster = self.prime_db_ops.Prism_select(SQL_ProductMaster)

        # 出库数据
        SQL_Outbound = "SELECT * FROM Outbound WHERE JNJ_Date in %s ;" % JNJ_12Month
        Outbound = self.prime_db_ops.Prism_select(SQL_Outbound)
        if Outbound.empty:
            self._lst_missing.append("出库")

        # 缺货数据
        SQL_Backorder = "SELECT * FROM Backorder WHERE JNJ_Date in %s ;" % JNJ_12Month
        Backorder = self.prime_db_ops.Prism_select(SQL_Backorder)
        if Backorder.empty:
            self._lst_missing.append("缺货")

        # 调整后的需求数据
        str_month_list = '(\"' + '\",\"'.join(self.get_jnj_month(13)[0:12]) + '\")'
        SQL_AdjustFCSTDemand = "SELECT * FROM AdjustFCSTDemand WHERE JNJ_Date in %s ;" % str_month_list
        AdjustFCSTDemand = self.prime_db_ops.Prism_select(SQL_AdjustFCSTDemand)
        # 为使相应数据合并，所有月份加1
        for i in range(len(AdjustFCSTDemand)):
            AdjustFCSTDemand.loc[i, "JNJ_Date"] = (datetime.datetime.strptime(
                AdjustFCSTDemand.loc[i, "JNJ_Date"], "%Y%m") + relativedelta(months=1)).strftime("%Y%m")

        if AdjustFCSTDemand.empty:
            self._lst_missing.append("调整后的需求数据")

        # 合并
        df_merge_temp_1 = pd.merge(Outbound, Backorder, how="outer", on=["Material", "JNJ_Date"])
        df_merge_temp_2 = pd.merge(df_merge_temp_1, AdjustFCSTDemand, how="outer", on=["Material", "JNJ_Date"])
        merge_all = pd.merge(df_merge_temp_2, ProductMaster, how="outer", on="Material")
        merge_all = merge_all[merge_all["FCST_state"] == "MTS"]

        merge_all["Outbound_QTY"].fillna(0, inplace=True)
        merge_all["Backorder_QTY"].fillna(0, inplace=True)
        merge_all["FCST_Demand1"].fillna(0, inplace=True)
        merge_all["Remark"].fillna("", inplace=True)
        merge_all["total"] = merge_all["Outbound_QTY"] + merge_all["Backorder_QTY"]

        # 聚合
        total = merge_all.groupby(merge_all["JNJ_Date"]).sum()
        total.reset_index(inplace=True)

        # Remark界面显示
        # 抓取有Remark的,前端显示
        Remark = merge_all[merge_all["Remark"] != ""]

        # 重命名及排序
        Remark.rename(columns={"Material": "规格型号", "Outbound_QTY": "实际出库",
                               "Backorder_QTY": "缺货", "FCST_Demand1": "预测需求",
                               "JNJ_Date": "年月"},
                      inplace=True)
        Remark = Remark[["规格型号", "ABC", "年月", "Remark", "实际出库", "缺货", "预测需求"]]

        # 返回字典型数据
        self._dictionary["total"] = total
        self._dictionary["Remark"] = Remark
        self._dictionary["merge_all"] = merge_all

        return self._dictionary

    # 读取补货计划所需基础数据
    def read_replishment(self) -> pd.DataFrame:
        # 主数据
        SQL_ProductMaster = "SELECT Material,GTS,ABC,MOQ From ProductMaster WHERE FCST_state='MTS';"
        ProductMaster = self.prime_db_ops.Prism_select(SQL_ProductMaster)
        ProductMaster.rename(columns={'ABC': "Class"}, inplace=True)

        # 出库数据6个月，计算置信度
        jnj_month_6 = str(tuple(self.get_jnj_month(6)))
        SQL_Outbound = "SELECT * From Outbound WHERE JNJ_Date in %s " % jnj_month_6
        Outbound = self.prime_db_ops.Prism_select(SQL_Outbound)
        Outbound = Outbound.pivot_table(index='Material', columns='JNJ_Date')
        Outbound_QTY_col = []  # 换列名，方便直接取出相应月份的数据
        for i in range(6):
            Outbound_QTY_col.append(Outbound.columns.values[i][1])
        Outbound.columns = Outbound_QTY_col
        Outbound = Outbound.reset_index()
        if Outbound.empty:
            self._lst_missing.append("出库")

        last_month = self.get_jnj_month(3)[2]

        # Intransit
        SQL_Intransit = "SELECT Material,Intransit_QTY From Intransit WHERE JNJ_Date = '%s'" % last_month
        Intransit = self.prime_db_ops.Prism_select(SQL_Intransit)
        if Intransit.empty:
            self._lst_missing.append("在途")

        # Onhand_QTY
        SQL_Onhand = "SELECT Material,Onhand_QTY From Onhand WHERE JNJ_Date = '%s'" % last_month
        Onhand = self.prime_db_ops.Prism_select(SQL_Onhand)
        if Onhand.empty:
            self._lst_missing.append("可发")

        # Putaway
        SQL_Putaway = "SELECT Material,Putaway_QTY From Putaway WHERE JNJ_Date = '%s'" % last_month
        Putaway = self.prime_db_ops.Prism_select(SQL_Putaway)
        if Putaway.empty:
            self._lst_missing.append("预入库")

        # Backorder
        SQL_Backorder = "SELECT Material,Backorder_QTY From Backorder WHERE JNJ_Date = '%s'" % last_month
        Backorder = self.prime_db_ops.Prism_select(SQL_Backorder)
        if Backorder.empty:
            self._lst_missing.append("缺货")

        # AdjustFCSTDemand
        SQL_Adjust = "SELECT Material,FCST_Demand1,Remark From AdjustFCSTDemand WHERE JNJ_Date = '%s'" % last_month
        AdjustFCSTDemand = self.prime_db_ops.Prism_select(SQL_Adjust)
        AdjustFCSTDemand['FCST_Demand1'] = self.new_round(AdjustFCSTDemand['FCST_Demand1'], 0)
        if AdjustFCSTDemand.empty:
            self._lst_missing.append("需求")

        # SafetyStockDay安全库存天数
        SQL_SafetyStockDay = "SELECT [Class],[Safetystock_Day] From SafetyStockDay;"
        SafetyStockDay = self.prime_db_ops.Prism_select(SQL_SafetyStockDay)
        if SafetyStockDay.empty:
            self._lst_missing.append("安全库存天数")

        if self._lst_missing:
            print(self._lst_missing)
        #     tkinter.messagebox.showwarning("警告",str(lack)+"数据缺失!")

        # 合并所需信息
        df_merge_tmp1 = pd.merge(AdjustFCSTDemand, Intransit, how="outer", on="Material")
        df_merge_tmp2 = pd.merge(df_merge_tmp1, Onhand, how="outer", on="Material")
        df_merge_tmp3 = pd.merge(df_merge_tmp2, Putaway, how="outer", on="Material")
        df_merge_tmp4 = pd.merge(df_merge_tmp3, Backorder, how="outer", on="Material")
        df_merge_tmp5 = pd.merge(df_merge_tmp4, Outbound, how="outer", on="Material")
        df_merge_tmp6 = pd.merge(ProductMaster, df_merge_tmp5, how="left", on="Material")
        merge_all = pd.merge(df_merge_tmp6, SafetyStockDay, how="left", on="Class")
        merge_all.drop_duplicates(keep='first', inplace=True)  # 删除因merge产生的意外重复code
        merge_all.fillna(0, inplace=True)

        return merge_all

    # 计算补货拆周逻辑
    def acl_excrete_week(self, merge_all: pd.DataFrame) -> pd.DataFrame:
        # 获取补货权值
        WK1, WK2, WK3, WK4 = self.get_weekly_pattern()
        merge_all['W1'] = 0
        merge_all['W2'] = 0
        merge_all['W3'] = 0
        merge_all['W4'] = 0

        # 计算拆周
        for i in range(len(merge_all)):
            # 第一、二、三周
            if merge_all['Rep_QTY'].iloc[i] <= 300:
                if merge_all['TotalINV_QTY'].iloc[i] <= merge_all['FCST_Demand1'].iloc[i] * (WK1 + WK2):
                    merge_all['W1'].iloc[i] = merge_all['Rep_QTY'].iloc[i]
                elif (merge_all['FCST_Demand1'].iloc[i] * (WK1 + WK2) < merge_all['TotalINV_QTY'].iloc[i] <
                      merge_all['FCST_Demand1'].iloc[i]):
                    merge_all['W2'].iloc[i] = merge_all['Rep_QTY'].iloc[i]
                else:
                    merge_all['W3'].iloc[i] = merge_all['Rep_QTY'].iloc[i]
            elif merge_all['Rep_QTY'].iloc[i] < 900:
                if merge_all['TotalINV_QTY'].iloc[i] <= merge_all['FCST_Demand1'].iloc[i] * (WK1 + WK2):
                    merge_all['W1'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * (WK1 + WK2) /
                                                             merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
                    merge_all['W3'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * (1 - WK1 - WK2) /
                                                             merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
                else:
                    merge_all['W2'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * (WK1 + WK2) /
                                                             merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
                    merge_all['W3'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * (1 - WK1 - WK2) /
                                                             merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
            else:
                merge_all['W1'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * WK1 /
                                                         merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
                merge_all['W2'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * WK2 /
                                                         merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
                merge_all['W3'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * WK3 /
                                                         merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]

                # 第四周
            merge_all['W4'].iloc[i] = (merge_all['Rep_QTY'].iloc[i] - merge_all['W1'].iloc[i] -
                                       merge_all['W2'].iloc[i] - merge_all['W3'].iloc[i])
            if merge_all['W4'].iloc[i] < 0:
                merge_all['W4'].iloc[i] = 0
                merge_all['W3'].iloc[i] = merge_all['W3'].iloc[i] - merge_all['W4'].iloc[i]

            return merge_all

    # 计算补货拆周计划
    def acl_replishment(self) -> dict:
        merge_all = self.read_replishment()
        # 计算安全库存量(四舍五入取整)
        merge_all['Safetystock_QTY'] = 0
        month_1 = self.get_jnj_month(3)[0]
        month_2 = self.get_jnj_month(3)[1]
        month_3 = self.get_jnj_month(3)[2]
        for i in range(len(merge_all)):
            merge_all['Safetystock_QTY'].iloc[i] = self.new_round(merge_all['Safetystock_Day'].iloc[i] *
                                                                  (merge_all[month_1].iloc[i] +
                                                                   merge_all[month_2].iloc[i] +
                                                                   merge_all[month_3].iloc[i]) / 90, 0)
        # 计算置信度
        merge_all['miu'] = np.mean(merge_all[self.get_jnj_month(6)].iloc[:], axis=1)
        merge_all['sigma'] = np.std(merge_all[self.get_jnj_month(6)].iloc[:], axis=1)

        # 历史数据为0，置信度应该为高
        # 根据出库记录，计算置信度，按可信程度区分，切片替换,|x-μ|的距离判断即可
        merge_all['sigma_level'] = abs(merge_all['FCST_Demand1'] - merge_all['miu']) / merge_all['sigma']
        merge_all['置信度'] = ""
        merge_all.loc[merge_all['sigma_level'] > 3, '置信度'] = '低'
        merge_all.loc[(merge_all['sigma_level'] > 2) & (merge_all['sigma_level'] <= 3), '置信度'] = '较低'
        merge_all.loc[(merge_all['sigma_level'] > 1) & (merge_all['sigma_level'] <= 2), '置信度'] = '较高'
        merge_all.loc[merge_all['sigma_level'] <= 1, '置信度'] = '高'

        # 计算完成，删除不需要显示的列
        merge_all = merge_all.drop(self.get_jnj_month(6)[0:2] + ['sigma_level'], axis=1)

        # 计算总库存
        merge_all['TotalINV_QTY'] = merge_all['Intransit_QTY'] + merge_all['Onhand_QTY'] + merge_all['Putaway_QTY']

        # 计算初始补货量
        merge_all['RepV1_QTY'] = merge_all['FCST_Demand1'] + merge_all['Safetystock_QTY'] + merge_all[
            'Backorder_QTY'] - merge_all['TotalINV_QTY']

        # 补货量凑整（二期：关于整托凑整）
        merge_all['Rep_QTY'] = 0
        for i in range(len(merge_all)):
            if merge_all['RepV1_QTY'].iloc[i] >= 300:
                merge_all['Rep_QTY'].iloc[i] = self.new_round(merge_all['RepV1_QTY'].iloc[i] / merge_all['MOQ'].iloc[i],
                                                              0) * merge_all['MOQ'].iloc[i]
            elif merge_all['RepV1_QTY'].iloc[i] >= 50:
                merge_all['Rep_QTY'].iloc[i] = self.new_round(merge_all['RepV1_QTY'].iloc[i] / merge_all['MOQ'].iloc[i],
                                                              0) * merge_all['MOQ'].iloc[i]
            elif merge_all['RepV1_QTY'].iloc[i] >= 2:
                merge_all['Rep_QTY'].iloc[i] = math.ceil(merge_all['RepV1_QTY'].iloc[i] / merge_all['MOQ'].iloc[i]
                                                         ) * merge_all['MOQ'].iloc[i]
            elif merge_all['RepV1_QTY'].iloc[i] >= 1:
                merge_all['Rep_QTY'].iloc[i] = 1
            else:
                merge_all['Rep_QTY'].iloc[i] = 0

        # 计算补货金额
        merge_all['Rep_value'] = merge_all['Rep_QTY'] * merge_all['GTS']

        # 计算拆周
        merge_all = self.acl_excrete_week(merge_all)
        # # 获取补货权值
        # WK1, WK2, WK3, WK4 = self.get_weekly_pattern()
        # merge_all['W1'] = 0
        # merge_all['W2'] = 0
        # merge_all['W3'] = 0
        # merge_all['W4'] = 0
        #
        # # 计算拆周
        # for i in range(len(merge_all)):
        #     # 第一、二、三周
        #     if merge_all['Rep_QTY'].iloc[i] <= 300:
        #         if merge_all['TotalINV_QTY'].iloc[i] <= merge_all['FCST_Demand1'].iloc[i] * (WK1 + WK2):
        #             merge_all['W1'].iloc[i] = merge_all['Rep_QTY'].iloc[i]
        #         elif (merge_all['FCST_Demand1'].iloc[i] * (WK1 + WK2) < merge_all['TotalINV_QTY'].iloc[i] <
        #               merge_all['FCST_Demand1'].iloc[i]):
        #             merge_all['W2'].iloc[i] = merge_all['Rep_QTY'].iloc[i]
        #         else:
        #             merge_all['W3'].iloc[i] = merge_all['Rep_QTY'].iloc[i]
        #     elif merge_all['Rep_QTY'].iloc[i] < 900:
        #         if merge_all['TotalINV_QTY'].iloc[i] <= merge_all['FCST_Demand1'].iloc[i] * (WK1 + WK2):
        #             merge_all['W1'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * (WK1 + WK2) /
        #                                                      merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
        #             merge_all['W3'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * (1 - WK1 - WK2) /
        #                                                      merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
        #         else:
        #             merge_all['W2'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * (WK1 + WK2) /
        #                                                      merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
        #             merge_all['W3'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * (1 - WK1 - WK2) /
        #                                                      merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
        #     else:
        #         merge_all['W1'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * WK1 /
        #                                                  merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
        #         merge_all['W2'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * WK2 /
        #                                                  merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
        #         merge_all['W3'].iloc[i] = self.new_round(merge_all['Rep_QTY'].iloc[i] * WK3 /
        #                                                  merge_all['MOQ'].iloc[i], 0) * merge_all['MOQ'].iloc[i]
        #
        #         # 第四周
        #     merge_all['W4'].iloc[i] = (merge_all['Rep_QTY'].iloc[i] - merge_all['W1'].iloc[i] -
        #                                merge_all['W2'].iloc[i] - merge_all['W3'].iloc[i])
        #     if merge_all['W4'].iloc[i] < 0:
        #         merge_all['W4'].iloc[i] = 0
        #         merge_all['W3'].iloc[i] = merge_all['W3'].iloc[i] - merge_all['W4'].iloc[i]

        # 输出结果
        rep_result = merge_all[['Material', 'FCST_Demand1', '置信度', 'Rep_QTY', 'W1', 'W2', 'W3', 'W4',
                                self.get_jnj_month(3)[0], self.get_jnj_month(3)[1], self.get_jnj_month(3)[2],
                                'Backorder_QTY', 'TotalINV_QTY', 'Onhand_QTY', 'Putaway_QTY', 'Intransit_QTY',
                                'Safetystock_QTY']]
        rep_result.rename(columns={'Material': '规格型号', 'FCST_Demand1': '二级需求',
                                   'Backorder_QTY': '缺货量', 'TotalINV_QTY': '总库存',
                                   'Onhand_QTY': '可发量', 'Putaway_QTY': '预入库',
                                   'Intransit_QTY': '在途', 'Safetystock_QTY': '安全库存',
                                   'Rep_QTY': '补货量'}, inplace=True)

        # 返回字典数据
        self._dictionary["merge_all"] = merge_all
        self._dictionary["rep_result"] = rep_result

        return self._dictionary

    # 计算修改补货计划
    def get_modify_replishment(self) -> dict:
        # 将数据合并到一起
        df_merge_tmp = self.read_replishment()
        # 读取补货计划
        next_month = self.get_next_jnj_month()
        SQL_AdjustRepPlan = "SELECT * From AdjustRepPlan WHERE JNJ_Date= %s" % next_month
        AdjustRepPlan = self.prime_db_ops.Prism_select(SQL_AdjustRepPlan)
        if AdjustRepPlan.empty:
            self._lst_missing.append("补货计划")
        # Rep_Remark存储在W1中
        AdjustRepPlan_m = AdjustRepPlan[["Material", "JNJ_Date", "week_No", "Rep_Remark"]]
        AdjustRepPlan = pd.pivot_table(AdjustRepPlan, index=["Material", "JNJ_Date"], columns=["week_No"],
                                       values=["RepWeek_QTY"])
        AdjustRepPlan_col = []
        for i in range(len(AdjustRepPlan.columns)):
            if type(AdjustRepPlan.columns.values[i][1]) == str:
                AdjustRepPlan_col.append(AdjustRepPlan.columns.values[i][1])
            else:
                AdjustRepPlan_col.append(str(AdjustRepPlan.columns.values[i][1]))
        AdjustRepPlan.columns = AdjustRepPlan_col
        AdjustRepPlan = AdjustRepPlan.reset_index()  # 重排索引

        # 合并备注数据
        AdjustRepPlan = pd.merge(AdjustRepPlan, AdjustRepPlan_m, on=["Material", "JNJ_Date"], how="left")
        # 切片，并只保留含W1的数据
        AdjustRepPlan = AdjustRepPlan[AdjustRepPlan["week_No"] == "W1"]
        AdjustRepPlan = AdjustRepPlan[["Material", "W1", "W2", "W3", "W4", "Rep_Remark"]]

        if self._lst_missing:
            print(self._lst_missing)
            # tkinter.messagebox.showwarning("警告",str(lack)+"数据缺失!")

        # 合并所需信息
        merge_all = pd.merge(df_merge_tmp, AdjustRepPlan, how="left", on="Material")
        merge_all.drop_duplicates(keep='first', inplace=True)  # 删除因merge产生的意外重复code
        # merge_all["Rep_Remark"].fillna("-", inplace=True)
        merge_all.fillna(0, inplace=True)

        # 计算安全库存量(四舍五入取整)
        merge_all['Safetystock_QTY'] = 0
        month_1 = self.get_jnj_month(3)[0]
        month_2 = self.get_jnj_month(3)[1]
        month_3 = self.get_jnj_month(3)[2]
        for i in range(len(merge_all)):
            merge_all['Safetystock_QTY'].iloc[i] = self.new_round(merge_all['Safetystock_Day'].iloc[i] *
                                                                  (merge_all[month_1].iloc[i] +
                                                                   merge_all[month_2].iloc[i] +
                                                                   merge_all[month_3].iloc[i]) / 90, 0)
        # 计算置信度
        merge_all['miu'] = np.mean(merge_all[self.get_jnj_month(6)].iloc[:], axis=1)
        merge_all['sigma'] = np.std(merge_all[self.get_jnj_month(6)].iloc[:], axis=1)

        # 历史数据为0，置信度应该为高
        # 根据出库记录，计算置信度，按可信程度区分，切片替换,|x-μ|的距离判断即可
        merge_all['sigma_level'] = abs(merge_all['FCST_Demand1'] - merge_all['miu']) / merge_all['sigma']
        merge_all['置信度'] = ""
        merge_all.loc[merge_all['sigma_level'] > 3, '置信度'] = '低'
        merge_all.loc[(merge_all['sigma_level'] > 2) & (merge_all['sigma_level'] <= 3), '置信度'] = '较低'
        merge_all.loc[(merge_all['sigma_level'] > 1) & (merge_all['sigma_level'] <= 2), '置信度'] = '较高'
        merge_all.loc[merge_all['sigma_level'] <= 1, '置信度'] = '高'

        # 计算完成，删除不需要显示的列
        merge_all = merge_all.drop(self.get_jnj_month(6)[0:2] + ['sigma_level'], axis=1)

        # 计算总库存
        merge_all['TotalINV_QTY'] = merge_all['Intransit_QTY'] + merge_all['Onhand_QTY'] + merge_all['Putaway_QTY']

        # 计算初始补货量
        merge_all['Rep_QTY'] = merge_all['W1'] + merge_all['W2'] + merge_all['W3'] + merge_all['W4']

        # 计算补货金额
        merge_all['Rep_value'] = merge_all['Rep_QTY'] * merge_all['GTS']

        # 重命名及排序
        rep_result = merge_all[['Material', 'FCST_Demand1', '置信度', 'Rep_QTY', 'W1', 'W2', 'W3', 'W4',
                                'Rep_Remark', self.get_jnj_month(3)[0], self.get_jnj_month(3)[1],
                                self.get_jnj_month(3)[2], 'Backorder_QTY', 'TotalINV_QTY', 'Onhand_QTY',
                                'Putaway_QTY', 'Intransit_QTY', 'Safetystock_QTY']]
        rep_result.rename(columns={'Material': '规格型号', 'FCST_Demand1': '二级需求',
                                   'Backorder_QTY': '缺货量', 'TotalINV_QTY': '总库存',
                                   'Onhand_QTY': '可发量', 'Putaway_QTY': '预入库',
                                   'Intransit_QTY': '在途', 'Safetystock_QTY': '安全库存',
                                   'Rep_QTY': '补货量', 'Rep_Remark': '备注'}, inplace=True)

        # 返回至字典
        self._dictionary["merge_all"] = merge_all
        self._dictionary["rep_result"] = rep_result

        return self._dictionary

    # 计算上个月在途金额（补货追踪模块）
    def acl_intransit(self) -> float:
        # 上个月在途
        SQL_Intransit = "SELECT * FROM Intransit WHERE JNJ_Date = '%s'" % self.get_jnj_month(1)[0]
        Intransit = self.prime_db_ops.Prism_select(SQL_Intransit)

        # 计算上个月在途金额
        SQL_ProductMaster_all = "SELECT Material,GTS,FCST_state From ProductMaster;"
        ProductMaster_all = self.prime_db_ops.Prism_select(SQL_ProductMaster_all)
        merge_Intransit = pd.merge(Intransit, ProductMaster_all, on='Material', how='left')
        merge_Intransit.fillna(0, inplace=True)
        Intransit_value = int(sum(merge_Intransit['GTS'] * merge_Intransit['Intransit_QTY']))

        return Intransit_value

    # 计算补货追踪数据，到货金额、订单金额、周到货数据、MTD等
    def acl_access(self) -> dict:
        # 导入主数据
        SQL_ProductMaster = "SELECT Material,GTS,FCST_state From ProductMaster WHERE FCST_state = 'MTS';"
        ProductMaster = self.prime_db_ops.Prism_select(SQL_ProductMaster)

        # 获取到货信息
        next_month = self.get_next_jnj_month()
        SQL_WeeklyInbound = "SELECT * From WeeklyInbound WHERE JNJ_Date = '%s'" % next_month
        WeeklyInbound = self.prime_db_ops.Prism_select(SQL_WeeklyInbound)
        # 提前对其进行聚合
        WeeklyInbound = WeeklyInbound.groupby(by=["JNJ_Date", "Material", "week_No"], as_index=False).sum()
        if WeeklyInbound.empty:
            self._lst_missing.append("实际到货")

        # 获取订货信息
        SQL_WeeklyOrder = "SELECT * From WeeklyOrder WHERE JNJ_Date = '%s'" % next_month
        WeeklyOrder = self.prime_db_ops.Prism_select(SQL_WeeklyOrder)
        if WeeklyOrder.empty:
            self._lst_missing.append("实际订货")

        # 获取指标
        SQL_OrderTarget = "SELECT * FROM OrderTarget WHERE JNJ_Date = '%s'" % next_month
        OrderTarget = self.prime_db_ops.Prism_select(SQL_OrderTarget)
        if OrderTarget.empty:
            self._lst_missing.append("指标")

        # 读取补货计划
        SQL_Rep_plan = "SELECT JNJ_Date,Material,RepWeek_QTY,week_No FROM AdjustRepPlan WHERE JNJ_Date = '%s'" % \
                       next_month
        Rep_plan = self.prime_db_ops.Prism_select(SQL_Rep_plan)
        if Rep_plan.empty:
            self._lst_missing.append("补货计划")

        # 上个月在途
        SQL_Intransit = "SELECT * FROM Intransit WHERE JNJ_Date = '%s'" % self.get_jnj_month(1)[0]
        Intransit = self.prime_db_ops.Prism_select(SQL_Intransit)
        if Intransit.empty:
            self._lst_missing.append("上月在途")

        if self._lst_missing:
            print(self._lst_missing)
            # tkinter.messagebox.showerror("警告", str(lack) + "数据缺失！")

        # Week
        week_No = self.prime_db_ops.Prism_select("SELECT week_No FROM WeeklyInbound WHERE JNJ_Date = '%s'" % next_month)

        # 计算上个月在途金额
        # SQL_ProductMaster_all = "SELECT Material,GTS,FCST_state From ProductMaster;"
        # ProductMaster_all = self.prime_db_ops.Prism_select(SQL_ProductMaster_all)
        # merge_Intransit = pd.merge(Intransit, ProductMaster_all, on='Material', how='left')
        # merge_Intransit.fillna(0, inplace=True)
        # Intransit_value = int(sum(merge_Intransit['GTS'] * merge_Intransit['Intransit_QTY']))

        # 合并数据--上个月计算的预测补货计划与本月实际到货
        df_merge_tmp1 = pd.merge(WeeklyInbound, Rep_plan, on=["Material", "JNJ_Date", "week_No"], how='outer')
        merge_all = pd.merge(df_merge_tmp1, ProductMaster, on=["Material"], how='left')
        merge_all.fillna(0, inplace=True)
        merge_all = merge_all[merge_all["FCST_state"] != 0]  # 只显示MTS

        # 计算每周实际到货金额与计算金额（显示和计算依据）
        merge_all.loc[:, "InboundWeekly_value"] = merge_all["GTS"] * merge_all["Inboundweek_QTY"]
        merge_all.loc[:, "Rep_plan_value"] = merge_all["GTS"] * merge_all["RepWeek_QTY"]
        # 删除week_No为0的情况
        merge_all.drop(index=merge_all[merge_all["week_No"] == 0].index, inplace=True)
        # # 去重复
        # merge_all.drop_duplicates(subset=['Material','week_No'], keep='first',inplace=True)
        merge_value = merge_all.copy()
        merge_value = pd.merge(merge_value, WeeklyOrder, on=["Material", "JNJ_Date", "week_No"], how='outer')
        # merge_value.fillna(0,inplace=True)#存在空
        merge_value.loc[:, "Orderweek_value"] = merge_value["GTS"] * merge_value["Orderweek_QTY"]
        merge_value.fillna(0, inplace=True)
        merge_value = merge_value.groupby(by=['week_No'], as_index=False).sum()
        merge_value = merge_value[["week_No", "InboundWeekly_value", "Orderweek_value", "Rep_plan_value"]]
        merge_value.loc[:, "contrast_value"] = merge_value["InboundWeekly_value"] / merge_value["Rep_plan_value"]
        merge_value = merge_value.T  # 转置
        merge_value.columns = merge_value.iloc[0, :]  # 将计算的第一行作为列名
        merge_value.drop(index='week_No', inplace=True)  # 删除第一行
        # print(merge_value)
        merge_value.reset_index(inplace=True)
        # print(merge_value.columns)
        # 若无W5数据，则手动添加为0
        if len(merge_value.columns) <= 5:
            merge_value['W5'] = 0
            merge_value.columns = ['week_No', 'W1', 'W2', 'W3', 'W4', 'W5']
        else:
            merge_value.columns = ['week_No', 'W1', 'W2', 'W3', 'W4', 'W5']
        merge_value.fillna(0, inplace=True)
        # print(merge_value)
        # MTD
        MTD_value = sum(merge_value.iloc[0][1:]) / int(OrderTarget["order_target"].iloc[0])
        # 计算本月剩余金额=目标金额-本月已到货金额
        InboundWeekly_value = int(OrderTarget["order_target"].iloc[0] -
                                  sum(merge_all["GTS"] * merge_all["Inboundweek_QTY"]))

        # 计算本周到货金额、本周订货金额
        if len(week_No['week_No'].unique()) == 1:
            inbound_value = int(merge_value.iloc[0, :]["W1"])
            order_value = int(merge_value.iloc[1, :]["W1"])
        elif len(week_No['week_No'].unique()) == 2:
            inbound_value = int(merge_value.iloc[0, :]["W2"])
            order_value = int(merge_value.iloc[1, :]["W2"])
        elif len(week_No['week_No'].unique()) == 3:
            inbound_value = int(merge_value.iloc[0, :]["W3"])
            order_value = int(merge_value.iloc[1, :]["W3"])
        elif len(week_No['week_No'].unique()) == 4:
            inbound_value = int(merge_value.iloc[0, :]["W4"])
            order_value = int(merge_value.iloc[1, :]["W4"])
        elif len(week_No['week_No'].unique()) == 5:
            inbound_value = int(merge_value.iloc[0, :]["W5"])
            order_value = int(merge_value.iloc[1, :]["W5"])
        else:
            inbound_value = 0
            order_value = 0
        merge_view = merge_all.pivot_table(values=['RepWeek_QTY', 'Inboundweek_QTY'],
                                           index=['Material', 'FCST_state'], columns='week_No')
        merge_view.reset_index(inplace=True)  # 重设index，方便显示和计算数据
        merge_view = pd.merge(merge_view, Intransit[["Material", "Intransit_QTY"]], on="Material", how="outer")
        merge_view.fillna(0, inplace=True)
        del merge_view["Material"]  # merge后会产生两列不同类型的material列，删除一列即可
        # 复核列名重置为一维列名
        merge_view_col = []
        for i in range(len(merge_view.columns)):
            if type(merge_view.columns[i]) == str:
                merge_view_col.append(str(merge_view.columns[i]))
            else:
                merge_view_col.append(str(merge_view.columns[i][0]) + str(merge_view.columns[i][1]))
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

        merge_view.loc[:, "Inboundweek_QTY"] = merge_view["Inboundweek_QTYW1"] + merge_view["Inboundweek_QTYW2"] + \
                                               merge_view["Inboundweek_QTYW3"] + merge_view["Inboundweek_QTYW4"] + \
                                               merge_view["Inboundweek_QTYW5"]
        merge_view.loc[:, "RepWeek_QTY"] = merge_view["RepWeek_QTYW1"] + merge_view["RepWeek_QTYW2"] + \
                                           merge_view["RepWeek_QTYW3"] + merge_view["RepWeek_QTYW4"]
        merge_view.drop(merge_view[merge_view["Material"] == 0].index, inplace=True)
        merge_view = merge_view[["Material", "FCST_state", "Intransit_QTY", "Inboundweek_QTY",
                                 "Inboundweek_QTYW1", "Inboundweek_QTYW2", "Inboundweek_QTYW3",
                                 "Inboundweek_QTYW4", "Inboundweek_QTYW5", "RepWeek_QTY",
                                 "RepWeek_QTYW1", "RepWeek_QTYW2", "RepWeek_QTYW3", "RepWeek_QTYW4"]]

        # 重命名,"RepWeek_QTYW5":"W5"不显示
        merge_view.rename(columns={"Material": "规格型号", "Intransit_QTY": "上月在途",
                                   "Inboundweek_QTY": "月到货量", "Inboundweek_QTYW1": "W1 ",
                                   "Inboundweek_QTYW2": "W2 ", "Inboundweek_QTYW3": "W3 ",
                                   "Inboundweek_QTYW4": "W4 ", "Inboundweek_QTYW5": "W5 ",
                                   "RepWeek_QTY": "月补货量", "RepWeek_QTYW1": "W1",
                                   "RepWeek_QTYW2": "W2", "RepWeek_QTYW3": "W3",
                                   "RepWeek_QTYW4": "W4", "FCST_state": "预测状态"}, inplace=True)
        # 输出字典型数据
        self._dictionary["inbound_value"] = inbound_value
        self._dictionary["order_value"] = order_value
        self._dictionary["merge_value"] = merge_value
        self._dictionary["merge_all"] = merge_all
        self._dictionary["merge_view"] = merge_view
        self._dictionary["MTD_value"] = MTD_value
        self._dictionary["InboundWeekly_value"] = InboundWeekly_value

        return self._dictionary

    # 初始值,直接复制AdjustRepPlan表格输入到AdjustRollingRepPlan表格
    def default_rolling(self) -> pd.DataFrame:
        SQL_adj_rolling = "SELECT JNJ_Date FROM AdjustRollingRepPlan GROUP BY JNJ_Date"
        next_month = self.get_next_jnj_month()
        if next_month not in self.prime_db_ops.Prism_select(SQL_adj_rolling)['JNJ_Date'].tolist():
            SQL_AdjustRepPlan = "SELECT * FROM AdjustRepPlan WHERE JNJ_Date='%s'" % next_month
            AdjustRepPlan = self.prime_db_ops.Prism_select(SQL_AdjustRepPlan)
            self.prime_db_ops.Prism_insert('AdjustRollingRepPlan', AdjustRepPlan)

        return AdjustRepPlan

    # rolling补货计划
    def read_acl_rolling_rep(self) -> pd.DataFrame:
        # 读取并合并数据
        # 导入主数据
        SQL_ProductMaster = "SELECT Material,GTS,FCST_state,ABC,MOQ From ProductMaster "
        ProductMaster = self.prime_db_ops.Prism_select(SQL_ProductMaster)
        ProductMaster.rename(columns={"ABC": "Class"}, inplace=True)

        # 下个月时间
        next_month = self.get_next_jnj_month()

        # ⭐每周获取出货数据
        SQL_WeeklyOutbound = "SELECT Material,Outboundweek_QTY,week_No From WeeklyOutbound WHERE JNJ_Date = '%s'" % \
                             next_month
        WeeklyOutbound = self.prime_db_ops.Prism_select(SQL_WeeklyOutbound)
        if WeeklyOutbound.empty:
            self._lst_missing.append("出货")

        # ⭐每周下单量数据
        SQL_WeeklyOrder = "SELECT Material,Orderweek_QTY,week_No FROM WeeklyOrder WHERE JNJ_Date = '%s'" % next_month
        WeeklyOrder = self.prime_db_ops.Prism_select(SQL_WeeklyOrder)
        if WeeklyOrder.empty:
            self._lst_missing.append("每周下单数据")

        # ⭐每周取最新版的缺货数据
        SQL_Backorder = "SELECT Material,Backorderweek_QTY,week_No FROM WeeklyBackorder WHERE JNJ_Date = '%s'" % \
                        next_month
        WeeklyBackorder = self.prime_db_ops.Prism_select(SQL_Backorder)
        if WeeklyBackorder.empty:
            self._lst_missing.append("每周缺货数据")

        # 调整后的需求数据
        SQL_AdjustFCSTDemand = "SELECT Material,FCST_Demand1 From AdjustFCSTDemand WHERE JNJ_Date = '%s'" % \
                               self.get_jnj_month(1)[0]
        AdjustFCSTDemand = self.prime_db_ops.Prism_select(SQL_AdjustFCSTDemand)

        # 安全库存
        # 出库数据3个月
        jnj_month_3 = str(tuple(self.get_jnj_month(3)))
        SQL_Outbound = "SELECT * From Outbound WHERE JNJ_Date in %s" % jnj_month_3
        Outbound = self.prime_db_ops.Prism_select(SQL_Outbound)

        Outbound = Outbound.pivot_table(index='Material', columns='JNJ_Date')
        Outbound.reset_index(inplace=True)
        Outbound_QTY_col = []  # 换列名，方便直接取出相应月份的数据
        for i in range(4):
            Outbound_QTY_col.append(Outbound.columns.values[i][1])
        Outbound_QTY_col[0] = 'Material'
        Outbound.columns = Outbound_QTY_col

        # 安全库存天数
        SQL_SafetyStockDay = "SELECT [Class],[Safetystock_Day] From SafetyStockDay;"
        SafetyStockDay = self.prime_db_ops.Prism_select(SQL_SafetyStockDay)
        if SafetyStockDay.empty:
            self._lst_missing.append("安全库存天数")

        # 上月Intransit
        SQL_Intransit = "SELECT Material,Intransit_QTY From Intransit WHERE JNJ_Date = '%s'" % self.get_jnj_month(1)[0]
        Intransit = self.prime_db_ops.Prism_select(SQL_Intransit)

        # 上月Onhand_QTY
        SQL_Onhand = "SELECT Material,Onhand_QTY From Onhand WHERE JNJ_Date = '%s'" % self.get_jnj_month(1)[0]
        Onhand = self.prime_db_ops.Prism_select(SQL_Onhand)

        # 上月Putaway
        SQL_Putaway = "SELECT Material,Putaway_QTY From Putaway WHERE JNJ_Date = '%s'" % self.get_jnj_month(1)[0]
        Putaway = self.prime_db_ops.Prism_select(SQL_Putaway)

        # 读取补货计划
        SQL_Rep_plan = "SELECT Material,RepWeek_QTY,week_No FROM AdjustRollingRepPlan WHERE JNJ_Date = '%s'" % \
                       next_month
        Rep_plan = self.prime_db_ops.Prism_select(SQL_Rep_plan)
        if Rep_plan.empty:
            self._lst_missing.append("补货计划")

        # 如有数缺失则显示
        if self._lst_missing:
            print(self._lst_missing)
            # tkinter.messagebox.showerror("警告", str(lack) + "数据缺失！")

        # 合并数据--上个月的数据合并，并计算安全库存
        df_merge_tmp0 = pd.merge(AdjustFCSTDemand, Onhand, on=["Material"], how='outer')
        df_merge_tmp1 = pd.merge(df_merge_tmp0, Intransit, on=["Material"], how='outer')
        df_merge_tmp2 = pd.merge(df_merge_tmp1, Putaway, on=["Material"], how='outer')
        df_merge_tmp3 = pd.merge(df_merge_tmp2, Outbound, on=["Material"], how='outer')
        df_merge_tmp4 = pd.merge(df_merge_tmp3, ProductMaster, on=["Material"], how='outer')
        merge_last = pd.merge(df_merge_tmp4, SafetyStockDay, on=["Class"], how='outer')
        # 计算安全库存量(四舍五入取整)
        merge_last.fillna(0, inplace=True)
        merge_last.loc[:, 'Safetystock_QTY'] = 0
        month_1 = self.get_jnj_month(3)[0]
        month_2 = self.get_jnj_month(3)[1]
        month_3 = self.get_jnj_month(3)[2]
        for i in range(len(merge_last)):
            merge_last['Safetystock_QTY'].iloc[i] = self.new_round(
                merge_last['Safetystock_Day'].iloc[i] * (merge_last[month_1].iloc[i] +
                                                         merge_last[month_2].iloc[i] +
                                                         merge_last[month_3].iloc[i]) / 90, 0)

        # 每周的数据更新，并更换列名
        merge_5 = pd.merge(WeeklyOutbound, Rep_plan, on=["Material", "week_No"], how='outer')
        merge_6 = pd.merge(merge_5, WeeklyBackorder, on=["Material", "week_No"], how='outer')
        merge_7 = pd.merge(merge_6, WeeklyOrder, on=["Material", "week_No"], how='outer')
        merge_8 = merge_7.pivot_table(index='Material', columns='week_No').reset_index()

        merge_columns = []
        for i in range(len(merge_8.columns)):
            merge_columns.append(merge_8.columns[i][0] + merge_8.columns[i][1])
        merge_8.columns = merge_columns

        merge_all = pd.merge(merge_8, merge_last, on=["Material"], how='outer')
        # 只选取预测状态为MTS的code
        merge_all = merge_all[merge_all["FCST_state"] == "MTS"]
        merge_all.fillna(0, inplace=True)

        # 当Wk不同时输出不同的补货计划
        merge_all.loc[:, "Rolling_RepW1_QTY"] = merge_all["RepWeek_QTYW1"]
        merge_all.loc[:, "Rolling_RepW2_QTY"] = merge_all["RepWeek_QTYW2"]
        merge_all.loc[:, "Rolling_RepW3_QTY"] = merge_all["RepWeek_QTYW3"]
        merge_all.loc[:, "Rolling_RepW4_QTY"] = merge_all["RepWeek_QTYW4"]

        # 当前缺货数
        merge_all.loc[:, "Backorderweek_QTY"] = merge_all["Backorderweek_QTYW1"]

        # 已出货量
        merge_all.loc[:, "Outboundweek_QTY"] = merge_all["Outboundweek_QTYW1"]

        # 计算补货计划值
        merge_all.loc[:, "Rolling_Rep_QTY"] = (merge_all["Rolling_RepW1_QTY"] + merge_all["Rolling_RepW2_QTY"] +
                                               merge_all["Rolling_RepW3_QTY"] + merge_all["Rolling_RepW4_QTY"])
        merge_all.loc[:, "Rolling_Rep_value"] = merge_all["Rolling_Rep_QTY"] * merge_all["GTS"]
        merge_all.fillna(0, inplace=True)

        return merge_all

    # 计算rolling逻辑
    def rolling_logistic(self, merge_all: pd.DataFrame) -> pd.DataFrame:
        # 待确认数据梳理部分(当wk为0时，则按月进行计算补货数据)
        # 当Wk不同时输出不同的补货计划
        next_month = self.get_next_jnj_month()
        WK1, WK2, WK3, WK4 = self.get_weekly_pattern()
        wk_No = self.get_week_no()
        if wk_No == "W1":
            merge_all.loc[:, "Rolling_RepW1_QTY"] = merge_all["Orderweek_QTYW1"]
            merge_all.loc[:, "Rolling_RepW2_QTY"] = self.new_round((WK2 * merge_all["FCST_Demand1"] +
                                                                    merge_all["Backorderweek_QTYW1"] +
                                                                    merge_all["Safetystock_QTY"] -
                                                                    (merge_all["Onhand_QTY"] + merge_all[
                                                                        "Intransit_QTY"] +
                                                                     merge_all["Putaway_QTY"] + merge_all[
                                                                         'Orderweek_QTYW1']
                                                                     - merge_all['Outboundweek_QTYW1'])) /
                                                                   merge_all['MOQ'], 0) * merge_all['MOQ']
            # 小于0则替换成0
            merge_all.loc[merge_all["Rolling_RepW2_QTY"] <= 0, "Rolling_RepW2_QTY"] = 0
            merge_all.loc[:, "Rolling_RepW3_QTY"] = self.new_round(WK3 * merge_all["FCST_Demand1"] /
                                                                   merge_all['MOQ'], 0) * merge_all['MOQ']
            merge_all.loc[:, "Rolling_RepW4_QTY"] = self.new_round(WK4 * merge_all["FCST_Demand1"] /
                                                                   merge_all['MOQ'], 0) * merge_all['MOQ']
            # 当前缺货数
            merge_all.loc[:, "Backorderweek_QTY"] = merge_all["Backorderweek_QTYW1"]
            # 已出货量
            merge_all.loc[:, "Outboundweek_QTY"] = merge_all["Outboundweek_QTYW1"]

        elif wk_No == "W2":
            merge_all.loc[:, "Rolling_RepW1_QTY"] = merge_all["Orderweek_QTYW1"]
            merge_all.loc[:, "Rolling_RepW2_QTY"] = merge_all["Orderweek_QTYW2"]
            merge_all.loc[:, "Rolling_RepW3_QTY"] = self.new_round((WK3 * merge_all["FCST_Demand1"] +
                                                                    merge_all["Backorderweek_QTYW2"] +
                                                                    merge_all["Safetystock_QTY"] -
                                                                    (merge_all["Onhand_QTY"] + merge_all[
                                                                        "Intransit_QTY"] +
                                                                     merge_all["Putaway_QTY"] + merge_all[
                                                                         'Orderweek_QTYW1'] +
                                                                     merge_all['Orderweek_QTYW2'] -
                                                                     merge_all['Outboundweek_QTYW1'] -
                                                                     merge_all['Outboundweek_QTYW2'])) /
                                                                   merge_all['MOQ'], 0) * merge_all['MOQ']
            # 小于0则替换成0
            merge_all.loc[merge_all["Rolling_RepW3_QTY"] <= 0, "Rolling_RepW3_QTY"] = 0
            merge_all.loc[:, "Rolling_RepW4_QTY"] = self.new_round(WK4 * merge_all["FCST_Demand1"]
                                                                   / merge_all['MOQ'], 0) * merge_all['MOQ']
            # 当前缺货数
            merge_all.loc[:, "Backorderweek_QTY"] = merge_all["Backorderweek_QTYW2"]
            # 已出货量
            merge_all.loc[:, "Outboundweek_QTY"] = merge_all["Outboundweek_QTYW1"] + merge_all['Outboundweek_QTYW2']

        elif wk_No == "W3":
            merge_all.loc[:, "Rolling_RepW1_QTY"] = merge_all["Orderweek_QTYW1"]
            merge_all.loc[:, "Rolling_RepW2_QTY"] = merge_all["Orderweek_QTYW2"]
            merge_all.loc[:, "Rolling_RepW3_QTY"] = merge_all["Orderweek_QTYW3"]
            merge_all.loc[:, "Rolling_RepW4_QTY"] = self.new_round((WK4 * merge_all["FCST_Demand1"] +
                                                                    merge_all["Backorderweek_QTYW3"] +
                                                                    merge_all["Safetystock_QTY"] -
                                                                    (merge_all["Onhand_QTY"] +
                                                                     merge_all["Intransit_QTY"] +
                                                                     merge_all["Putaway_QTY"] +
                                                                     merge_all['Orderweek_QTYW1'] +
                                                                     merge_all['Orderweek_QTYW2'] +
                                                                     merge_all['Orderweek_QTYW3'] -
                                                                     merge_all['Outboundweek_QTYW1'] -
                                                                     merge_all['Outboundweek_QTYW2'] -
                                                                     merge_all['Outboundweek_QTYW3'])) /
                                                                   merge_all['MOQ'], 0) * merge_all['MOQ']
            # 小于0则替换成0
            merge_all.loc[merge_all["Rolling_RepW4_QTY"] <= 0, "Rolling_RepW4_QTY"] = 0
            # 当前缺货数
            merge_all.loc[:, "Backorderweek_QTY"] = merge_all["Backorderweek_QTYW3"]
            # 已出货量
            merge_all.loc[:, "Outboundweek_QTY"] = merge_all["Outboundweek_QTYW1"] + merge_all['Outboundweek_QTYW2'] + \
                                                   merge_all['Outboundweek_QTYW3']

        elif wk_No == "W4":
            merge_all.loc[:, "Rolling_RepW1_QTY"] = merge_all["Orderweek_QTYW1"]
            merge_all.loc[:, "Rolling_RepW2_QTY"] = merge_all["Orderweek_QTYW2"]
            merge_all.loc[:, "Rolling_RepW3_QTY"] = merge_all["Orderweek_QTYW3"]
            merge_all.loc[:, "Rolling_RepW4_QTY"] = merge_all["Orderweek_QTYW4"]
            # 当前缺货数
            merge_all.loc[:, "Backorderweek_QTY"] = merge_all["Backorderweek_QTYW4"]
            # 已出货量
            merge_all.loc[:, "Outboundweek_QTY"] = merge_all["Outboundweek_QTYW1"] + merge_all['Outboundweek_QTYW2'] + \
                                                   merge_all['Outboundweek_QTYW3'] + merge_all['Outboundweek_QTYW4']

        else:
            print("数据缺失")
            # tkinter.messagebox.showinfo("提示", "数据缺失")
        # 计算补货计划值
        merge_all.loc[:, "Rolling_Rep_QTY"] = (merge_all["Rolling_RepW1_QTY"] + merge_all["Rolling_RepW2_QTY"] +
                                               merge_all["Rolling_RepW3_QTY"] + merge_all["Rolling_RepW4_QTY"])
        merge_all.loc[:, "Rolling_Rep_value"] = merge_all["Rolling_Rep_QTY"] * merge_all["GTS"]
        merge_all.fillna(0, inplace=True)
        # ⭐存储，需先逆透视，再加上JNJ_Date和remark才行,取消计算功能，当存在数据时，先读取数据
        # 点击更新按钮，覆盖数据，再做保存
        RollingRep = merge_all[["Material", "Rolling_RepW1_QTY", "Rolling_RepW2_QTY",
                                "Rolling_RepW3_QTY", "Rolling_RepW4_QTY"]]
        RollingRep.rename(columns={"Rolling_RepW1_QTY": "W1", "Rolling_RepW2_QTY": "W2",
                                   "Rolling_RepW3_QTY": "W3", "Rolling_RepW4_QTY": "W4"},
                          inplace=True)
        RollingRep = RollingRep.melt(id_vars=['Material'], var_name='week_No', value_name='RepWeek_QTY')
        RollingRep.loc[:, "JNJ_Date"] = next_month
        RollingRep.loc[:, "Rep_Remark"] = ""

        return RollingRep

    # 获取周编号，以outbound为准
    def get_week_no(self) -> str:
        # 下个月时间
        next_month = self.get_next_jnj_month()
        # wk_No，以WeeklyOutbound的周数为准
        SQL_week = "SELECT week_No FROM WeeklyOutbound WHERE JNJ_Date = '%s'" % next_month
        try:
            wk_no = sorted(list(self.prime_db_ops.Prism_select(SQL_week)["week_No"].unique()))[-1]
        except:
            wk_no = "无"

        return wk_no


# Jeffrey - 调试函数，创造入口
if __name__ == '__main__':
    module_test = PrismCalculation()
    # print(str(tuple(module_test.get_jnj_month(12))))
    # print(module_test.jnj_date_exist("Outbound", "201003"))
    # print(module_test.get_outbound_record(month_qty=5).head())
    # print(module_test.forecast_generation())
    # print(module_test.mape_bias())
    # print(module_test.mape_bias()["mape_df"])
    # print(module_test.history_data())
    # print(module_test.acl_replishment())
    # print(module_test.get_modify_replishment())
    # print(module_test.acl_intransit())
    # print(module_test.acl_access())
    # print(module_test.read_acl_rolling_rep())
    # df_test = module_test.read_acl_rolling_rep()
    # print(module_test.rolling_logistic(df_test))
    # print(module_test.get_week_no())

    pass
