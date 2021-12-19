from tkinter import Label, Frame, ttk, filedialog, Button, Entry, messagebox
from matplotlib.pyplot import axis
import pandas as pd
import numpy as np
from prism_database_operation import PrismDatabaseOperation
import datetime
from dateutil.relativedelta import relativedelta


class PrismCalculation:
    def __init__(self) -> None:
        self.prime_db_ops = PrismDatabaseOperation()
        self._lst_missing = []
        pass
    
    # jeffrey - 没必要
    # def new_round(self, _float, _len):
    #     if isinstance(_float, float):
    #         if str(_float)[::-1].find('.') <= _len:
    #             return(_float)
    #         if str(_float)[-1] == '5':
    #             return(round(float(str(_float)[:-1]+'6'), _len))
    #         else:
    #             return(round(_float, _len))
    #     else:
    #         return(round(_float, _len))
    
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
    
    def get_season_factor(self, input_month: str) -> float:
        sql_cmd = 'SELECT season_factor FROM SeasonFactor WHERE JNJ_Date=\"%s\"' % input_month
        df_season_factor  = self.prime_db_ops.Prism_select(sql_cmd)
        lst_season_factor = df_season_factor.values.tolist()
        if len(lst_season_factor[0]) == 0:
            return 1
        else:
            return lst_season_factor[0][0]

    def get_outbound_record(self, month_qty: int) -> pd.DataFrame:
        lst_month = self.get_jnj_month(month_qty)
        str_month_list = '(\"' + '\",\"'.join(lst_month) + '\")'
        sql_cmd = "SELECT * FROM Outbound WHERE JNJ_Date in %s" % str_month_list
        Outbound_QTY = self.prime_db_ops.Prism_select(sql_cmd)
        Outbound_QTY = Outbound_QTY.pivot_table(index='Material',columns="JNJ_Date", values='Outbound_QTY')
        Outbound_QTY = Outbound_QTY.reset_index()
        return Outbound_QTY

    # Jeffrey 独立模块拆成独立函数
    def update_actual_demand(self) -> None:
        last_month = self.get_jnj_month(1)[0]
        # Jeffrey: 在变量命名中最好能表明它的数据类型，尤其对于dataframe这种特殊数据类型
        # Jeffrey：使用标准的字符串拼接
        df_outbound = self.prime_db_ops.Prism_select("SELECT * FROM Outbound WHERE JNJ_Date=\"%s\"" % last_month)
        df_backorder = self.prime_db_ops.Prism_select("SELECT * FROM Backorder WHERE JNJ_Date=\"%s\"" % last_month)
        df_backorder_ttl = pd.merge(df_outbound,df_backorder,how="outer",on=['Material','JNJ_Date'])
        df_season_factor = self.prime_db_ops.Prism_select("SELECT * FROM SeasonFactor;")
        df_act_demand = pd.merge(df_backorder_ttl,df_season_factor,how="left",on=['JNJ_Date'])
        # Jeffrey: 尽可能缩减重复语句
        df_act_demand.fillna({'Backorder_QTY': 0, 'Outbound_QTY': 0}, inplace=True)
        # df_act_demand['Backorder_QTY'].fillna(0,inplace=True)
        # df_act_demand['Outbound_QTY'].fillna(0,inplace=True)
        df_act_demand['ActDemand_QTY'] = (df_act_demand['Outbound_QTY'] + df_act_demand['Backorder_QTY']) / df_act_demand['season_factor']
        # Jeffrey: 同名覆盖会引起歧义，尽量避免
        df_act_demand.drop(columns=['Outbound_QTY', 'Backorder_QTY', 'season_factor'], inplace=True)
        # df_act_demand = df_act_demand[['JNJ_Date','Material','ActDemand_QTY']]
        # 插入之前判断是否已有数据
        df_current_demand = self.prime_db_ops.Prism_select("SELECT * FROM ActDemand WHERE JNJ_Date=\"%s\"" % last_month)
        if df_current_demand.empty:
            self.prime_db_ops.Prism_insert('ActDemand',df_act_demand)
        else:
            df_act_demand = df_current_demand.copy()
            self._lst_missing.append("模型数据")
        pass

    # Jeffrey - 将计算forecast模块单列，将来可复用
    def generate_forecast(self, df_input: pd.DataFrame, lst_weight: list, fcst_mth: int) -> pd.DataFrame:
        # get following month list
        lst_month_name = []
        for i in range(fcst_mth):
            following_mth = datetime.datetime.strptime(df_input.columns[-1],'%Y%m')+relativedelta(months=+ (i+1))
            lst_month_name.append(following_mth.strftime('%Y%m'))
        lst_cycle_month = [3, 6, 12]
        for i in range(fcst_mth):
            for j in range(len(lst_cycle_month)):
                df_input['sum_%s' % lst_cycle_month[j]] = (df_input.iloc[:, 0-lst_cycle_month[j]:].mean(axis=1)) * lst_weight[i]
            df_input[lst_month_name[i]] = df_input.iloc[:, (0 - len(lst_cycle_month)): ].sum(axis=1) * self.get_season_factor(lst_month_name[i])
            # remove sum (-4:-2)
            df_input.drop(columns=df_input.columns[(-1-len(lst_cycle_month)):-1], axis=1, inplace=True)
        return df_input

    def forecast_generation(self):
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
        df_mst  = self.prime_db_ops.Prism_select(SQL_state)

        # 链接数据，得到所需的计算12个月的数据
        df_acl_demand = pd.merge(df_mst,df_fcst_model,how="outer",on=['Material'])
        df_acl_demand.fillna(0,inplace=True)

        # 数透，material为行，JNJ_Date为列
        df_demand_pivot = df_acl_demand.pivot_table(index='Material',columns='JNJ_Date', values='ActDemand_QTY')
        df_demand_pivot.fillna(0,inplace=True)

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
            del df_demand_pivot["0"] # 当master里有而需求中没有就会出现不必要的0列，删除即可
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
        df_fcst_demand = df_forecast_result.iloc[:,-3:].reset_index() # 将数透格式转为df
        df_fcst_demand.columns = ['Material','FCST_Demand1','FCST_Demand2','FCST_Demand3']
        # jeffrey - 直接将月份列放在第一列，以免以后调整
        df_fcst_demand.insert(0, 'JNJ_Date', len(df_fcst_demand)*self.get_jnj_month(1))

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
        df_fcst_demand.fillna(0,inplace=True)
        # 全部取整
        # df_fcst_demand['FCST_Demand1'] = self.new_round(df_fcst_demand['FCST_Demand1'],0)
        # df_fcst_demand['FCST_Demand2'] = self.new_round(df_fcst_demand['FCST_Demand2'],0)
        # df_fcst_demand['FCST_Demand3'] = self.new_round(df_fcst_demand['FCST_Demand3'],0)

        # Jeffrey - 没必要做精确取整，为确保demand，可以考虑上取整，否则所有小于0.5的都会被清零
        df_fcst_demand[['FCST_Demand1', 'FCST_Demand2', 'FCST_Demand3']] = df_fcst_demand[['FCST_Demand1', 'FCST_Demand2', 'FCST_Demand3']].applymap(np.int64)

        # 判断数据库中是否已存在，将计算出来的数值插入到数据库中
        # Jeffrey - 已经存在不覆盖？？逻辑需要调整，如果存在就不覆盖，应该先计算是不是已经有了，可以直接不计算
        SQL_select = "SELECT DISTINCT(JNJ_Date) AS jnj_month FROM FCSTDemand"
        FCST_Demand_Date = self.prime_db_ops.Prism_select(SQL_select)['jnj_month'].values.tolist()
        if self.get_jnj_month(1) not in FCST_Demand_Date:
            self.prime_db_ops.Prism_insert('FCSTDemand',df_fcst_demand)
        else:
            SQL = 'SELECT * FROM FCSTDemand WHERE JNJ_Date =\"$s\"' % self.get_jnj_month(1)[0]
            df_fcst_demand = self.prime_db_ops.Prism_select(SQL)
            self._lst_missing.append("预测数据")
            # tkinter.messagebox.showinfo("提示",JNJ_Month(1)[0]+"预测数据已存在！")

        # 断数据库中是否已存在，将FCST_Demand1存入预测需求调整表格中
        AdjustFCSTDemand = df_fcst_demand[['JNJ_Date','Material','FCST_Demand1']]
        AdjustFCSTDemand["Remark"] = ""
        SQL_select = "SELECT DISTINCT(JNJ_Date) AS jnj_month FROM AdjustFCSTDemand"
        Adjust_FCST_Demand_Date = self.prime_db_ops.Prism_select(SQL_select)['jnj_month'].values.tolist()
        if self.get_jnj_month(1)  not in Adjust_FCST_Demand_Date:
            self.prime_db_ops.Prism_insert('AdjustFCSTDemand',AdjustFCSTDemand)
        else:
            SQL = 'SELECT * FROM AdjustFCSTDemand WHERE JNJ_Date =\"$s\"' % self.get_jnj_month(1)[0]
            AdjustFCSTDemand = self.prime_db_ops.Prism_select(SQL)
            self._lst_missing.append("已调整预测需求")      

        # 如果missing已存在，则提示哪些已存在
        if self._lst_missing != []:
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
        df_merge_temp_1 = pd.merge(ProductMaster,Outbound_QTY,how='left',on="Material")
        df_merge_temp_2 = pd.merge(df_merge_temp_1,df_fcst_demand,how='left',on="Material")
        df_fcst_final = pd.merge(df_merge_temp_2,AdjustFCSTDemand,how='left',on="Material")
        # Jeffrey - 删除列尽量用标准函数
        df_fcst_final.drop(columns=['JNJ_Date_x', 'JNJ_Date_y'], inplace=True)
        df_fcst_final['miu'] = np.mean(df_fcst_final[self.get_jnj_month(6)].iloc[:],axis=1)
        df_fcst_final['sigma'] = np.std(df_fcst_final[self.get_jnj_month(6)].iloc[:],axis=1)
        df_fcst_final.rename(columns={'Material':'规格型号', 
                                      'FCST_Demand1_x':'模型预测值', 
                                      'FCST_Demand1_y':'最终预测值', 
                                      'Remark':'修改原因'},
                    inplace=True)
        # # 历史数据为0，置信度应该为高
        # # 根据出库记录，计算置信度，按可信程度区分，切片替换,|x-μ|的距离判断即可
        # Jeffrey - 用简单的逻辑计算sigma level
        df_fcst_final['sigma_level'] = abs(df_fcst_final['最终预测值']-df_fcst_final['miu']) / df_fcst_final['sigma']
        df_fcst_final['置信度'] = ""
        df_fcst_final.loc[df_fcst_final['sigma_level']>3, '置信度'] = '低'
        df_fcst_final.loc[(df_fcst_final['sigma_level']>2) & (df_fcst_final['sigma_level']<=3), '置信度'] = '较低'
        df_fcst_final.loc[(df_fcst_final['sigma_level']>1) & (df_fcst_final['sigma_level']<=2), '置信度'] = '较高'
        df_fcst_final.loc[df_fcst_final['sigma_level']<=1, '置信度'] = '高'
        # Jeffrey - 这个啥意思？？
        df_fcst_final.loc[((np.mean(df_fcst_final.iloc[:,-8:-2],axis=1))<=abs(0.1)),'置信度'] = "高"
        # df_fcst_final.loc[(abs(df_fcst_final['最终预测值']-miu)>abs(3*sigma)),'置信度'] = '低'
        # df_fcst_final.loc[(abs(df_fcst_final['最终预测值']-miu)<=abs(3*sigma)) & 
        #          (abs(df_fcst_final['最终预测值']-miu)>abs(2*sigma)),'置信度'] = '较低'
        # df_fcst_final.loc[(abs(df_fcst_final['最终预测值']-miu)<=abs(2*sigma)) & 
        #          (abs(df_fcst_final['最终预测值']-miu)>abs(sigma)),'置信度'] = '较高'
        # df_fcst_final.loc[(abs(df_fcst_final['最终预测值']-miu)<=abs(sigma)),'置信度'] = "高"
        # df_fcst_final.loc[((np.mean(df_fcst_final.iloc[:,-8:-2],axis=1))<=abs(0.1)),'置信度'] = "高"
        #         print(round(FCST['FCST_Demand1']))
        df_fcst_final.fillna(0,inplace=True)
        # Jeffrey 尽可能利用完整数组
        month_list = self.get_jnj_month(6)
        month_list.reverse()
        # 重新截取数据并排序
        df_fcst_final = df_fcst_final[["规格型号","分类Level4","模型预测值","最终预测值","修改原因","置信度"] + month_list]
        # Jeffrey - 这部分添加千分位的操作不要放在这里进行，在前段展示的时候进行转换
        # 转换格式
        # for i in df_fcst_final.columns:
        #     try:
        #         for j in range(len(df_fcst_final)):
        #             df_fcst_final[i].iloc[j] = "{:,}".format(int(float(df_fcst_final[i].iloc[j])))
        #     except:
        #         df_fcst_final[i] = df_fcst_final[i].astype(str)
        # pass
    
        return df_fcst_final

# Jeffrey - 调试函数，创造入口
if __name__ == '__main__':
    module_test = PrismCalculation()
    forecast_result = module_test.forecast_generation()
    print(forecast_result.head())
    pass