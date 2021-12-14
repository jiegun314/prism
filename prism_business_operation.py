from tkinter import Label, Frame, ttk, filedialog, Button, Entry, messagebox
import pandas as pd
import numpy as np
from prism_database_operation import PrismDatabaseOperation
import datetime
from dateutil.relativedelta import relativedelta


    
class PrismCalculation:
    def __init__(self) -> None:
        self.prime_db_ops = PrismDatabaseOperation()
        pass
    
    def new_round(self, _float, _len):
        if isinstance(_float, float):
            if str(_float)[::-1].find('.') <= _len:
                return(_float)
            if str(_float)[-1] == '5':
                return(round(float(str(_float)[:-1]+'6'), _len))
            else:
                return(round(_float, _len))
        else:
            return(round(_float, _len))
    
    def get_jnj_month(self, n):
        SQL_select_date = "SELECT JNJ_Date From Outbound"
        # 建议修改成 select JNJ_Date FROM Outbound ORDER BY JNJ_Date
        Outbound_month = list(self.prime_db_ops.Prism_select(SQL_select_date)['JNJ_Date'].unique())
        last_month = sorted(Outbound_month)[-n:]
        return last_month
    
    def forecast_generation(self):
        last_month = self.get_jnj_month(1)[0]
        Outbound = self.prime_db_ops.Prism_select("SELECT * FROM Outbound WHERE JNJ_Date='"+last_month+"';")
        Backorder = self.prime_db_ops.Prism_select("SELECT * FROM Backorder WHERE JNJ_Date='"+
                                          last_month+"';")
        B_O = pd.merge(Outbound,Backorder,how="outer",on=['Material','JNJ_Date'])
        SeasonFactor = self.prime_db_ops.Prism_select("SELECT * FROM SeasonFactor;")
        ActDemand = pd.merge(B_O,SeasonFactor,how="left",on=['JNJ_Date'])
        ActDemand['Backorder_QTY'].fillna(0,inplace=True)
        ActDemand['Outbound_QTY'].fillna(0,inplace=True)
        ActDemand['ActDemand_QTY'] = (ActDemand['Outbound_QTY']+
                                      ActDemand['Backorder_QTY'])/ActDemand['season_factor']
        ActDemand = ActDemand[['JNJ_Date','Material','ActDemand_QTY']]
        # 插入之前判断是否已有数据
        ActDemand_db = self.prime_db_ops.Prism_select("SELECT * FROM ActDemand WHERE JNJ_Date='"+
                                             last_month+"';")
        #         ActDemand.to_excel(r"ActDemand.xlsx")
        missing = []
        if ActDemand_db.empty:
            self.prime_db_ops.Prism_insert('ActDemand',ActDemand)
        else:
            ActDemand = ActDemand_db
            missing.append("模型数据")
        #             tkinter.messagebox.showinfo("提示","模型数据已存在！")

        month_12 = self.get_jnj_month(12)
        FCSTmodel = pd.DataFrame()
        for i in range(12):
            SQL_select = "SELECT * FROM ActDemand WHERE JNJ_Date='"+month_12[i]+"';"
            FCSTmodel = FCSTmodel.append(self.prime_db_ops.Prism_select(SQL_select))

        # 获取所有FCST_state为开的ProductMaster数据
        SQL_state = "SELECT Material,分类Level4 FROM ProductMaster WHERE FCST_state = 'MTS'"
        state_MTS  = self.prime_db_ops.Prism_select(SQL_state)

        # 链接数据，得到所需的计算12个月的数据
        acl_FCSTDemand = pd.merge(state_MTS,FCSTmodel,how="outer",on=['Material'])
        acl_FCSTDemand.fillna(0,inplace=True)
        #         print(acl_FCSTDemand)

        # 获取权值
        FCSTWeight = self.prime_db_ops.Prism_select("SELECT * FROM FCSTWeight")
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
        FCST_Demand['JNJ_Date'] = len(FCST_Demand)*self.get_jnj_month(1)

        # 前面除过的季节因子，现在乘回来，需获取当月的JNJ_Date并获取后三个月的季节因子
        try:
            SeasonFactor = self.prime_db_ops.Prism_select("SELECT * FROM SeasonFactor")
            months_1 = (datetime.datetime.strptime(self.get_jnj_month(1)[0],"%Y%m")+
                        relativedelta(months=1)).strftime("%Y%m")
            months_2 = (datetime.datetime.strptime(self.get_jnj_month(1)[0],"%Y%m")+
                        relativedelta(months=2)).strftime("%Y%m")
            months_3 = (datetime.datetime.strptime(self.get_jnj_month(1)[0],"%Y%m")+
                        relativedelta(months=3)).strftime("%Y%m")
            SeasonFactor_1 = SeasonFactor[SeasonFactor['JNJ_Date']==months_1
                                         ]['season_factor'].iloc[0]
            SeasonFactor_2 = SeasonFactor[SeasonFactor['JNJ_Date']==months_2
                                         ]['season_factor'].iloc[0]
            SeasonFactor_3 = SeasonFactor[SeasonFactor['JNJ_Date']==months_3
                                         ]['season_factor'].iloc[0]
        except:
            messagebox.showerror("错误","季节因子不存在，请维护季节因子数据")

        #         print(FCST_Demand[FCST_Demand['Material']=='W9932'])
        FCST_Demand['FCST_Demand1'] = FCST_Demand['FCST_Demand1']*SeasonFactor_1
        FCST_Demand['FCST_Demand2'] = FCST_Demand['FCST_Demand2']*SeasonFactor_2
        FCST_Demand['FCST_Demand3'] = FCST_Demand['FCST_Demand3']*SeasonFactor_3
        #         FCST_Demand.to_excel(r"FCST_Demand.xlsx")
        FCSTDemand = FCST_Demand[['JNJ_Date','Material','FCST_Demand1','FCST_Demand2',
                                  'FCST_Demand3']]
        FCSTDemand.fillna(0,inplace=True)
        # 全部取整
        FCSTDemand['FCST_Demand1'] = self.new_round(FCSTDemand['FCST_Demand1'],0)
        FCSTDemand['FCST_Demand2'] = self.new_round(FCSTDemand['FCST_Demand2'],0)
        FCSTDemand['FCST_Demand3'] = self.new_round(FCSTDemand['FCST_Demand3'],0)

        # 判断数据库中是否已存在，将计算出来的数值插入到数据库中
        SQL_select = "SELECT JNJ_Date FROM FCSTDemand"
        FCST_Demand_Date = self.prime_db_ops.Prism_select(SQL_select)['JNJ_Date']
        if self.get_jnj_month(1) not in FCST_Demand_Date.unique() or FCST_Demand_Date.empty:
            self.prime_db_ops.Prism_insert('FCSTDemand',FCSTDemand)
        else:
            SQL = "SELECT * FROM FCSTDemand WHERE JNJ_Date ='"+self.get_jnj_month(1)[0]+"';"
            FCSTDemand = self.prime_db_ops.Prism_select(SQL)
            missing.append("预测数据")
        #             tkinter.messagebox.showinfo("提示",JNJ_Month(1)[0]+"预测数据已存在！")

        # 断数据库中是否已存在，将FCST_Demand1存入预测需求调整表格中
        AdjustFCSTDemand = FCSTDemand[['JNJ_Date','Material','FCST_Demand1']]
        AdjustFCSTDemand["Remark"] = ""
        SQL_select = "SELECT JNJ_Date FROM AdjustFCSTDemand"
        Adjust_FCST_Demand_Date = self.prime_db_ops.Prism_select(SQL_select)['JNJ_Date']
        if self.get_jnj_month(1)  not in Adjust_FCST_Demand_Date.unique() or Adjust_FCST_Demand_Date.empty:
            self.prime_db_ops.Prism_insert('AdjustFCSTDemand',AdjustFCSTDemand)
        else:
            SQL = "SELECT * FROM AdjustFCSTDemand WHERE JNJ_Date ='"+self.get_jnj_month(1)[0]+"';"
            AdjustFCSTDemand = self.prime_db_ops.Prism_select(SQL)
            missing.append("已调整预测需求")
        #             tkinter.messagebox.showinfo("提示",JNJ_Month(1)[0]+"已调整预测需求已存在")        

        # 如果missing已存在，则提示哪些已存在
        if missing != []:
            messagebox.showinfo("提示",str(missing)+"已存在!")

        # 将预测结果显示到主窗口，规格型号、产品家族、出库记录（6个月）、置信度（6个月数据，千分位）
        SQL_select = "SELECT [Material],[分类Level4] FROM ProductMaster WHERE FCST_state = 'MTS'"
        ProductMaster = self.prime_db_ops.Prism_select(SQL_select)

        # 获取出库记录（6个月）
        Outbound_QTY = pd.DataFrame()
        for i in range(6):
            SQL_select_outbound = "SELECT * FROM Outbound WHERE JNJ_Date = '"+self.get_jnj_month(6)[i]+"';"
            Outbound_QTY = Outbound_QTY.append(self.prime_db_ops.Prism_select(SQL_select_outbound))
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
        miu = np.mean(FCST[[self.get_jnj_month(6)[0],self.get_jnj_month(6)[1],self.get_jnj_month(6)[2],
                            self.get_jnj_month(6)[3],self.get_jnj_month(6)[4],self.get_jnj_month(6)[5]]].iloc[:],axis=1)
        sigma = np.std(FCST[[self.get_jnj_month(6)[0],self.get_jnj_month(6)[1],self.get_jnj_month(6)[2],
                             self.get_jnj_month(6)[3],self.get_jnj_month(6)[4],self.get_jnj_month(6)[5]]].iloc[:],axis=1)
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
        pass
    
        return FCST