# coding:utf-8
import sys
import xlrd
import numpy as np
import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
df = pd.read_csv('smartq2_Python.csv',header = 0)
print(df.shape)
# 下面这句实现了宽表变长表的功能，是本脚本的核心
df1=df.set_index(['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','is_gaofan']).stack().reset_index()
df1.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','is_gaofan','jidu','cash']
print(df1.shape)
# print(df1.dtypes)
df2 = df1[df1['cash']!=0]
print(df2.shape)
# 1新增辅助列1和2，
df2['fuzhulie1'] = df2['jidu'].str[0:4].astype('int')*4+df2['jidu'].str[-1:].astype('int')
df2['fuzhulie2'] = df2['jidu'].str[0:4].astype('int')*4+df2['jidu'].str[-1:].astype('int')+1
# 2新增首消季度字段
df2['first_csm_date'] = pd.to_datetime(df2['first_csm_date'],format='%Y-%m-%d')
df2['first_csm_quarter']=df2['first_csm_date'].dt.year.fillna(0).astype('int').map(str)+'q'+df2['first_csm_date'].dt.quarter.fillna(0).astype('int').map(str)
# 3用新增辅助列1和2来关联，主要是判断老户留存和老户召回
df3 = pd.merge(df2, df2,  how='left',left_on=['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','is_gaofan','first_csm_quarter','fuzhulie1'],right_on=['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','is_gaofan','first_csm_quarter','fuzhulie2'])
print(df3.shape)


def cal_test(jidu_x, first_csm_quarter,jidu_y):
    if jidu_x == first_csm_quarter:
        return '纯新户'
    elif pd.notnull(jidu_y):
        return '老户留存'
    else:
        return '老户召回'
df3['cust_flag'] = df3.apply(lambda x: cal_test(x['jidu_x'], x['first_csm_quarter'],x['jidu_y']), axis=1)
df4 =pd.DataFrame(df3,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','is_gaofan','first_csm_quarter','cash_x','jidu_x','cust_flag'])
df4.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','is_gaofan','first_csm_quarter','cash','jidu','cust_flag']
print(df4.shape)
df5=df4.groupby(['jidu','acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_flag'])['cash'].agg(['count', 'sum'])

df5.to_csv('Result3.csv',encoding="utf_8_sig")
