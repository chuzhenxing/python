# coding:utf-8
import sys
import xlrd
import numpy as np
import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
df = pd.read_csv('smartq2_Python.csv',header = 0)
print(df.shape)
# 1、把宽表拆分为分季度的块儿
df2017q1 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2017q1'])
df2017q2 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2017q2'])
df2017q3 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2017q3'])
df2017q4 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2017q4'])
df2018q1 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2018q1'])
df2018q2 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2018q2'])
df2018q3 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2018q3'])
df2018q4 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2018q4'])
df2019q1 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2019q1'])
df2019q2 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2019q2'])
df2019q3 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2019q3'])
df2019q4 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2019q4'])
df2020q1 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2020q1'])
df2020q2 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2020q2'])
df2020q3 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2020q3'])
df2020q4 =pd.DataFrame(df,columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','2020q4'])
# 2、把每个块儿分别添加季度标签
df2017q1['flag']='2017q1'
df2017q2['flag']='2017q2'
df2017q3['flag']='2017q3'
df2017q4['flag']='2017q4'
df2018q1['flag']='2018q1'
df2018q2['flag']='2018q2'
df2018q3['flag']='2018q3'
df2018q4['flag']='2018q4'
df2019q1['flag']='2019q1'
df2019q2['flag']='2019q2'
df2019q3['flag']='2019q3'
df2019q4['flag']='2019q4'
df2020q1['flag']='2020q1'
df2020q2['flag']='2020q2'
df2020q3['flag']='2020q3'
df2020q4['flag']='2020q4'
# 3、把每个块儿分别修改现金字段为cash,便于merge
df2017q1.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2017q2.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2017q3.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2017q4.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2018q1.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2018q2.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2018q3.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2018q4.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2019q1.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2019q2.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2019q3.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2019q4.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2020q1.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2020q2.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2020q3.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
df2020q4.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','first_csm_date','cash','flag']
# 4、merge所有块儿，整合为一个大表df2
df2=pd.concat([df2017q1,df2017q2,df2017q3,df2017q4,df2018q1,df2018q2,df2018q3,df2018q4,df2019q1,df2019q2,df2019q3,df2019q4,df2020q1,df2020q2,df2020q3,df2020q4])
print(df2.shape)
# 5、筛掉所有的cash为0的行，保留有消费的行；df3
df3 = df2[df2['cash']!=0]
print(df3.shape)

# print(df3.groupby(['flag','acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2'])['cash'].apply(sum))
# print(df3.groupby(['flag'])['cash'].apply(sum))
# print(df3.groupby(['flag'])['cash'].agg(['sum','count']))
# df3_copy = df3.copy()
# 6.1、新增辅助列1和2，
df3['fuzhulie1'] = df3['flag'].str[0:4].astype('int')*4+df3['flag'].str[-1:].astype('int')
df3['fuzhulie2'] = df3['flag'].str[0:4].astype('int')*4+df3['flag'].str[-1:].astype('int')+1
# 6.2、新增首消季度字段
df3['first_csm_date'] = pd.to_datetime(df3['first_csm_date'],format='%Y-%m-%d')
# df3['year']=df3['first_csm_date'].dt.year.fillna(0).astype('int')
# df3['quarter']=df3['first_csm_date'].dt.quarter.fillna(0).astype('int')
df3['first_csm_quarter']=df3['first_csm_date'].dt.year.fillna(0).astype('int').map(str)+'q'+df3['first_csm_date'].dt.quarter.fillna(0).astype('int').map(str)
# print(df3.shape)
# print(df3.dtypes)
# print(pd.DataFrame(df3,columns = ['first_csm_date','first_csm_quarter']).head(5))
# df3_copy['fuzhulie'] = df3_copy['flag'].str[0:4].astype('int')*4+df3_copy['flag'].str[-1:].astype('int')+1
# def fuzhu(flag):
#     year = int(flag[0:4])
#     q = int(flag[-1])
#     return year * 4 + q
# df3_copy = df3.copy()
# df3['fuzhulie'] = df3_copy['flag'].apply(fuzhu)
# print(df3.dtypes)
# 7.1、用新增辅助列1和2来关联，主要是判断老户留存和老户召回
df4 = pd.merge(df3, df3,  how='left',left_on=['fuzhulie1','cust_id','op_unit_name'],right_on=['fuzhulie2','cust_id','op_unit_name'])
print(df4.shape)
print(df4.dtypes)
# df4[df4['flag_x']=='2018q1'].head(100000).to_csv('Result.csv',encoding="utf_8_sig")
# 方法 1
# df1['year_month'] = df1['date'].apply(lambda x : x.strftime('%Y-%m'))
# 7.2、新增客户分层映射字段：判断：纯新户、老户留存、老户召回
# 参考链接：https://www.zhihu.com/question/340831983/answer/790465786
# https://blog.csdn.net/hejp_123/article/details/107819087
def cal_test(flag_x, first_csm_quarter_x,flag_y):
    if flag_x == first_csm_quarter_x:
        return '纯新户'
    elif pd.notnull(flag_y):
        return '老户留存'
    else:
        return '老户召回'
df4['cust_flag'] = df4.apply(lambda x: cal_test(x['flag_x'], x['first_csm_quarter_x'],x['flag_y']), axis=1)

# print(df4.shape)
# print(df4.dtypes)
# df4[df4['flag_x']=='2018q1'].head(100000).to_csv('Result.csv',encoding="utf_8_sig")
df5 =pd.DataFrame(df4,columns = ['acct_type_new_x','op_unit_name','ssg_trade_name_1_x','ssg_trade_name_2_x','cust_id','cust_name_x','cash_x','flag_x','cust_flag'])
df5.columns = ['acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_id','cust_name','cash','flag','cust_flag']

df6=df5.groupby(['flag','acct_type_new','op_unit_name','ssg_trade_name_1','ssg_trade_name_2','cust_flag'])['cash'].agg(['count', 'sum'])

df6.to_csv('Result.csv',encoding="utf_8_sig")
