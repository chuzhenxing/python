# coding:utf-8
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

# 1#################################################
def MakeList(x):
    T = list(x)
    return T
# 2#################################################
def loadDataSet():
     # return [[1, 3, 4], [2, 3, 5], [1, 2, 3, 5], [2, 5]]

    path = os.getcwd()+'\\1020.csv'
    f = open(path)
    df = pd.read_csv(f)
    df['name']=df['name'].apply(lambda  x:x.decode('utf-8'))
    DFGrouped = df.groupby('deviceid', as_index = False)
    DF_Agg = DFGrouped.agg({'name' : MakeList})
    data1 = list(DF_Agg['name'])
    return data1

# 3#################################################
def createC1(dataSet):                                       # return C1 frequent item set
    C1 = []                                                    # C1为大小为1的项的集合
    for transaction in dataSet:                               # 遍历数据集中的每一条交易
        for item in transaction:                              # 遍历每一条交易中的每个商品
            if not [item] in C1:
                C1.append([item])
    C1.sort()
    return map(frozenset,C1)                                 # map函数表示遍历C1中的每一个元素执行forzenset，frozenset表示“冰冻”的集合，即不可改变
# 4#################################################
# Ck表示数据集，D表示候选集合的列表，minSupport表示最小支持度
# 该函数用于从C1生成L1，L1表示满足最低支持度的元素集合
def scanD(D,CK,minSupport):
    ssCnt = {}
    for tid in D:
        for can in CK:
            if can.issubset(tid):                              # issubset：表示如果集合can中的每一元素都在tid中则返回true
                if not ssCnt.has_key(can): ssCnt[can]=1       # 统计各个集合scan出现的次数，存入ssCnt字典中，字典的key是集合，value是统计出现的次数
                else:ssCnt[can]+=1
    numItems = float(len(D))
    retList = []
    supportData = {}
    for key in ssCnt:                                         # 计算每个项集的支持度，如果满足条件则把该项集加入到retList列表中
        support = ssCnt[key]/numItems
        if support >= minSupport:
            retList.insert(0, key)
        supportData[key]= support                              # 构建支持的项集的字典

    return retList,supportData # return result list and support data is a map
# 5#####################################################
def aprioriGen(Lk,k):
    #建立频繁项集
    retList = []
    lenLk = len(Lk)
    for i in range(lenLk):
        for j in range(i+1,lenLk):
            #第k-2个项相同时，将两个集合合并
            L1 = list(Lk[i])[:k-2] ;  L2 = list(Lk[j])[:k-2]
            L1.sort();L2.sort()
            if L1 == L2:
                retList.append(Lk[i]| Lk[j])
    return retList
# 6#######################################################z
def apriori(dataSet,minSupport = 0.5):
    # 创建单元素的频繁项集列表
    C1 = createC1(dataSet)
    D = map(set, dataSet)
    L1,supportData = scanD(D, C1, minSupport)
    L = [L1]
    k = 2
    while(len(L[k-2])>0):
        Ck = aprioriGen(L[k-2], k)
        Lk,supK = scanD(D, Ck, minSupport)
        supportData.update(supK)
        L.append(Lk)
        k+=1
    return L,supportData
# 7#######################################################
#========================            关联规则生成函数                     ========================
#调用下边两个函数
#L：表示频繁项集列表，supportData：包含那些频繁项集支持数据的字典，minConf：表示最小可信度阀值
def generateRules(L, supportData, minConf=0.7):
    bigRuleList = []   #存放可信度，后面可以根据可信度排名
    for i in range(1, len(L)):
        for freqSet in L[i]:
            H1 = [frozenset([item]) for item in freqSet]
            rulesFromConseq(freqSet, H1, supportData, bigRuleList, minConf)
    return bigRuleList
# 8#######################################################
#从最初的项集中产生更多的关联规则，H为当前的候选规则集，产生下一层的候选规则集
#freqSet：频繁项集 H：可以出现在规则右部的元素列表  supportData：保存项集的支持度，brl保存生成的关联规则，minConf：最小可信度阀值
def rulesFromConseq(freqSet, H, supportData, brl, minConf):
    m = len(H[0])
    while (len(freqSet) > m): # 判断长度 > m，这时即可求H的可信度
        H = calcConf(freqSet, H, supportData, brl, minConf)
        if (len(H) > 1): # 判断求完可信度后是否还有可信度大于阈值的项用来生成下一层H
            H = aprioriGen(H, m + 1)
            m += 1
        else: # 不能继续生成下一层候选关联规则，提前退出循环
            break
#计算规则的可信度，并找到满足最小可信度的规则存放在prunedH中，作
# 9#######################################################为返回值返回
def calcConf(freqSet,H,supportData,br1,minConf):
    prunedH = []
    for conseq in H:
        conf = supportData[freqSet]/supportData[freqSet-conseq]
        if conf>= minConf:
            # print freqSet-conseq,' --> ',conseq, 'confidence:',conf
            br1.append((freqSet-conseq,conseq,conf))
            prunedH.append(conseq)
    return prunedH

# 10#######################################################
if __name__ == "__main__":
    '''
    dataSet = loadDataSet()
    print(dataSet)
    C1 = createC1(dataSet)
    print(C1)
    D = map(set, dataSet)
    print(D)
    L1,supportData = scanD(D, C1, 0.5)
    print(L1)
    print(supportData)
    '''
    minSupport = input("input minSupport: ")      # 输入最小支持度
    minConf = input("input  minConf: ")        # 输入最小置信度
    dataSet = loadDataSet()
    L,supportData = apriori(dataSet,minSupport)  #apriori(dataSet)
    # print(L[0])
    # print(L[1])
    # print(L[2])
    # print(L)
    # print(supportData)

    result1 = pd.DataFrame(supportData.items(), columns=['Data', 'support'])   # 数据框
    # print result1

    result2 = result1[result1['Data'].str.len() == 2]                         # 2列数据框 把2项集筛选出来
    # print result2
    df3 = pd.DataFrame(list(result2.Data),columns=['Data1', 'Data2'])
    # print df3
    df4 = pd.DataFrame(list(result2.support),columns=['support'])
    # print df4
    df5=pd.concat([df3,df4],axis = 1)
    # print df5
    df7=df3.Data1
    df8=df3.Data2
    df9=pd.concat([df8,df7],axis=1)
    df10=pd.concat([df9,df4],axis=1)
    df10.columns = ['Data1', 'Data2', 'support']
    # print df10
    df11 = pd.concat([df5,df10],axis=0)
    # print df11                    #二项集最终输出（里面的前2列转置后合并）

    # df11.to_csv('stockIBM1.csv', encoding='utf-8', index=True,header=True)

    # result3=result2.sort_values('support',ascending=False)                     #数据框逆序排序
    # result3.to_csv('stockIBM.csv', encoding='utf-8', index=True,header=True)

    #  关联规则开始
    guanlian1  = generateRules(L, supportData, minConf)
    # print guanlian1

    guanlian2 =pd.DataFrame(guanlian1,columns=['Data1', 'Data2', 'Conf'])

    guanlian3 = guanlian2[guanlian2['Data1'].str.len() == 1]
    guanlian4 = guanlian3[guanlian3['Data2'].str.len() == 1]
    # print guanlian4
    df33 = pd.DataFrame(list(guanlian4.Data1),columns=['Data1'])
    df34 = pd.DataFrame(list(guanlian4.Data2),columns=['Data2'])
    df35 = pd.DataFrame(list(guanlian4.Conf),columns=['Conf'])
    df36 = pd.concat([df33,df34],axis=1)
    df37 = pd.concat([df36,df35],axis=1)
    df37.columns = ['Data1', 'Data2', 'Conf']
    # print df37                            # 置信度最终输出

    # df37.to_csv('stockIBM2.csv', encoding='utf-8', index=True,header=True)
    resultz = pd.merge(df11, df37,how='inner', on=['Data1', 'Data2'])
    # df.sort_values(by=['A','B'],ascending=[0,1],inplace=True)
    resulty=resultz.sort_values(by = ['support', 'Conf'],ascending=[0,1])
    print resulty
    resulty.to_csv('stockIBM2.csv', encoding='utf-8', index=True,header=True)

    '''

select deviceid, name
from
ods.sdk_dimension_pv
where
dt = '2017-07-01'
and subtype1 = 'enter'
and name not like '%�%'
and name is not null
and name <> ''
and name <> '-'
group by deviceid,name
order by deviceid,name

'''
