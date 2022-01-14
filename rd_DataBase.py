import pandas as pd
import os

### 使用value_counts的方法，根据col列中的项目，对input_pd进行分类排列：
def sort_Col_item(Input_pd,col):
    Item_counts = Input_pd.iloc[:, col].value_counts()
    counts_result = pd.DataFrame({'col1': Item_counts.index, 'col2': Item_counts.values})
    for i in range(0,counts_result.shape[0]):
        target = counts_result.iloc[i, 0]
        result1 = Input_pd.loc[Input_pd.iloc[:,col] == target]
        if i == 0:
            result = result1
        else:
            result = pd.concat([result, result1])
        print(i)
    return result

### 使用groupby的方法，根据col列中的项目，对input_pd进行分类排列：
def sort_Groupby(Input_pd,col):
    grouped_pd = Input_pd.groupby(col)
    i = 0
    for item, item_df in grouped_pd:
        if i == 0:
            Result_pd = item_df
        else:
            Result_pd = Result_pd.append(item_df, ignore_index=True)
        i = i + 1
        print(i)
    return Result_pd

### 根据在产项目列表pd_PrjList， 将pd_Infor中的内容（例如：）过滤出来; mode=0 时是宽松过滤，采用str.contains()来过滤，mode=1是严格过滤，采用isin()来过滤
def OnGoing_Prjinfo_Filter(pd_PrjList, pd_Infor,mode):
    Task_list = pd_PrjList['任务单号']
    if mode ==0:
        for i in range(0,Task_list.shape[0]):
            if i ==0:
                if '备注(cMemo)' in pd_Infor.columns :
                    indx=pd_Infor['备注(cMemo)'].str.contains(Task_list.iloc[i]).fillna(False)
                else:
                    indx = pd_Infor['备注(cmemo)'].str.contains(Task_list.iloc[i]).fillna(False)
                pd_Infor_inTask = pd_Infor[indx]
            else:
                if '备注(cMemo)' in pd_Infor.columns:
                    pd_Infor_inTask1 = pd_Infor[pd_Infor['备注(cMemo)'].str.contains(Task_list.iloc[i]).fillna(False)]
                else:
                    pd_Infor_inTask1 = pd_Infor[pd_Infor['备注(cmemo)'].str.contains(Task_list.iloc[i]).fillna(False)]
                pd_Infor_inTask = pd.concat([pd_Infor_inTask,pd_Infor_inTask1])
    else:
        pd_Infor_inTask = pd_Infor[pd_Infor['备注(cMemo)'].isin(Task_list)]
    return pd_Infor_inTask

def Purchase_Rawdata_analyze():
    #######  导入在产生产任务单， 并整理出对应在产任务单的采购合同
    FileNameStr_PrjList = './Purchase_Rawdata/00在产任务单/近期在产生产任务单-2022-good.xls'
    pd_PrjList = pd.DataFrame(pd.read_excel(FileNameStr_PrjList))

    ####### 导入库存信息
    # The file name of stock information
    # StockInfor_Filename = './Purchase_Rawdata/11库存记录/ERP库存表格20210521.XLS'
    StockInfor_Filename = './Purchase_Rawdata/11库存记录/ERP现存量20211008.XLS'
    pd_StockInfor = pd.DataFrame(pd.read_excel(StockInfor_Filename))

    #StockInfor_pd_byERPNUM = sort_Col_item(StockInfor_pd,2)
    #pd_StockInfor_byERPNUM = sort_Groupby(pd_StockInfor1,'存货编码')
    #pd_StockInfor = pd_StockInfor1

    #StockInfor_pd_byERPNUM.to_excel('./Purchase_Rawdata/results/ERPStock_byERPNUM_20210412.xlsx')


    ####### 导入所有的采购合同
    # FileNameStr_AllContract: The file name of all history purchase contract record
    FolderNameStr = './Purchase_Rawdata/21采购合同记录/'
    FileNameStr_contract = 'ERP20150101-20210420.xls'
    pd_Contract_Infor0 = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr_contract))
    pd_Contract_Infor0 = pd_Contract_Infor0.loc[:,['备注(cMemo)','存货编号(cInvCode)','数量(iQuantity)']]
    #pd_Contract_Infor0.columns = ['Task','ERP','Qty']

    ###### ****************************************************************************** ########
    ######                                 导入询价结果                                     #######
    ###### ****************************************************************************** ########
    Contract_FolderNameStr = './Purchase_Rawdata/23询价结果/'
    Contract_FileNameStr1 = '合并汇总表-版本27-20220110.xlsx'
    Contract_FileNameStr2 = '汇总表2-版本12-20211008.xlsx'   ## 后续不用了

    pd_Contract_Infor1 = pd.DataFrame(pd.read_excel(Contract_FolderNameStr + Contract_FileNameStr1,sheet_name='操作'))
    pd_Contract_Infor1 = pd_Contract_Infor1.loc[:,['生产计划单号','ERP编码','报价数量']]
    pd_Contract_Infor1 = pd_Contract_Infor1[pd_Contract_Infor1['报价数量']>0]
    pd_Contract_Infor1.columns = ['备注(cMemo)','存货编号(cInvCode)','数量(iQuantity)']

    '''pd_Contract_Infor2 = pd.DataFrame(pd.read_excel(Contract_FolderNameStr + Contract_FileNameStr2,sheet_name='操作'))
    pd_Contract_Infor2 = pd_Contract_Infor2.loc[:, ['生产计划单号', 'ERP编码', '报价数量']]
    pd_Contract_Infor2 = pd_Contract_Infor2[pd_Contract_Infor2['报价数量'] > 0]
    pd_Contract_Infor2.columns = ['备注(cMemo)','存货编号(cInvCode)','数量(iQuantity)']'''

    #pd_Contract_Infor = pd.concat([pd_Contract_Infor0,pd_Contract_Infor1,pd_Contract_Infor2])
    pd_Contract_Infor = pd.concat([pd_Contract_Infor0, pd_Contract_Infor1])

    #xls = pd.concat([xls1,xls2,xls3,xls4,xls5])
    #xls.to_excel('./Purchase_Rawdata/results/Contracts2016-202011.xlsx')


    pd_Contract_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_Contract_Infor,1)


    #######  导入材料出库记录，并整理出对应在产任务单的材料出库记录
    # FileNameStr_Outstock: The file name of all history Out stock record
    FolderNameStr = './Purchase_Rawdata/13材料出库记录/'
    FileNameStr_outstock = '材料出库单列表20211008.xls'
    pd_OutStock = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr_outstock))
    pd_Outstock_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_OutStock,0)


    #######  导入材料入库记录，并整理出对应在产任务单的材料入库记录
    # FileNameStr_Instock: The file name of all history In stock record
    FolderNameStr = './Purchase_Rawdata/12材料入库记录/'
    FileNameStr_instock = '采购入库单列表20211008.xls'
    pd_InStock = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr_instock))
    pd_Instock_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_InStock,0)

    #######  导入产成品入库记录


    #######  导入销售出库记录


    #######  导入委外加工出库记录
    FolderNameStr = './Purchase_Rawdata/15委外材料出库记录/'
    FileNameStr_outsourcing = '委外材料出库单列表20211008.xls'
    pd_OutSourcing = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr_outsourcing))
    pd_OutSourcing_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_OutSourcing,0)

    #######  导入委外加工入库记录
    FolderNameStr = './Purchase_Rawdata/16委外产成品入库记录/'
    FileNameStr_outsourcingBack = '委外产成品入库单20211008.xls'
    pd_OutSourcingBack = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr_outsourcingBack))
    pd_OutSourcingBack_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_OutSourcingBack,0)



    #######  计算虚拟库存 = （库存记录 I + 在产任务单的材料出库记录 II ） + （在产任务单的采购合同 III - 在产任务单的采购材料入库记录 IV ）
    ## pd_Virtual_Stock =  (pd_StockInfor + pd_Outstock_inTask) + (pd_Contract_inTask - pd_Instock_inTask)

    ##### 整理库存信息，将相同ERP编码的库存项进行合并 ---- I
    pd_StockInfor2 = pd_StockInfor.loc[:,['存货编码','现存数量']]
    pd_StockInfor2.columns = ['ERP','Qty']
    pd_StockInfor3 = pd_StockInfor2.groupby('ERP').sum()
    pd_StockInfor3['ERP']= pd_StockInfor3.index
    pd_StockInfor3.index = range(0,pd_StockInfor3.shape[0])

    ##### 整理在产任务单的出库信息，将相同ERP编码的库存项进行合并  ----- II
    pd_Outstock_inTask2 = pd_Outstock_inTask.loc[:,['材料编码(cInvCode)','数量(iQuantity)']]
    pd_Outstock_inTask2.columns = ['ERP','Qty']
    pd_Outstock_inTask3 = pd_Outstock_inTask2.groupby('ERP').sum()
    pd_Outstock_inTask3['ERP']= pd_Outstock_inTask3.index
    pd_Outstock_inTask3.index = range(0,pd_Outstock_inTask3.shape[0])

    ##### 整理在产任务单的委外加工出库信息，将相同的ERP编码的库存项进行合并  -----A
    if not pd_OutSourcing_inTask.empty:
        pd_A_inTask2 = pd_OutSourcing_inTask.loc[:,['材料编码(cInvCode)','数量(iQuantity)']]
        pd_A_inTask2.columns = ['ERP','Qty']
        pd_A_inTask3 = pd_A_inTask2.groupby('ERP').sum()
        pd_A_inTask3['ERP']= pd_A_inTask3.index
        pd_A_inTask3.index = range(0,pd_A_inTask3.shape[0])
    else:
        pd_A_inTask3 = pd.DataFrame(columns=['ERP','Qty'])


    ##### 整理在产任务单的委外返回入库信息，将相同的ERP编码的库存项进行合并  -----B
    if not pd_OutSourcingBack_inTask.empty:
        pd_B_inTask2 = pd_OutSourcingBack_inTask.loc[:,['产品编码(cInvCode)','数量(iQuantity)']]
        pd_B_inTask2.columns = ['ERP','Qty']
        pd_B_inTask3 = pd_B_inTask2.groupby('ERP').sum()
        pd_B_inTask3['ERP']= pd_B_inTask3.index
        pd_B_inTask3.index = range(0,pd_B_inTask3.shape[0])
    else:
        pd_B_inTask3 = pd.DataFrame(columns=['ERP', 'Qty'])

    ##### 整理委外加工但是还没有返回的物料  ----- (A - B)
    pd_B_inTask3['Qty'] = pd_B_inTask3.apply(lambda x: x['Qty']* -1, axis=1)
    pd_OutSoucingNotback0 = pd.concat([pd_A_inTask3,pd_B_inTask3])
    pd_OutSoucingNotback1 = pd_OutSoucingNotback0.groupby('ERP').sum()
    pd_OutSoucingNotback1['ERP']= pd_OutSoucingNotback1.index
    pd_OutSoucingNotback1.index = range(0,pd_OutSoucingNotback1.shape[0])



    #####  将在产项目的材料出库记录，做退库处理   ---- [（ I + II ）+ （A-B）]
    pd_VirtualStock0 = pd.concat([pd_StockInfor3,pd_Outstock_inTask3,pd_OutSoucingNotback1])
    pd_VirtualStock1 = pd_VirtualStock0.groupby('ERP').sum()
    pd_VirtualStock1['ERP'] = pd_VirtualStock1.index
    pd_VirtualStock1.index = range(0,pd_VirtualStock1.shape[0])

    ##### 整理在产任务单的采购合同   ------  III
    pd_Contract_inTask2 = pd_Contract_inTask.loc[:,['存货编号(cInvCode)','数量(iQuantity)']]
    pd_Contract_inTask2.columns = ['ERP','Qty']
    pd_Contract_inTask3 = pd_Contract_inTask2.groupby('ERP').sum()
    pd_Contract_inTask3['ERP'] = pd_Contract_inTask3.index
    pd_Contract_inTask3.index = range(0,pd_Contract_inTask3.shape[0])

    ##### 整理在产任务单的采购材料入库记录  ----- IV
    pd_Instock_inTask2 = pd_Instock_inTask.loc[:,['存货编码(cInvCode)','数量(iQuantity)']]
    pd_Instock_inTask2.columns = ['ERP','Qty']
    pd_Instock_inTask3 = pd_Instock_inTask2.groupby('ERP').sum()
    pd_Instock_inTask3['ERP'] = pd_Instock_inTask3.index
    pd_Instock_inTask3.index = range(0,pd_Instock_inTask3.shape[0])

    ##### 整理已经签订合同但是还没有到货的物料  ----- (III - IV)
    pd_Instock_inTask3['Qty'] = pd_Instock_inTask3.apply(lambda x: x['Qty']* -1, axis=1)
    pd_Contract_not_delivered0 = pd.concat([pd_Contract_inTask3,pd_Instock_inTask3])
    pd_Contract_not_delivered1 = pd_Contract_not_delivered0.groupby('ERP').sum()
    pd_Contract_not_delivered1['ERP']= pd_Contract_not_delivered1.index
    pd_Contract_not_delivered1.index = range(0,pd_Contract_not_delivered1.shape[0])


    #####  整理虚拟库存 ------- [( I + II ）+ （ A -B ）] + （ III - IV ）
    pd_VirtualStock = pd.concat([pd_VirtualStock1,pd_Contract_not_delivered1]).groupby('ERP').sum()
    pd_VirtualStock['ERP'] = pd_VirtualStock.index
    pd_VirtualStock.index = range(0,pd_VirtualStock.shape[0])
    return pd_VirtualStock,pd_PrjList,FileNameStr_PrjList,StockInfor_Filename,FileNameStr_contract,\
           Contract_FileNameStr1,Contract_FileNameStr2,FileNameStr_outstock,FileNameStr_instock,FileNameStr_outsourcing,FileNameStr_outsourcingBack

def main():
    pd_VirtualStock, pd_PrjList = Purchase_Rawdata_analyze()
    print('OK!')

if __name__ == "__main__":
    main()