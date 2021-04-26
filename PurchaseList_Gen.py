import pandas as pd
import os


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


def OnGoing_Prjinfo_Filter(pd_PrjList, pd_Infor,mode):
    Task_list = pd_PrjList['任务单号']
    if mode ==0:
        for i in range(0,Task_list.shape[0]):
            if i ==0:
                indx=pd_Infor['备注(cMemo)'].str.contains(Task_list.iloc[i]).fillna(False)
                pd_Infor_inTask = pd_Infor[indx]
            else:
                pd_Infor_inTask1 = pd_Infor[pd_Infor['备注(cMemo)'].str.contains(Task_list.iloc[i]).fillna(False)]
                pd_Infor_inTask = pd.concat([pd_Infor_inTask,pd_Infor_inTask1])
    else:
        pd_Infor_inTask = pd_Infor[pd_Infor['备注(cMemo)'].isin(Task_list)]
    return pd_Infor_inTask


#def removeERP(ERP,Qty,pd_VertualStock):


def removeL2L3(BOM_subfolder,Qty,pd_VirtualStock):
    if os.path.exists(BOM_subfolder+'L2'):
        #列举L2目录下所有子编码的半成品列表
        for root, dirs, files in os.walk(BOM_subfolder + 'L2/'):
            L2_ERPlist = pd.Series(files).apply(lambda x: x.upper().strip('.XLS'))
        #在L1的BOM列表中，按照数量剔除L2半成品列表中的内容
        for root, dirs, files in os.walk(BOM_subfolder + 'L1/'):
            for file in files:
                pd_BOML1 = pd.DataFrame(pd.read_excel(BOM_subfolder + 'L1/' + file))
                pd_BOML1['基本用量分子(ipsquantity)']=pd_BOML1.apply(lambda x: x['基本用量分子(ipsquantity)'] * Qty,axis=1)
                pd_BOML2_inL1 = pd_BOML1[pd_BOML1['子件编码(cpscode)'].isin(L2_ERPlist)]
                pd_BOML2_inL1_x = pd_BOML2_inL1[['子件编码(cpscode)','基本用量分子(ipsquantity)']]
            for i in range(pd_BOML2_inL1_x.shape[0]):
                ERP = pd_BOML2_inL1_x.iloc[i,0]
                num = pd_BOML2_inL1_x.iloc[i,1]
                pd_inStock = pd_VirtualStock[pd_VirtualStock.ERP.str.contains(ERP)]
                if not pd_inStock.empty:
                    Qty_inStock = pd_inStock.iat[0,0]
                    tst = pd_VirtualStock[pd_VirtualStock.ERP.str.contains(ERP)]
                    tst2 = pd_BOML1[pd_BOML1.loc[:, '子件编码(cpscode)'].str.contains(ERP)]
                    if Qty_inStock >= num:
                        pd_VirtualStock.at[tst.index,'Qty'] =  tst.iat[0,0] - num
                        pd_BOML1.at[tst2.index,'基本用量分子(ipsquantity)'] = 0
                        print('The ERP number:', ERP, 'is enough in stock:', 'We need' ,num, 'pieces, and there are ',tst.iat[0,0],' in the stock')
                    else:
                        pd_VirtualStock.at[tst.index,'Qty'] = 0
                        pd_BOML1.at[tst2.index,'基本用量分子(ipsquantity)'] = num - tst.iat[0,0]
                        print('There are not enough pieces for ERP number:', ERP, 'We need' ,num, 'pieces, but there are only',tst.iat[0,0],' in the stock' )
                        # 当前ERP编码的半成品库存不够，需要看看这个半成品是否还有L3层的子半成品可以用，且将L2层半成品的底层物料组成加入到pd_BOML1中去
                else:
                    print('there is no EPR number:',ERP,'in Stock')
                    # 库存中没有当前ERP编码的半成品，需要看看这个半成品是否还有L3层的子半成品可以用，且将L2层半成品的底层物料组成加入到pd_BOML1中去

    return 1,1

   # for root, dirs, files in os.walk(BOM_subfolder+'L1/'):
    #    for file in files:
    #        pd_BOML1 = pd.DataFrame(pd.read_excel(BOM_subfolder+'L1/'+file))


def BOM2Buy_gen(pd_PrjList,folderNameStr,pd_VirtualStock):
    Task_list = pd_PrjList['任务单号']
    Task_ERP_list = pd_PrjList['物料编码']
    Task_Qty_list = pd_PrjList['任务单投产数量']
    for i in range(0,Task_ERP_list.shape[0]):
        BOM_subfolder = folderNameStr+Task_ERP_list[i]+'/'
        Qty = Task_Qty_list[i]
        pd_BOM2Buy,pd_VirtualStock_left = removeL2L3(BOM_subfolder,Qty,pd_VirtualStock)
        print(BOM_subfolder,':')
    return(1)

'''''        print('the root is:',root)
        for dir in dirs:
            print('one dir in  ',root,' is: ', dir)
            for file in files:
                print(file)
    return(1)'''''
#######
# FileNameStr_PrjList: The file name of the on-going project list
# FolderNameStr_BOM: The folder includes all BOM information
#  The result pd_BOMofALLPrj contains BOM list for all project together

# pd_BOMofALLPrj = BOM_of_ALLPrjGen(FileNameStr_PrjList,FolderNameStr_BOM)


#######
# The file name of stock information
StockInfor_Filename = './Purchase_Rawdata/库存记录/ERP库存表格20210423.XLS'
pd_StockInfor = pd.DataFrame(pd.read_excel(StockInfor_Filename))
#StockInfor_pd_byERPNUM = sort_Col_item(StockInfor_pd,2)
#pd_StockInfor_byERPNUM = sort_Groupby(pd_StockInfor,'存货编码')
#StockInfor_pd_byERPNUM.to_excel('./Purchase_Rawdata/results/ERPStock_byERPNUM_20210412.xlsx')


#######
# FileNameStr_AllContract: The file name of all history purchase contract record
FolderNameStr = './Purchase_Rawdata/采购合同记录/'
FileNameStr1 = 'ERP20150101-20210420.xls'
pd_Contract_Infor = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr1))
#xls = pd.concat([xls1,xls2,xls3,xls4,xls5])
#xls.to_excel('./Purchase_Rawdata/results/Contracts2016-202011.xlsx')

#######
FileNameStr_PrjList = './Purchase_Rawdata/在产任务单/近期在产生产任务单-4-19-good.xls'
pd_PrjList = pd.DataFrame(pd.read_excel(FileNameStr_PrjList))
pd_Contract_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_Contract_Infor,1)


#######
# FileNameStr_Outstock: The file name of all history Out stock record
FolderNameStr = './Purchase_Rawdata/材料出库记录/'
FileNameStr1 = '材料出库单列表20210423.xls'
pd_OutStock = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr1))
pd_Outstock_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_OutStock,0)


#######
# FileNameStr_Instock: The file name of all history In stock record
FolderNameStr = './Purchase_Rawdata/材料入库记录/'
FileNameStr1 = '采购入库单列表20210423.xls'
pd_InStock = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr1))
pd_Instock_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_InStock,0)



## pd_Virtual_Stock =  (pd_StockInfor + pd_Outstock_inTask) + (pd_Contract_inTask - pd_Instock_inTask)
pd_StockInfor2 = pd_StockInfor.loc[:,['存货编码','现存数量']]
pd_StockInfor2.columns = ['ERP','Qty']
pd_StockInfor3 = pd_StockInfor2.groupby('ERP').sum()
pd_StockInfor3['ERP']= pd_StockInfor3.index
pd_StockInfor3.index = range(0,pd_StockInfor3.shape[0])

pd_Outstock_inTask2 = pd_Outstock_inTask.loc[:,['材料编码(cInvCode)','数量(iQuantity)']]
pd_Outstock_inTask2.columns = ['ERP','Qty']
pd_Outstock_inTask3 = pd_Outstock_inTask2.groupby('ERP').sum()
pd_Outstock_inTask3['ERP']= pd_Outstock_inTask3.index
pd_Outstock_inTask3.index = range(0,pd_Outstock_inTask3.shape[0])
for i in range(0,pd_Outstock_inTask3.shape[0]):
    Qty = pd_Outstock_inTask3.iloc[i,0]
    ERP_item = pd_Outstock_inTask3.iloc[i, 1]
    pd_StockInfor3[pd_StockInfor3.ERP==ERP_item].Qty = pd_StockInfor3[pd_StockInfor3.ERP==ERP_item].Qty + Qty
pd_VirtualStock = pd_StockInfor3

BOM_folderNameStr = './BOM/'
BOM2Buy = BOM2Buy_gen(pd_PrjList,BOM_folderNameStr,pd_VirtualStock)

print('Finished')