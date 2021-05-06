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


def remove_subComponent(Parent_ERP,BOM_subfolder,Qty,L_upper,L_lower,pd_VirtualStock):
    stop = 1  #标识位，用来标志是库存里已经有足够的子部件，不需要在BOM结构里继续往下追踪原件
    pd_BOML1 = pd.DataFrame(pd.read_excel(BOM_subfolder + L_upper + Parent_ERP + '.XLS'))
    pd_BOML1['基本用量分子(ipsquantity)'] = pd_BOML1.apply(lambda x: x['基本用量分子(ipsquantity)'] * Qty, axis=1)
    if os.path.exists(BOM_subfolder+L_lower):
        #列举L2目录下所有子编码的半成品列表
        for root, dirs, files in os.walk(BOM_subfolder + L_lower):
            L2_ERPlist = pd.Series(files).apply(lambda x: x.upper().strip('.XLS'))
        #在L1的BOM列表中，按照数量剔除L2半成品列表中的内容
            #获取L2中所有子部件的列表（通过文件列表）
            #根据L2中的子部件列表，将L1的BOM中对应的BOM列出来
        pd_BOML2_inL1 = pd_BOML1[pd_BOML1['子件编码(cpscode)'].isin(L2_ERPlist)]
        pd_BOML2_inL1_x = pd_BOML2_inL1[['子件编码(cpscode)','基本用量分子(ipsquantity)']]
            #根据L1_BOM中用到的L2子部件以及数量，从库存以及L1层BOM中删除，如果库存信息足够，则stop=1
        if pd_BOML2_inL1_x.empty:
            print('There is no subComponents in :'+ Parent_ERP)
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
                    print('The ERP '+ L_upper +' number:', ERP, 'is enough in stock:', 'We need' ,num, 'pieces, and there are ',tst.iat[0,0],' in the stock')
                else:
                    pd_VirtualStock.at[tst.index,'Qty'] = 0
                    pd_BOML1.at[tst2.index,'基本用量分子(ipsquantity)'] = num - tst.iat[0,0]
                    print('There are not enough pieces for '+ L_upper +' ERP number:', ERP, 'We need' ,num, 'pieces, but there are only',tst.iat[0,0],' in the stock' )
                    # 当前ERP编码的半成品库存不够，需要看看这个半成品是否还有L3层的子半成品可以用，且将L2层半成品的底层物料组成加入到pd_BOML1中去
                    stop=0
            else:
                print('there is no '+ L_upper +' EPR number:',ERP,'in Stock')
                stop=0
                # 库存中没有当前ERP编码的半成品，需要看看这个半成品是否还有L3层的子半成品可以用，且将L2层半成品的底层物料组成加入到pd_BOML1中去
    else:
        pd_BOML2_inL1_x = []

    return pd_BOML1,pd_BOML2_inL1_x,pd_VirtualStock,stop  #返回根据库存情况删除子部件后的库存信息，以及L1层BOM，如果库存信息足够，则stop=1

def BOM_expand_componet(pd_BOM,pd_subBOM_list,BOM_subfolder, L_lower):
    for i in range(0,pd_subBOM_list.shape[0]):
        ERP = pd_subBOM_list.iloc[i,0]
        num = pd_BOM[pd_BOM.loc[:, '子件编码(cpscode)'].str.contains(ERP)]['基本用量分子(ipsquantity)']
        pd_BOM_subERP = pd.DataFrame(pd.read_excel(BOM_subfolder + L_lower + ERP+'.XLS'))
        pd_BOM_subERP['基本用量分子(ipsquantity)'] = pd_BOM_subERP.apply(lambda x: x['基本用量分子(ipsquantity)'] * num, axis=1)
        if num.values[0] >0:
            tst = pd_BOM[pd_BOM.loc[:, '子件编码(cpscode)'].str.contains(ERP)]
            pd_BOM.loc[tst.index,'基本用量分子(ipsquantity)'] = 0
            pd_BOM = pd.concat([pd_BOM,pd_BOM_subERP])
            pd_BOM.index = range(0,pd_BOM.shape[0])
    return pd_BOM

def BOM_expandBOM(ERP,pd_parent_BOM,pd_ERP_BOM):
    tst = pd_parent_BOM[pd_parent_BOM.loc[:, '子件编码(cpscode)'].str.contains(ERP)]
    pd_parent_BOM.loc[tst.index,'基本用量分子(ipsquantity)'] = 0
    pd_parent_BOM = pd.concat([pd_parent_BOM, pd_ERP_BOM])
    pd_parent_BOM.index = range(0, pd_parent_BOM.shape[0])
    return pd_parent_BOM

def remove_L2L3L4_andExpandBOM(Top_ERP,BOM_subfolder,Qty,pd_VirtualStock):
    pd_BOML1_left, pd_BOML2_inL1_list, pd_VirtualStock_left, stop = remove_subComponent(Top_ERP, BOM_subfolder, Qty,'L1/', 'L2/', pd_VirtualStock)
    if not stop:
        if os.path.exists(BOM_subfolder+'L2/'):
            for i in range(pd_BOML2_inL1_list.shape[0]):
                stop_L2 = 1
                ERP = pd_BOML2_inL1_list.iloc[i, 0]
                num = pd_BOML1_left[pd_BOML1_left.loc[:, '子件编码(cpscode)'].str.contains(ERP)]['基本用量分子(ipsquantity)']
                #num = pd_BOML1_left.at[tst2.index, '基本用量分子(ipsquantity)']
                if num.values[0] >0:
                    pd_BOML2_left, pd_BOML3_inL2_list, pd_VirtualStock_left, stop_L2 = remove_subComponent(ERP, BOM_subfolder, num, 'L2/', 'L3/', pd_VirtualStock_left)
                    if not stop_L2:
                        pd_BOML2_left = BOM_expand_componet(pd_BOML2_left,pd_BOML3_inL2_list,BOM_subfolder, 'L3/')
                    pd_BOML1_left = BOM_expandBOM(ERP,pd_BOML1_left,pd_BOML2_left)
    return pd_BOML1_left, pd_VirtualStock_left

   # for root, dirs, files in os.walk(BOM_subfolder+'L1/'):
    #    for file in files:
    #        pd_BOML1 = pd.DataFrame(pd.read_excel(BOM_subfolder+'L1/'+file))


def BOM_component_needed_gen(pd_PrjList,folderNameStr,pd_VirtualStock):
    Task_list = pd_PrjList['任务单号']
    Task_ERP_list = pd_PrjList['物料编码']
    Task_Qty_list = pd_PrjList['任务单投产数量']
    for i in range(0,Task_ERP_list.shape[0]):
        BOM_subfolder = folderNameStr+Task_ERP_list[i]+'/'
        Qty = Task_Qty_list[i]
        if i ==0:
            pd_BOM_needed,pd_VirtualStock= remove_L2L3L4_andExpandBOM(Task_ERP_list[i],BOM_subfolder,Qty,pd_VirtualStock)
            pd_BOM_needed.iloc[:, 2] = Task_list[i]
        else:
            pd_BOM_needed_x, pd_VirtualStock = remove_L2L3L4_andExpandBOM(Task_ERP_list[i], BOM_subfolder, Qty,pd_VirtualStock)
            pd_BOM_needed_x.iloc[:,2] = Task_list[i]
            pd_BOM_needed = pd.concat([pd_BOM_needed,pd_BOM_needed_x])
            pd_BOM_needed.index = range(0,pd_BOM_needed.shape[0])
        Str_filename = folderNameStr+Task_ERP_list[i]+'_VirtualStock_tst.xlsx'
        #pd_VirtualStock.to_excel(Str_filename)
    return pd_BOM_needed,pd_VirtualStock

def BOM_component2Buy_gen(pd_BOMcomponent_needed,pd_VirtualStock):
    BOM2Buy = pd_BOMcomponent_needed

    for i in range(0,pd_BOMcomponent_needed.shape[0]):
        target_ERP = pd_BOMcomponent_needed.iloc[i,:]['子件编码(cpscode)']
        Qty_target = pd_BOMcomponent_needed.iloc[i,:]['基本用量分子(ipsquantity)']
        Qty_instock = pd_VirtualStock[pd_VirtualStock.ERP == target_ERP].Qty
        print(i)
        if i == 773:
            print(i)
        if Qty_instock.empty:
            print('There is no Component: ' + target_ERP + ' in the stock')
            # 库存为0，
        elif Qty_instock.values < Qty_target:
            # 库存不够
            print('There is no enough component: ' + target_ERP + 'in the stock. We need ' , Qty_target , ', but only ' , Qty_instock.values , ' in the stock')
            print(pd_BOMcomponent_needed.index.values[i])
            BOM2Buy.loc[pd_BOMcomponent_needed.index.values[i],'基本用量分子(ipsquantity)'] = Qty_target - Qty_instock.values
            tst = pd_VirtualStock[pd_VirtualStock.ERP.str.contains(target_ERP)]
            pd_VirtualStock.loc[tst.index, 'Qty'] = 0
        else:
            #库存足够
            print('There is enough component: ' + target_ERP + 'in the stock. We need ' , Qty_target , ', and ' , Qty_instock.values , ' in the stock')
            tst = pd_VirtualStock[pd_VirtualStock.ERP.str.contains(target_ERP)]
            pd_VirtualStock.loc[tst.index, 'Qty'] = Qty_instock.values - Qty_target
            BOM2Buy.loc[pd_BOMcomponent_needed.index.values[i],'基本用量分子(ipsquantity)'] =0
            #BOM2Buy.index = range(0,BOM2Buy.shape[0])
        print(i)
    BOM2Buy_x = BOM2Buy[['母件编码 *(cpspcode)H','母件名称 *(cinvname)H','规格型号(cinvstd)H','子件编码(cpscode)','子件名称(cinvname)','主计量单位(ccomunitname)','基本用量分子(ipsquantity)']]
    index_x =  BOM2Buy_x[BOM2Buy_x.loc[:,'基本用量分子(ipsquantity)'] ==0].index
    BOM2Buy_x.drop(index = index_x , axis=0 ,inplace = True)
    BOM2Buy_x.index = range(0,BOM2Buy_x.shape[0])
    return BOM2Buy_x





#######
# FileNameStr_PrjList: The file name of the on-going project list
# FolderNameStr_BOM: The folder includes all BOM information
#  The result pd_BOMofALLPrj contains BOM list for all project together

# pd_BOMofALLPrj = BOM_of_ALLPrjGen(FileNameStr_PrjList,FolderNameStr_BOM)


####### 导入库存信息
# The file name of stock information
StockInfor_Filename = './Purchase_Rawdata/库存记录/ERP库存表格20210423.XLS'
pd_StockInfor = pd.DataFrame(pd.read_excel(StockInfor_Filename))
#StockInfor_pd_byERPNUM = sort_Col_item(StockInfor_pd,2)
#pd_StockInfor_byERPNUM = sort_Groupby(pd_StockInfor1,'存货编码')
#pd_StockInfor = pd_StockInfor1

#StockInfor_pd_byERPNUM.to_excel('./Purchase_Rawdata/results/ERPStock_byERPNUM_20210412.xlsx')


####### 导入所有的采购合同
# FileNameStr_AllContract: The file name of all history purchase contract record
FolderNameStr = './Purchase_Rawdata/采购合同记录/'
FileNameStr1 = 'ERP20150101-20210420.xls'
pd_Contract_Infor = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr1))
#xls = pd.concat([xls1,xls2,xls3,xls4,xls5])
#xls.to_excel('./Purchase_Rawdata/results/Contracts2016-202011.xlsx')

#######  导入在产生产任务单， 并整理出对应在产任务单的采购合同
FileNameStr_PrjList = './Purchase_Rawdata/在产任务单/近期在产生产任务单-4-19-good.xls'
pd_PrjList = pd.DataFrame(pd.read_excel(FileNameStr_PrjList))
pd_Contract_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_Contract_Infor,1)


#######  导入材料出库记录，并整理出对应在产任务单的材料出库记录
# FileNameStr_Outstock: The file name of all history Out stock record
FolderNameStr = './Purchase_Rawdata/材料出库记录/'
FileNameStr1 = '材料出库单列表20210423.xls'
pd_OutStock = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr1))
pd_Outstock_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_OutStock,0)


#######  导入材料入库记录，并整理出对应在产任务单的材料入库记录
# FileNameStr_Instock: The file name of all history In stock record
FolderNameStr = './Purchase_Rawdata/材料入库记录/'
FileNameStr1 = '采购入库单列表20210423.xls'
pd_InStock = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr1))
pd_Instock_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_InStock,0)


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

#####  将在产项目的材料出库记录，做退库处理   ---- （ I + II ）
pd_VirtualStock0 = pd.concat([pd_StockInfor3,pd_Outstock_inTask3])
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


#####  整理虚拟库存 ------- （ I + II ） + （ III - IV ）
pd_VirtualStock = pd.concat([pd_VirtualStock1,pd_Contract_not_delivered1]).groupby('ERP').sum()
pd_VirtualStock['ERP'] = pd_VirtualStock.index
pd_VirtualStock.index = range(0,pd_VirtualStock.shape[0])


BOM_folderNameStr = './BOM/'

######  根据在产项目列表，统计所有的元器件级别需求 （排除库存中的成品和办成品，将剩下的需求全部转化成元器件采购级别）
pd_BOMcomponent_needed,pd_VirtualStock2 = BOM_component_needed_gen(pd_PrjList,BOM_folderNameStr,pd_VirtualStock)


######  根据库存中的元器件存量，计算需要新采购的元器件清单
BOM_component2Buy = BOM_component2Buy_gen(pd_BOMcomponent_needed,pd_VirtualStock2)
BOM_component2Buy.to_excel(BOM_folderNameStr + 'BOM2Buy.xlsx')

print('Finished')