import pandas as pd
import os

### 本程序用于生成缺料表


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
                indx=pd_Infor['备注(cMemo)'].str.contains(Task_list.iloc[i]).fillna(False)
                pd_Infor_inTask = pd_Infor[indx]
            else:
                pd_Infor_inTask1 = pd_Infor[pd_Infor['备注(cMemo)'].str.contains(Task_list.iloc[i]).fillna(False)]
                pd_Infor_inTask = pd.concat([pd_Infor_inTask,pd_Infor_inTask1])
    else:
        pd_Infor_inTask = pd_Infor[pd_Infor['备注(cMemo)'].isin(Task_list)]
    return pd_Infor_inTask




#### 对库存信息进行搜索，如果发现库存中有上层Parent_ERP用到的半成品子部件，就将该子部件从Parent_ERP的顶层BOM，以及从库存信息中删除
### 返回根据库存情况删除子部件后的库存信息，以及L1层BOM，如果库存信息足够，则stop=1，即不需要再对子部件BOM进一步展开了
def remove_subComponent(Parent_ERP,BOM_subfolder,Qty,L_upper,L_lower,pd_VirtualStock):
    stop = 1  #标识位，用来标志是库存里已经有足够的子部件，不需要在BOM结构里继续往下追踪原件
    ### 计算上层Parent_ERP的总需求pd_BOML1，其中可能包含下层子部件
    pd_BOML1 = pd.DataFrame(pd.read_excel(BOM_subfolder + L_upper + Parent_ERP + '.XLS'))
    pd_BOML1['基本用量分子(ipsquantity)'] = pd_BOML1.apply(lambda x: x['基本用量分子(ipsquantity)'] * Qty, axis=1)
    ### 如果有下层BOM文件的目录存在
    if os.path.exists(BOM_subfolder+L_lower):
        #列举下层BOM目录下所有子部件的半成品列表
        for root, dirs, files in os.walk(BOM_subfolder + L_lower):
            L2_ERPlist = pd.Series(files).apply(lambda x: x.upper().strip('.XLS'))
        #在上层总需求中，按照数量剔除下层子部件中的内容
            #获取下层中所有子部件的列表（通过文件列表）
            #根据下层子部件列表，将上层总需求中对应的下层总需求列出来
        pd_BOML2_inL1 = pd_BOML1[pd_BOML1['子件编码(cpscode)'].isin(L2_ERPlist)]
        pd_BOML2_inL1_x = pd_BOML2_inL1[['子件编码(cpscode)','基本用量分子(ipsquantity)']]
            #根据上层总需求中用到的下层子部件以及其总数量，从库存以及上层总需求中删除，如果库存信息足够，则stop=1
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
                # 库存中没有当前ERP编码的半成品，但需要看看这个半成品是否还有L3层的子半成品可以用，且将L2层半成品的底层物料组成加入到pd_BOML1中去
    else:
        pd_BOML2_inL1_x = []

    return pd_BOML1,pd_BOML2_inL1_x,pd_VirtualStock,stop  #返回根据库存情况删除子部件后的库存信息，以及L1层BOM，如果库存信息足够，则stop=1



### 将上层总需求 pd_parent_BOM 中ERP编码对应的子部件，替换成下一层的元器件级别的总需求pd_ERP_BOM，使得上层总需求也成为元器件级别
###  1）将子部件ERP编码在上层总需求的数量设置为0
###  2）将子部件ERP编码所对应的总需求，拓展成元器件，替换并入上层总需求并返回
def BOM_expand_LowComponent_into_High(ERP,pd_parent_BOM,pd_ERP_BOM):
    tst = pd_parent_BOM[pd_parent_BOM.loc[:, '子件编码(cpscode)'].str.contains(ERP)]
    pd_parent_BOM.loc[tst.index,'基本用量分子(ipsquantity)'] = 0
    pd_parent_BOM = pd.concat([pd_parent_BOM, pd_ERP_BOM])
    pd_parent_BOM.index = range(0, pd_parent_BOM.shape[0])
    return pd_parent_BOM

### 调用了BOM_expand_LowComponent_into_High()
### 如果库存中没有足够的L3子部件，那么将L2层总需求pd_BOM中还需要的L3子部件，替换扩展成L2元器件级别的总需求，这样返回值的L2层总需求就都是元器件级别的了
### 输入 ：pd_subBOM_list是L2层子部件还需要的L3子部件的列表，pd_BOM是L2层总需求，其中还可能包含着L3子部件
### BOM_subfolder是对应了当前生产计划单Top_ERP的BOM文件存放地址，例如'./BOM/P020xxxx/',其下包含了'L1','L2','L3','L4'等子目录，L_lower 是最低级别子部件目录'L3'
def BOM_expand_L3component_into_L2(pd_BOM,pd_subBOM_list,BOM_subfolder, L_lower):
    ## 针对pd_subBOM_list中的每一个L3子部件
    for i in range(0,pd_subBOM_list.shape[0]):
        ## 获取每一个L3子部件的需求总数量
        ERP = pd_subBOM_list.iloc[i,0]
        num = pd_BOM[pd_BOM.loc[:, '子件编码(cpscode)'].str.contains(ERP)]['基本用量分子(ipsquantity)']
        ## 根据该L3子部件的BOM文件，以及需求总数量，计算出该L3子部件的所有元器件需求数量（L3层子部件，已经不再含有下层子部件了）
        pd_BOM_subERP = pd.DataFrame(pd.read_excel(BOM_subfolder + L_lower + ERP+'.XLS'))
        pd_BOM_subERP['基本用量分子(ipsquantity)'] = pd_BOM_subERP.apply(lambda x: x['基本用量分子(ipsquantity)'] * num, axis=1)
        if num.values[0] >0:
            pd_BOM = BOM_expand_LowComponent_into_High(ERP, pd_BOM, pd_BOM_subERP)
            '''tst = pd_BOM[pd_BOM.loc[:, '子件编码(cpscode)'].str.contains(ERP)]
            pd_BOM.loc[tst.index,'基本用量分子(ipsquantity)'] = 0
            pd_BOM = pd.concat([pd_BOM,pd_BOM_subERP])
            pd_BOM.index = range(0,pd_BOM.shape[0])'''
    return pd_BOM




###
### 调用了BOM_expand_LowComponent_into_High() 和 BOM_expand_L3component_into_L2()
### 对某个生产计划单所对应的Top_ERP、Qty信息，执行：
###    1）从BOM和虚拟库存中分别删除库存中所有的L2L3L4层的子部件，
###    2）对于库存中不够的L2L3L4层子部件，将这些子部件以及对应的不足数量扩展成元器件级别
###    3）返回最终需要的（包含多套设备）的元器件总需求：pd_BOML1_left，以及（删除了库存子部件后）剩下的虚拟库存pd_VirtualStock_left
### 输入参数BOM_subfolder是对应了当前生产计划单Top_ERP的BOM文件存放地址，例如'./BOM/P020xxxx/',其下包含了'L1','L2','L3','L4'等子目录；pd_VirtualStock 是虚拟库存
def remove_L2L3L4_andExpandBOM(Top_ERP,BOM_subfolder,Qty,pd_VirtualStock):
    ### 首先，从L1层中删除L2的现有库存
    ###      返回去除库存后，剩下的L2库存外总需求，以及L1总需求，但剩下的L2总需求可能还有L3子部件，L1总需求可能还还有L2子部件
    pd_BOML1_left, pd_BOML2_inL1_list, pd_VirtualStock_left, stop = remove_subComponent(Top_ERP, BOM_subfolder, Qty,'L1/', 'L2/', pd_VirtualStock)
    ### 如果L2库存不够：
    if not stop:
        if os.path.exists(BOM_subfolder+'L2/'):
            ## 从/L2目录下，针对每一个L2子部件：
            for i in range(pd_BOML2_inL1_list.shape[0]):
                stop_L2 = 1
                ## 逐个找到L2子部件的剩余总需求数量
                ERP = pd_BOML2_inL1_list.iloc[i, 0]
                num = pd_BOML1_left[pd_BOML1_left.loc[:, '子件编码(cpscode)'].str.contains(ERP)]['基本用量分子(ipsquantity)']
                if num.values[0] >0:
                    ## 如果这个L2子部件的库存不够，剩余总需求数量不为0，则先从L2的总需求中，根据库存信息，删掉所需要L3的子部件（注：因为L4一般都是元器件级别，所以子部件最深只存在于L3）
                    ##      并返回库存以外，还需要的L3总需求，以及还需要的L2的库存外总需求（但此时的L2总需求还不是元器件级别，里面还有可能含有L3子部件哦）
                    pd_BOML2_left, pd_BOML3_inL2_list, pd_VirtualStock_left, stop_L2 = remove_subComponent(ERP, BOM_subfolder, num, 'L2/', 'L3/', pd_VirtualStock_left)
                    ## 如果L3的子部件库存不够，即此时的L2总需求还不是元器件级别，里面还含有L3子部件
                    if not stop_L2:
                        ## 将该生产计划单中，所有L2对L3子部件的需求，全部拓展为元器件级别，替换并入L2的总需求
                        pd_BOML2_left = BOM_expand_L3component_into_L2(pd_BOML2_left,pd_BOML3_inL2_list,BOM_subfolder, 'L3/')
                    ## 现在L2子部件总需求已经是元器件级别了。将L2子部件的总需求，以元器件级别，替换并入L1的总需求
                    pd_BOML1_left = BOM_expand_LowComponent_into_High(ERP,pd_BOML1_left,pd_BOML2_left)

                    ## 现在L1总需求，也都是元器件级别了，没有子部件了
    return pd_BOML1_left, pd_VirtualStock_left

   # for root, dirs, files in os.walk(BOM_subfolder+'L1/'):
    #    for file in files:
    #        pd_BOML1 = pd.DataFrame(pd.read_excel(BOM_subfolder+'L1/'+file))

### 调用了remove_L2L3L4_andExpandBOM()
### 对在产项目列表pd_PrjList中的生产任务单以及数量信息，执行：
###    1）从BOM和虚拟库存中分别删除库存中所有的L2L3L4层的子部件，
###    2）对于库存中不够的L2L3L4层子部件，将这些子部件以及对应的不足数量扩展成元器件级别，
##     3）返回最终需要的（包含多套设备）的元器件总BOM：pd_BOM_needed，以及删除了库存子部件后）剩下的虚拟库存pd_VirtualStock
### 输入参数： folderNameStr是所有BOM文件的存放文件夹'./BOM/'; pd_VirtualStock 是虚拟库存
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
        #Str_filename = folderNameStr+Task_ERP_list[i]+'_VirtualStock_tst.xlsx'
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
StockInfor_Filename = './Purchase_Rawdata/11库存记录/ERP库存表格20210521.XLS'
pd_StockInfor = pd.DataFrame(pd.read_excel(StockInfor_Filename))
#StockInfor_pd_byERPNUM = sort_Col_item(StockInfor_pd,2)
#pd_StockInfor_byERPNUM = sort_Groupby(pd_StockInfor1,'存货编码')
#pd_StockInfor = pd_StockInfor1

#StockInfor_pd_byERPNUM.to_excel('./Purchase_Rawdata/results/ERPStock_byERPNUM_20210412.xlsx')


####### 导入所有的采购合同
# FileNameStr_AllContract: The file name of all history purchase contract record
FolderNameStr = './Purchase_Rawdata/21采购合同记录/'
FileNameStr1 = 'ERP20150101-20210420.xls'
pd_Contract_Infor = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr1))
#xls = pd.concat([xls1,xls2,xls3,xls4,xls5])
#xls.to_excel('./Purchase_Rawdata/results/Contracts2016-202011.xlsx')

#######  导入在产生产任务单， 并整理出对应在产任务单的采购合同
FileNameStr_PrjList = './Purchase_Rawdata/00在产任务单/近期在产生产任务单-6-16-good.xls'
pd_PrjList = pd.DataFrame(pd.read_excel(FileNameStr_PrjList))
pd_Contract_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_Contract_Infor,1)


#######  导入材料出库记录，并整理出对应在产任务单的材料出库记录
# FileNameStr_Outstock: The file name of all history Out stock record
FolderNameStr = './Purchase_Rawdata/13材料出库记录/'
FileNameStr1 = '材料出库单列表20210521.xls'
pd_OutStock = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr1))
pd_Outstock_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_OutStock,0)


#######  导入材料入库记录，并整理出对应在产任务单的材料入库记录
# FileNameStr_Instock: The file name of all history In stock record
FolderNameStr = './Purchase_Rawdata/12材料入库记录/'
FileNameStr1 = '采购入库单列表20210521.xls'
pd_InStock = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr1))
pd_Instock_inTask = OnGoing_Prjinfo_Filter(pd_PrjList, pd_InStock,0)

#######  导入产成品入库记录


#######  导入销售出库记录


#######  导入委外加工出库记录


#######  导入委外加工入库记录



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
BOM_component2Buy.to_excel('./Results/' + 'BOM2Buy.xlsx')

print('Finished')