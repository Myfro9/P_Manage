import pandas as pd
import os
import time
import shutil
import rd_DataBase

### 本程序根据在产计划单以及对应的BOM，生成：
##  1）每一个在产计划单的BOM库存-采购分析清单（按照BOM文件的格式）
##  2）每一个在产计划单的缺料表

def Remove_Items_BOMandStock(Prj,Parent_ERP,pd_BOM,pd_VirtualStock):
    pd_BOM.dropna(axis=0,how='all')
    for i in range(pd_BOM.shape[0]):
        ERP = pd_BOM.loc[pd_BOM.index[i], '子件编码(cpscode)']
        num = pd_BOM.loc[pd_BOM.index[i], '任务单总需求量']
        #if Prj == 'RW202107-A-3':
        #    print(ERP)
        pd_inStock = pd_VirtualStock[pd_VirtualStock.ERP.str.contains(ERP).fillna(False)]
        if not pd_inStock.empty:
            Qty_inStock = pd_inStock.iat[0, 0]
            tst = pd_VirtualStock[pd_VirtualStock.ERP.str.contains(ERP).fillna(False)]
            tst2 = pd_BOM[pd_BOM.loc[:, '子件编码(cpscode)'].str.contains(ERP).fillna(False)]
            if Qty_inStock >= num:
                pd_VirtualStock.at[tst.index, 'Qty'] = tst.iat[0, 0] - num
                pd_BOM.at[tst2.index, '需新增采购量'] = 0
                Comments1 = '/{' + Prj + '_' + Parent_ERP + '_' + str(num) + '/' + str(num) + '}'
                #print('The ERP number:', ERP, 'is enough in stock:', 'We need', num,
                      #'pieces, and there are ', tst.iat[0, 0], ' in the stock')
            else:
                pd_VirtualStock.at[tst.index, 'Qty'] = 0
                pd_BOM.at[tst2.index, '需新增采购量'] = num - tst.iat[0, 0]
                Comments1 = '/{' + Prj + '_' + Parent_ERP + '_' + str(Qty_inStock) + '/' + str(num) + '}'
                #print('There are not enough pieces for  ERP number:', ERP, 'We need', num,
                      #'pieces, but there are only', tst.iat[0, 0], ' in the stock')
            Comments = tst.iat[0, 2]
            if Comments == None:
                pd_VirtualStock.at[tst.index, 'Comments'] = Comments1
            else:
                pd_VirtualStock.at[tst.index, 'Comments'] = Comments + Comments1
        else:
            tst2 = pd_BOM[pd_BOM.loc[:, '子件编码(cpscode)'].str.contains(ERP).fillna(False)]
            Qty = pd_BOM.loc[tst2.index, '任务单总需求量']
            print('there is no EPR number:', ERP, 'in Stock record information. ', Qty, 'is needed')
            pd_BOM.at[tst2.index, '需新增采购量'] = pd_BOM.loc[tst2.index, '任务单总需求量']
            pd_BOM.at[tst2.index, '基本用量分母(tdqtyd)'] = '无ERP库存信息'
    return pd_BOM, pd_VirtualStock

def Check_ChildERP(pd_BOM,BOM_subfolder):
    # 列举下层BOM_subfolder下所有子部件的半成品列表
    pd_BOM.dropna(axis=0, how='all')
    for root, dirs, files in os.walk(BOM_subfolder):
        sub_ERPlist = pd.Series(files).apply(lambda x: x.upper().strip('.XLS'))
    pd_subList = pd_BOM[pd_BOM['子件编码(cpscode)'].isin(sub_ERPlist.fillna(False))]
    if not pd_subList.empty:
        for i in range(pd_subList.shape[0]):
            pd_BOM.at[pd_subList.index[i],'是否为半成品'] = 'Yes'
            ERP = pd_BOM.loc[pd_subList.index[i], '子件编码(cpscode)']
            FileNameStr = BOM_subfolder + ERP + '.XLS'
            Qty = pd_BOM.loc[pd_subList.index[i], '需新增采购量']
            if os.path.exists(FileNameStr):
                xls_df = pd.DataFrame(pd.read_excel(FileNameStr))
                Str1 = '任务单总需求量'
                #print(FileNameStr)
                if Str1 in xls_df.columns:
                    xls_df1 = xls_df[
                        ['母件编码 *(cpspcode)H', '母件名称 *(cinvname)H', '子件编码(cpscode)', '子件名称(cinvname)', '规格型号(cinvstd)',
                         '主计量单位(ccomunitname)', '基本用量分子(ipsquantity)', '基本用量分母(tdqtyd)','任务单总需求量']]
                    xls_df1['任务单总需求量'] = xls_df1.apply(lambda x: x['基本用量分子(ipsquantity)'] * Qty + x['任务单总需求量'] , axis=1)
                else:
                    xls_df1 = xls_df[
                        ['母件编码 *(cpspcode)H', '母件名称 *(cinvname)H', '子件编码(cpscode)', '子件名称(cinvname)', '规格型号(cinvstd)',
                         '主计量单位(ccomunitname)', '基本用量分子(ipsquantity)', '基本用量分母(tdqtyd)']]
                    xls_df1['任务单总需求量'] = None
                    assert  Qty != None , '新增采购数量为None ！'
                    xls_df1['任务单总需求量'] = xls_df1.apply(lambda x: x['基本用量分子(ipsquantity)'] * Qty , axis=1)
                xls_df1.to_excel(FileNameStr.lower())
    return pd_BOM


def clear_ChildERP(BOM_subfolder):
    for root, dirs, files in os.walk(BOM_subfolder):
        for file in files:
            if os.path.splitext(file)[1].lower() == '.xls':
                xls_df = pd.DataFrame(pd.read_excel(root+file))
                Str1 = '任务单总需求量'
                if Str1 in xls_df.columns:
                    if xls_df['任务单总需求量'].sum() == 0:
                        os.remove(root+file)
                        print(root+file + 'is enough in stock and removed')
                else:
                    os.remove(root + file)
                    print(root + file + 'is useless and removed')

def Purchase_Production_BOMgen(pd_PrjList,BOM_folderNameStr, Target_folderNameStr, pd_VirtualStock):
    Task_list = pd_PrjList['任务单号']
    Task_ERP_list = pd_PrjList['物料编码']
    Task_Qty_list = pd_PrjList['任务单投产数量']
    for i in range(0,Task_ERP_list.shape[0]):
        BOM_subfolder = BOM_folderNameStr + Task_ERP_list[i] + '/'
        Result_subfolder = Target_folderNameStr + Task_list[i] + '/'
        Qty = Task_Qty_list[i]
        # 根据生产计划单号以及对应的产品ERP编码，在Result目录下建立相应的目录，并将BOM文件列表复制过去
        assert os.path.exists(BOM_subfolder), 'Can not find BOM folder ! '
        assert not os.path.exists(Result_subfolder), '目标文件夹已经存在了,再等几分钟！'
        shutil.copytree(BOM_subfolder, Result_subfolder)
        # 处理L1_L2目录下的文件
        for root, dirs, files in os.walk(Result_subfolder + 'L1/'):
            #print(root)
            for file in files:
                if os.path.splitext(file)[1].lower() == '.xls':
                    xls_df = pd.DataFrame(pd.read_excel(root+file))
                    pd_BOML1 = xls_df[
                        ['母件编码 *(cpspcode)H','母件名称 *(cinvname)H','子件编码(cpscode)','子件名称(cinvname)','规格型号(cinvstd)',
                         '主计量单位(ccomunitname)','基本用量分子(ipsquantity)','基本用量分母(tdqtyd)']]
                    pd_BOML1['任务单总需求量'] = pd_BOML1.apply(lambda x: x['基本用量分子(ipsquantity)'] * Qty, axis=1)
                    pd_BOML1['需新增采购量'] = None
                    pd_BOML1['是否为半成品'] = None
                    pd_BOML1, pd_VirtualStock = Remove_Items_BOMandStock(Task_list[i],Task_ERP_list[i],pd_BOML1,pd_VirtualStock)
                    if os.path.exists(Result_subfolder + 'L2/'):
                        pd_BOML1 = Check_ChildERP(pd_BOML1,Result_subfolder + 'L2/')
                    pd_BOML1.to_excel(Result_subfolder + 'L1/' + file.lower())
        # 处理L2_L3目录下的文件
        print(Result_subfolder)
        if os.path.exists(Result_subfolder + 'L2/'):
            clear_ChildERP(Result_subfolder + 'L2/')
            for root, dirs, files in os.walk(Result_subfolder + 'L2/'):
                for file in files:
                    if os.path.splitext(file)[1].lower() == '.xls':
                        xls_df = pd.DataFrame(pd.read_excel(root + file))
                        pd_BOML2 = xls_df
                        pd_BOML2['需新增采购量'] = None
                        pd_BOML2['是否为半成品'] = None
                        pd_BOML2, pd_VirtualStock = Remove_Items_BOMandStock(Task_list[i], file.upper().strip('.XLS'), pd_BOML2,
                                                                             pd_VirtualStock)
                        if os.path.exists(Result_subfolder + 'L3/'):
                            pd_BOML2 = Check_ChildERP(pd_BOML2, Result_subfolder + 'L3/')
                        pd_BOML2.to_excel(Result_subfolder + 'L2/' + file.lower())

        # 处理L3_L4目录下的文件
        if os.path.exists(Result_subfolder + 'L3/'):
            clear_ChildERP(Result_subfolder + 'L3/')
            for root, dirs, files in os.walk(Result_subfolder + 'L3/'):
                for file in files:
                    if os.path.splitext(file)[1].lower() == '.xls':
                        xls_df = pd.DataFrame(pd.read_excel(root + file))
                        pd_BOML3 = xls_df
                        pd_BOML3['需新增采购量'] = None
                        pd_BOML3['是否为半成品'] = None
                        pd_BOML3, pd_VirtualStock = Remove_Items_BOMandStock(Task_list[i], file.upper().strip('.XLS'), pd_BOML3, pd_VirtualStock)
                        if os.path.exists(Result_subfolder + 'L4/'):
                            pd_BOML3 = Check_ChildERP(pd_BOML3, Result_subfolder + 'L4/')
                        pd_BOML3.to_excel(Result_subfolder + 'L3/' + file.lower())
        # 处理L4目录下的文件
        if os.path.exists(Result_subfolder + 'L4/'):
            clear_ChildERP(Result_subfolder + 'L4/')
            for root, dirs, files in os.walk(Result_subfolder + 'L4/'):
                for file in files:
                    if os.path.splitext(file)[1].lower() == '.xls':
                        xls_df = pd.DataFrame(pd.read_excel(root + file))
                        pd_BOML4 = xls_df
                        pd_BOML4['需新增采购量'] = None
                        pd_BOML4['是否为半成品'] = None
                        pd_BOML4, pd_VirtualStock = Remove_Items_BOMandStock(Task_list[i], file.upper().strip('.XLS'), pd_BOML4, pd_VirtualStock)
                        pd_BOML4.to_excel(Result_subfolder + 'L4/' + file.lower())

    pd_VirtualStock.to_excel(Target_folderNameStr + 'VirtualStock_left'+ '.xls')
    return pd_VirtualStock

def PurchaseTable_Gen(pd_PrjList,pd_VirtualStock,Target_folderNameStr):
    Task_list = pd_PrjList['任务单号']
    Task_ERP_list = pd_PrjList['物料编码']
    data = [['A','A','A',0,0]]
    pd_PurchaseTable = pd.DataFrame(columns=['生产计划单号','ERP','名称','任务单总需求量','需新增采购量'])
    pd_PurchaseTable0 = pd.DataFrame(data,columns=['生产计划单号','ERP','名称','任务单总需求量','需新增采购量'])
    for i in range(0, Task_list.shape[0]):
        Task_folderNameStr = Target_folderNameStr + Task_list[i] + '/'
        # 处理L1_L2目录下的文件
        for root, dirs, files in os.walk(Task_folderNameStr ):
            print('the root is:',root)
            for file in files:
                print('the file is:',file)
                if os.path.splitext(file)[1].lower() == '.xls':
                    xls_df = pd.DataFrame(pd.read_excel(root + '/' + file))
                    for idx in range(xls_df.shape[0]):
                        Str1 = '是否为半成品'
                        if Str1 in xls_df.columns:
                            if xls_df.loc[xls_df.index[idx],'是否为半成品'] != 'Yes':
                                pd_PurchaseTable0.iat[0,0] = Task_list[i]
                                pd_PurchaseTable0.iat[0,1] = xls_df.loc[xls_df.index[idx],'子件编码(cpscode)']
                                pd_PurchaseTable0.iat[0,2] = xls_df.loc[xls_df.index[idx], '子件名称(cinvname)']
                                pd_PurchaseTable0.iat[0,3] = xls_df.loc[xls_df.index[idx], '任务单总需求量']
                                pd_PurchaseTable0.iat[0,4] = xls_df.loc[xls_df.index[idx], '需新增采购量']
                                pd_PurchaseTable = pd.concat([pd_PurchaseTable,pd_PurchaseTable0])
                        else:
                            pd_PurchaseTable0.iat[0, 0] = Task_list[i]
                            pd_PurchaseTable0.iat[0, 1] = xls_df.loc[xls_df.index[idx], '子件编码(cpscode)']
                            pd_PurchaseTable0.iat[0, 2] = xls_df.loc[xls_df.index[idx], '子件名称(cinvname)']
                            pd_PurchaseTable0.iat[0, 3] = xls_df.loc[xls_df.index[idx], '任务单总需求量']
                            pd_PurchaseTable0.iat[0, 4] = xls_df.loc[xls_df.index[idx], '需新增采购量']
                            pd_PurchaseTable = pd.concat([pd_PurchaseTable, pd_PurchaseTable0])
                            #print('test2')
        #print('test1')
    return pd_PurchaseTable

def main():
    pd_VirtualStock,pd_PrjList = rd_DataBase.Purchase_Rawdata_analyze()
    pd_VirtualStock['Comments'] = None
    BOM_folderNameStr = './BOM/'
    TimeStr = time.strftime("%Y-%m-%d-%H_%M", time.localtime(time.time()))
    #TimeStr = '2021-06-18-19_18'  # debug
    Target_folderNameStr = './results/' + 'PurchaseBOM_' + TimeStr +'/'
    pd_VirtualStock1 = Purchase_Production_BOMgen(pd_PrjList, BOM_folderNameStr, Target_folderNameStr, pd_VirtualStock)
    pd_PurchaseTable = PurchaseTable_Gen(pd_PrjList,pd_VirtualStock1,Target_folderNameStr)
    pd_PurchaseTable1 = pd_PurchaseTable[pd_PurchaseTable['需新增采购量'] >0]
    pd_PurchaseTable1.to_excel(Target_folderNameStr + 'PurchaseTable' + TimeStr  + '.xls')

if __name__ == "__main__":
    main()