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


#######
# FileNameStr_PrjList: The file name of the on-going project list
# FolderNameStr_BOM: The folder includes all BOM information
#  The result pd_BOMofALLPrj contains BOM list for all project together

# pd_BOMofALLPrj = BOM_of_ALLPrjGen(FileNameStr_PrjList,FolderNameStr_BOM)


#######
# The file name of stock information
StockInfor_Filename = './Purchase_Rawdata/ERPStock_20210412.XLS'
StockInfor_pd = pd.DataFrame(pd.read_excel(StockInfor_Filename))
#StockInfor_pd_byERPNUM = sort_Col_item(StockInfor_pd,2)
#StockInfor_pd_byERPNUM.to_excel('./Purchase_Rawdata/results/ERPStock_byERPNUM_20210412.xlsx')


#######
# FileNameStr_AllContract: The file name of all history purchase contract record
FolderNameStr = './Purchase_Rawdata/'
FileNameStr1 = 'ERP2016.xls'
FileNameStr2 = 'ERP2017.xls'
FileNameStr3 = 'ERP2018.xls'
FileNameStr4 = 'ERP2019.xls'
FileNameStr5 = 'ERP20200101-20201130.xls'

xls1 = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr1))
xls2 = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr2))
xls3 = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr3))
xls4 = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr4))
xls5 = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr5))

xls = pd.concat([xls1,xls2,xls3,xls4,xls5])
xls.to_excel('./Purchase_Rawdata/results/Contracts2016-202011.xlsx')
# pd_OnGoingContract = OnGoing_Prj_ContractGen(FileNameStr_PrjList, FileNameStr_AllContract)


#######
# FileNameStr_Outstock: The file name of all history Out stock record
# pd_OnGoingOutstock = OnGoing_OutStockGen(FileNameStr_PrjList, FileNameStr_Outstock)

#######
# FileNameStr_Instock: The file name of all history In stock record
# pd_OnGoingInstock = OnGoing_OutStockGen(pd_OnGoingContract, FileNameStr_Instock)


