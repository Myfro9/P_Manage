import pandas as pd
import os
import math
import warnings
import xlrd


##### 读取库存信息（其中已经包含了 'Stock_Analyse1.py' 计算的用途统计信息）
FolderNameStr1 = './Results/'
FileNameStr0 = 'ERP库存表格20210423_用途分析_0.xlsx'
FileNameStr1 = 'ERP库存表格20210423_用途分析_1.xlsx'
FileNameStr2 = 'ERP库存表格20210423_用途分析_2.xlsx'
FileNameStr3 = 'ERP库存表格20210423_用途分析_3.xlsx'
FileNameStr4 = 'ERP库存表格20210423_用途分析_4.xlsx'
FileNameStr5 = 'ERP库存表格20210423_用途分析_5.xlsx'
FileNameStr6 = 'ERP库存表格20210423_用途分析_6.xlsx'
xls0 = pd.DataFrame(pd.read_excel(FolderNameStr1+FileNameStr0))
xls1 = pd.DataFrame(pd.read_excel(FolderNameStr1+FileNameStr1))
xls2 = pd.DataFrame(pd.read_excel(FolderNameStr1+FileNameStr2))
xls3 = pd.DataFrame(pd.read_excel(FolderNameStr1+FileNameStr3))
xls4 = pd.DataFrame(pd.read_excel(FolderNameStr1+FileNameStr4))
xls5 = pd.DataFrame(pd.read_excel(FolderNameStr1+FileNameStr5))
xls6 = pd.DataFrame(pd.read_excel(FolderNameStr1+FileNameStr6))
xls = pd.concat([xls0,xls1,xls2,xls3,xls4,xls5,xls6])
xls.index = xls.iloc[:,0]

col_list = xls.columns.tolist()
xls.drop(columns=[col_list[0]],inplace=True)
col_list1 = col_list[1:len(col_list)]
col_list1.insert(7,'单价')
col_list1.insert(8,'总金额')
xls = xls.reindex(columns=col_list1)
pd_StockInfor = xls

###### 读取采购价格信息,包含了元器件 '采购价格信息' 和 '半成品折算BOM采购价格信息'
FileNameStr7 = 'RefPrice_byERPnum_202104.xlsx'
FileNameStr8 = 'PriceInfor_L3L2_V3_0.xlsx'
## 处理 '元器件采购价格信息'
xls7 = pd.DataFrame(pd.read_excel(FolderNameStr1+FileNameStr7))
col_list7 = xls7.columns.tolist()
xls7.drop(columns=[col_list7[0]],inplace=True)
## 处理 '半成品折算BOM采购价格信息'，并重新规划表格的列，以便合并
xls8 = pd.DataFrame(pd.read_excel(FolderNameStr1+FileNameStr8))
col_list8 = xls8.columns.tolist()
xls8.drop(columns=[col_list8[0]],inplace=True)
col_list = xls7.columns.tolist()
xls8 = xls8.reindex(columns=col_list)
## 合并两个表格
xls_Refprice = pd.concat([xls7,xls8])
xls_Refprice.index = range(0,xls_Refprice.shape[0])
pd_Refprice = xls_Refprice


###### 在库存信息中加入采购价格信息

for i in range(0,pd_StockInfor.shape[0]):
    target_ERP = pd_StockInfor.loc[pd_StockInfor.index[i],'存货编码']
    Sr_Match = pd_Refprice[pd_Refprice['ERP']== target_ERP]
    if not Sr_Match.empty:
        price = Sr_Match['Ref_UnitPrice'].values[0]
        pd_StockInfor.loc[pd_StockInfor.index[i],'单价']= price
        pd_StockInfor.loc[pd_StockInfor.index[i], '总金额'] = pd_StockInfor.loc[pd_StockInfor.index[i], '现存数量'] * price
        print(i)

pd_StockInfor_sorted = pd_StockInfor.sort_values(by='总金额',ascending= False)
pd_StockInfor.to_excel('./Results/'+'ERP库存表格20210423_用途及库存价格分析.xlsx')
pd_StockInfor_sorted.to_excel('./Results/'+'ERP库存表格20210423_用途及库存价格分析_按价格排序.xlsx')
print('OK')